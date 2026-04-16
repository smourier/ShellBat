using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace ShellBat.Utilities;

public class AuthenticodeSignerInfo
{
#pragma warning disable IDE1006 // Naming Styles
    private const string szOID_RSA_counterSign = "1.2.840.113549.1.9.6";
    private const string szOID_RFC3161_counterSign = "1.3.6.1.4.1.311.3.3.1";
    private const string szOID_NESTED_SIGNATURE = "1.3.6.1.4.1.311.2.4.1";
    private const string szOID_RSA_signingTime = "1.2.840.113549.1.9.5";
    private const string SPC_SP_OPUS_INFO_OBJID = "1.3.6.1.4.1.311.2.1.12";

    private const string sz_CERT_STORE_PROV_PKCS7 = "PKCS7";

    private const uint SPC_URL_LINK_CHOICE = 1;
    private const uint SPC_FILE_LINK_CHOICE = 3;

    private const int TIMESTAMP_INFO = 80;
    private const int PKCS7_SIGNER_INFO = 500;
#pragma warning restore IDE1006 // Naming Styles

    private readonly List<AuthenticodeSignature> _nestedSignatures = [];

    internal unsafe AuthenticodeSignerInfo(AuthenticodeSignature signature, CMSG_SIGNER_INFO* info)
    {
        Version = info->dwVersion;
        HashAlgorithm = new Oid(info->HashAlgorithm.pszObjId.ToString()!);
        HashEncryptionAlgorithm = new Oid(info->HashEncryptionAlgorithm.pszObjId.ToString()!);

        CertificateSerialNumberBytes = new byte[info->SerialNumber.cbData];
        Marshal.Copy(info->SerialNumber.pbData, CertificateSerialNumberBytes, 0, CertificateSerialNumberBytes.Length);

        EncryptedHash = new byte[info->EncryptedHash.cbData];
        Marshal.Copy(info->EncryptedHash.pbData, EncryptedHash, 0, EncryptedHash.Length);

        AuthenticatedAttributes = ReadAttributes(info->AuthAttrs);
        UnauthenticatedAttributes = ReadAttributes(info->UnauthAttrs);

        var issuer = new byte[info->Issuer.cbData];
        Marshal.Copy(info->Issuer.pbData, issuer, 0, issuer.Length);
        Issuer = new X500DistinguishedName(issuer);

        var cert = signature.Certificates.FirstOrDefault(c => c.SerialNumber.EqualsIgnoreCase(CertificateSerialNumber));
        if (cert != null)
        {
            Certificate = cert;
        }

        foreach (var attr in GetAttributes(info))
        {
            if (attr.pszObjId.ToString() == SPC_SP_OPUS_INFO_OBJID)
            {
                using var mem = GetObject(attr, attr.pszObjId);
                if (mem == null)
                    continue;

                var opusInfo = (SPC_SP_OPUS_INFO*)mem.Pointer;
                ProgramName = opusInfo->pwszProgramName.ToString();

                if (opusInfo->pPublisherInfo != 0)
                {
                    var publisherInfo = (SPC_LINK*)opusInfo->pPublisherInfo;
                    if (publisherInfo->dwLinkChoice == SPC_URL_LINK_CHOICE || publisherInfo->dwLinkChoice == SPC_FILE_LINK_CHOICE)
                    {
                        PublisherLink = publisherInfo->Anonymous.pwszUrl.ToString();
                    }
                }

                if (opusInfo->pMoreInfo != 0)
                {
                    var moreInfo = (SPC_LINK*)opusInfo->pMoreInfo;
                    if (moreInfo->dwLinkChoice == SPC_URL_LINK_CHOICE || moreInfo->dwLinkChoice == SPC_FILE_LINK_CHOICE)
                    {
                        MoreLink = moreInfo->Anonymous.pwszUrl.ToString();
                    }
                }
                continue;
            }

            if (attr.pszObjId.ToString() == szOID_NESTED_SIGNATURE)
            {
                for (var i = 0; i < attr.cValue; i++)
                {
                    var blob = (CRYPT_INTEGER_BLOB*)(attr.rgValue + i);
                    using var msg = Authenticode.Message.GetFromBlob(blob->pbData, blob->cbData);
                    if (msg == null)
                        continue;

                    var nestedSig = msg.GetSignature(false, out _);
                    if (nestedSig != null)
                    {
                        _nestedSignatures.Add(nestedSig);
                    }
                }
                continue;
            }

            if (attr.pszObjId.ToString() == szOID_RSA_counterSign)
            {
                var blob = (CRYPT_INTEGER_BLOB*)attr.rgValue;
                using var inf = GetObject(blob->pbData, blob->cbData, new PSTR(PKCS7_SIGNER_INFO));
                if (inf == null)
                    continue;

                var signerInfo = (CMSG_SIGNER_INFO*)inf.Pointer;
                var counterSignature = new AuthenticodeSignerInfo(signature, signerInfo);

                // version 0 => RSA counter signature
                TimestampSignature = new TimestampAuthenticodeSignature(inf.ToArray())
                {
                    HashAlgorithm = counterSignature.HashAlgorithm
                };
                TimestampSignature._signerInfos.Add(counterSignature);

                cert = signature.Certificates.FirstOrDefault(c => c.SerialNumber.EqualsIgnoreCase(counterSignature.CertificateSerialNumber));
                if (cert != null)
                {
                    TimestampSignature._certificates.Add(cert);
                    counterSignature.Certificate = cert;
                }

                foreach (var tsAttr in GetAttributes(signerInfo))
                {
                    if (tsAttr.pszObjId.ToString() == szOID_RSA_signingTime)
                    {
                        var tsBlob = (CRYPT_INTEGER_BLOB*)tsAttr.rgValue;
                        using var tsTime = GetObject(tsBlob->pbData, tsBlob->cbData, tsAttr.pszObjId);
                        if (tsTime != null)
                        {
                            var signingTime = (FILETIME*)tsTime.Pointer;
                            TimestampSignature.Time = GetDateTime(*signingTime);
                        }
                        break;
                    }
                }

                continue;
            }

            if (attr.pszObjId.ToString() == szOID_RFC3161_counterSign)
            {
                var blob = (CRYPT_INTEGER_BLOB*)attr.rgValue;
                using var msg = Authenticode.Message.GetFromBlob(blob->pbData, blob->cbData);
                if (msg == null)
                    continue;

                TimestampSignature = msg.GetSignature(true, out _) as TimestampAuthenticodeSignature;
                if (TimestampSignature != null)
                {
                    using var timestamp = GetObject(TimestampSignature.Content, new PSTR(TIMESTAMP_INFO));
                    if (timestamp != null)
                    {
                        // version 1 => RFC 3161 timestamp
                        var timestampInfo = (CRYPT_TIMESTAMP_INFO*)timestamp.Pointer;
                        TimestampSignature.Version = timestampInfo->dwVersion;
                        TimestampSignature.TSAPolicyId = timestampInfo->pszTSAPolicyId.ToString();
                        TimestampSignature.Time = GetDateTime(timestampInfo->ftTime);
                        TimestampSignature.HashAlgorithm = new Oid(timestampInfo->HashAlgorithm.pszObjId.ToString()!);
                    }
                }
                continue;
            }
        }
    }

    public uint Version { get; }
    public string? ProgramName { get; }
    public string? PublisherLink { get; }
    public string? MoreLink { get; }
    public Oid HashAlgorithm { get; }
    public Oid HashEncryptionAlgorithm { get; }
    public byte[] EncryptedHash { get; }
    public byte[] CertificateSerialNumberBytes { get; }
    public X509Certificate2? Certificate { get; private set; }
    public string CertificateSerialNumber => SerialNumberToHex(CertificateSerialNumberBytes);
    public CryptographicAttributeObject[] AuthenticatedAttributes { get; }
    public CryptographicAttributeObject[] UnauthenticatedAttributes { get; }
    public X500DistinguishedName? Issuer { get; }
    public TimestampAuthenticodeSignature? TimestampSignature { get; }
    public IReadOnlyList<AuthenticodeSignature> NestedSignatures => _nestedSignatures.AsReadOnly();
    public IEnumerable<CryptographicAttributeObject> AllAttributes
    {
        get
        {
            foreach (var attr in AuthenticatedAttributes)
            {
                yield return attr;
            }

            foreach (var attr in UnauthenticatedAttributes)
            {
                yield return attr;
            }
        }
    }

    private static DateTime GetDateTime(in FILETIME ft)
    {
        // these are not equivalent with DateTime.FromFileTime, so try to use the API first
        // so we're compatible with what the shell displays
        if (ShellN.Functions.FileTimeToLocalFileTime(ft, out var localFt) && ShellN.Functions.FileTimeToSystemTime(localFt, out var st))
            return new DateTime(st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);

        return DateTime.FromFileTime(((long)ft.dwHighDateTime << 32) | ft.dwLowDateTime);
    }

    // reverse order here
    private static string SerialNumberToHex(ReadOnlySpan<byte> bytes)
    {
        const string _hexChars = "0123456789ABCDEF";

        var chars = new char[bytes.Length * 2];
        var index = 0;
        for (var i = bytes.Length - 1; i >= 0; i--)
        {
            var b = bytes[i];
            chars[index++] = _hexChars[b >> 4];
            chars[index++] = _hexChars[b & 0x0F];
        }

        return new string(chars);
    }

    private static unsafe List<CRYPT_ATTRIBUTE> GetAttributes(CMSG_SIGNER_INFO* info)
    {
        var list = new List<CRYPT_ATTRIBUTE>();
        var pAttr = (CRYPT_ATTRIBUTE*)info->AuthAttrs.rgAttr;
        for (var i = 0; i < info->AuthAttrs.cAttr; i++)
        {
            list.Add(pAttr[i]);
        }

        pAttr = (CRYPT_ATTRIBUTE*)info->UnauthAttrs.rgAttr;
        for (var i = 0; i < info->UnauthAttrs.cAttr; i++)
        {
            list.Add(pAttr[i]);
        }
        return list;
    }

    private static unsafe ComMemory? GetObject(CRYPT_ATTRIBUTE attr, PSTR type)
    {
        var blob = (CRYPT_INTEGER_BLOB*)attr.rgValue;
        return GetObject(blob->pbData, blob->cbData, type);
    }

    private static unsafe ComMemory? GetObject(byte[] bytes, PSTR type)
    {
        fixed (byte* ptr = bytes)
        {
            return GetObject((nint)ptr, (uint)bytes.Length, type);
        }
    }

    private static ComMemory? GetObject(nint pointer, uint count, PSTR type)
    {
        uint size = 0;
        ShellN.Functions.CryptDecodeObjectEx(CERT_QUERY_ENCODING_TYPE.X509_ASN_ENCODING | CERT_QUERY_ENCODING_TYPE.PKCS_7_ASN_ENCODING,
            type,
            pointer,
            count,
            0,
            0,
            0,
            ref size);
        if (size == 0)
            return null;

        var mem = new ComMemory(size);
        if (!ShellN.Functions.CryptDecodeObjectEx(CERT_QUERY_ENCODING_TYPE.X509_ASN_ENCODING | CERT_QUERY_ENCODING_TYPE.PKCS_7_ASN_ENCODING,
            type,
            pointer,
            count,
            0,
            0,
            mem.Pointer,
            ref size))
        {
            mem.Dispose();
            return null;
        }

        return mem;
    }

    private static unsafe CryptographicAttributeObject[] ReadAttributes(CRYPT_ATTRIBUTES attributes)
    {
        var array = new CryptographicAttributeObject[attributes.cAttr];
        var attrs = (CRYPT_ATTRIBUTE*)attributes.rgAttr;
        if (array.Length > 0)
        {
            var blobs = (CRYPT_INTEGER_BLOB*)attrs->rgValue;
            for (var i = 0; i < array.Length; i++)
            {
                var oid = new Oid(attrs[i].pszObjId.ToString()!);
                array[i] = new CryptographicAttributeObject(oid);
                for (var j = 0; j < attrs->cValue; j++)
                {
                    array[i].Values.Add(new AsnEncodedData(oid, Authenticode.AsSpan(blobs + j)));
                }
            }
        }
        return array;
    }
}
