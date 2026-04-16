using System.Security.Cryptography.X509Certificates;

namespace ShellBat.Utilities;

public static partial class Authenticode
{
    public static unsafe HRESULT GetSignStatus(string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        using var pwstr = new Pwstr(filePath);
        var file = new WINTRUST_FILE_INFO
        {
            cbStruct = sizeof(WINTRUST_FILE_INFO),
            pcwszFilePath = pwstr
        };

        var data = new WINTRUST_DATA
        {
            cbStruct = sizeof(WINTRUST_DATA),
            dwUIChoice = WTD_UI_NONE,
            dwUnionChoice = WTD_CHOICE_FILE,
            fdwRevocationChecks = WTD_REVOKE_NONE,
            pFile = (nint)(&file),
        };

        return WinVerifyTrust(HANDLE.Invalid, WINTRUST_ACTION_GENERIC_VERIFY_V2, ref data);
    }

    public static AuthenticodeSignature? GetSignature(string filePath) => GetSignature(filePath, out _);
    public static AuthenticodeSignature? GetSignature(string filePath, out HRESULT error)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        using var msg = Message.GetFromFile(filePath, out error);
        if (msg == null)
            return null;

        return msg.GetSignature(false, out error);
    }

    private struct WINTRUST_FILE_INFO
    {
        public int cbStruct;
        public PWSTR pcwszFilePath;
        public nint hFile;
        public nint pgKnownSubject;
    }

    private struct WINTRUST_DATA
    {
        public int cbStruct;
        public nint pPolicyCallbackData;
        public nint pSIPClientData;
        public int dwUIChoice;
        public int fdwRevocationChecks;
        public int dwUnionChoice;
        public nint pFile;
        public int dwStateAction;
        public nint hWVTStateData;
        public nint pwszURLReference;
        public int dwProvFlags;
        public int dwUIContext;
        public nint pSignatureSettings;
    }

#pragma warning disable IDE1006 // Naming Styles
    private const int WTD_UI_NONE = 2;
    private const int WTD_REVOKE_NONE = 0;
    private const int WTD_CHOICE_FILE = 1;

    private const uint CMSG_CONTENT_PARAM = 2;
    private const uint CMSG_SIGNER_COUNT_PARAM = 5;
    private const uint CMSG_SIGNER_INFO_PARAM = 6;
    private const uint CMSG_CERT_COUNT_PARAM = 11;
    private const uint CMSG_CERT_PARAM = 12;

    private static readonly Guid WINTRUST_ACTION_GENERIC_VERIFY_V2 = new("{00AAC56B-CD44-11d0-8CC2-00C04FC295EE}");
#pragma warning restore IDE1006 // Naming Styles

    [LibraryImport("wintrust.dll")]
    private static partial HRESULT WinVerifyTrust(HANDLE hwnd, in Guid pgActionID, ref WINTRUST_DATA pWVTData);

    internal static unsafe ReadOnlySpan<byte> AsSpan(CRYPT_INTEGER_BLOB* blob) => new(blob->pbData.ToPointer(), (int)blob->cbData);

    internal sealed partial class Message : IDisposable
    {
        private nint _handle;

        private Message(nint handle)
        {
            _handle = handle;
        }

        public nint Handle => _handle;

        ~Message() { Dispose(); }
        public void Dispose()
        {
            var handle = Interlocked.Exchange(ref _handle, 0);
            if (handle != 0)
            {
                ShellN.Functions.CryptMsgClose(handle);
                GC.SuppressFinalize(this);
            }
        }

        public unsafe AuthenticodeSignature? GetSignature(bool timestamp, out HRESULT error)
        {
            var size = (uint)sizeof(uint);
            ShellN.Functions.CryptMsgGetParam(Handle, CMSG_CONTENT_PARAM, 0, 0, ref size);
            if (size == 0)
            {
                error = Marshal.GetLastPInvokeError();
                return null;
            }

            using var content = new ComMemory(size);
            if (!ShellN.Functions.CryptMsgGetParam(Handle, CMSG_CONTENT_PARAM, 0, content.Pointer, ref size))
            {
                error = Marshal.GetLastPInvokeError();
                return null;
            }

            var signature = timestamp ? new TimestampAuthenticodeSignature(content.ToArray()) : new AuthenticodeSignature(content.ToArray());

            // we read certificates first because signer infos may refer to them
            var count = 0;
            size = sizeof(uint);
            ShellN.Functions.CryptMsgGetParam(Handle, CMSG_CERT_COUNT_PARAM, 0, (nint)(&count), ref size);
            if (size == 0)
            {
                error = Marshal.GetLastPInvokeError();
                return null;
            }

            for (uint i = 0; i < count; i++)
            {
                size = 0;
                ShellN.Functions.CryptMsgGetParam(Handle, CMSG_CERT_PARAM, i, 0, ref size);
                if (size == 0)
                {
                    error = Marshal.GetLastPInvokeError();
                    return null;
                }

                using var certInfo = new ComMemory(size);
                if (!ShellN.Functions.CryptMsgGetParam(Handle, CMSG_CERT_PARAM, i, certInfo.Pointer, ref size))
                {
                    error = Marshal.GetLastPInvokeError();
                    return null;
                }

                var cert = X509CertificateLoader.LoadCertificate(certInfo.AsSpan());
                signature._certificates.Add(cert);
            }

            count = 0;
            size = sizeof(uint);
            if (!ShellN.Functions.CryptMsgGetParam(Handle, CMSG_SIGNER_COUNT_PARAM, 0, (nint)(&count), ref size))
            {
                error = Marshal.GetLastPInvokeError();
                return null;
            }

            for (uint i = 0; i < count; i++)
            {
                size = 0;
                ShellN.Functions.CryptMsgGetParam(Handle, CMSG_SIGNER_INFO_PARAM, i, 0, ref size);
                if (size == 0)
                {
                    error = Marshal.GetLastPInvokeError();
                    return null;
                }

                using var signerInfo = new ComMemory(size);
                if (!ShellN.Functions.CryptMsgGetParam(Handle, CMSG_SIGNER_INFO_PARAM, i, signerInfo.Pointer, ref size))
                {
                    error = Marshal.GetLastPInvokeError();
                    return null;
                }

                var pinfo = (CMSG_SIGNER_INFO*)signerInfo.Pointer;
                var info = new AuthenticodeSignerInfo(signature, pinfo);
                signature._signerInfos.Add(info);
            }

            error = DirectN.Constants.S_OK;
            return signature;
        }

        public static Message? GetFromBlob(nint pointer, uint count)
        {
            var type = CERT_QUERY_ENCODING_TYPE.X509_ASN_ENCODING | CERT_QUERY_ENCODING_TYPE.PKCS_7_ASN_ENCODING;
            var msg = ShellN.Functions.CryptMsgOpenToDecode((uint)type, 0, 0, 0, 0, 0);
            if (msg == 0)
                return null;

            if (!ShellN.Functions.CryptMsgUpdate(msg, pointer, count, true))
                throw new Win32Exception(Marshal.GetLastWin32Error());

            return new Message(msg);
        }

        public static unsafe Message? GetFromFile(string filePath, out HRESULT error)
        {
            nint handle;
            if (!ShellN.Functions.CryptQueryObject(
                CERT_QUERY_OBJECT_TYPE.CERT_QUERY_OBJECT_FILE,
                PWSTR.From(filePath).Value,
                CERT_QUERY_CONTENT_TYPE_FLAGS.CERT_QUERY_CONTENT_FLAG_PKCS7_SIGNED_EMBED,
                CERT_QUERY_FORMAT_TYPE_FLAGS.CERT_QUERY_FORMAT_FLAG_BINARY,
                0, 0, 0, 0, 0, (nint)(&handle), 0))
            {
                error = Marshal.GetLastPInvokeError();
                return null;
            }

            error = DirectN.Constants.S_OK;
            return new Message(handle);
        }
    }
}
