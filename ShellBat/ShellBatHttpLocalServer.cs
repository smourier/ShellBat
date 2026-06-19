namespace ShellBat;

public partial class ShellBatHttpLocalServer : IDisposable
{
    private HttpListener? _listener;
    private bool _stopped;

    private ShellBatHttpLocalServer(HttpListener listener, int port)
    {
        _listener = listener;
        Port = port;
        //LogRequests = true;
    }

    public int Port { get; }
    public bool IsDisposed => _listener == null;
    public Uri Url => new(GetUrl(Port));
    public bool LogRequests { get; set; }

    public override string ToString() => Url.ToString();

    public static Task<string?> GetImageCacheFilePath(UrlType type, string parsingName, int imageSize, GetImageOptions options = GetImageOptions.None)
    {
        ArgumentNullException.ThrowIfNull(parsingName);
        if (type != UrlType.Image && type != UrlType.Icon)
            throw new ArgumentOutOfRangeException(nameof(type));

        var name = $"{parsingName}\0{imageSize}\0{(int)options}".ComputeGuidHash().ToString("N");
        return EnsureCache(type, parsingName, name, () => GetImage(
                                        parsingName,
                                        imageSize,
                                        options), imageSize, options);
    }

    protected virtual async Task<HttpStatusCode> Get(HttpListenerRequest request, Stream outputStream, HttpListenerContext context)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(outputStream);
        if (request.Url != null)
        {
            var segments = request.Url.LocalPath.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (segments.Length > 0)
            {
                if (Enum.TryParse<UrlType>(segments[0], true, out var type))
                {
                    string? cacheFileName = null;
                    string parsingName;
                    string name;
                    var imageOptions = GetImageOptions.None;
                    if (CacheControlHeaderValue.TryParse(request.Headers["Cache-Control"], out var cacheControl))
                    {
                        if (cacheControl.NoCache || cacheControl.NoStore)
                        {
                            imageOptions |= GetImageOptions.DontCache;
                        }
                        else if (cacheControl.OnlyIfCached)
                        {
                            imageOptions |= GetImageOptions.CacheOnly;
                        }
                    }

                    switch (type)
                    {
                        case UrlType.Info:
                            outputStream.Write(Encoding.UTF8.GetBytes($"{Program.AppId};Pid:{SystemUtilities.CurrentProcess.Id}"));
                            return HttpStatusCode.OK;

                        case UrlType.Image:
                        case UrlType.Icon:
                            if (segments.Length < 2)
                                break;

                            var imageSize = Math.Max(-1, request.QueryString.GetValue<int>("size", CultureInfo.InvariantCulture));
                            imageOptions = request.QueryString.GetValue("options", CultureInfo.InvariantCulture, imageOptions);

                            parsingName = segments[1].Replace('/', '\\');
                            name = $"{parsingName}\0{imageSize}\0{(int)imageOptions}".ComputeGuidHash().ToString("N");
                            cacheFileName = await EnsureCache(type, parsingName, name, () => GetImage(
                                parsingName,
                                imageSize,
                                imageOptions), imageSize, imageOptions);
                            break;

                        case UrlType.PdfPage:
                            if (segments.Length < 2)
                                break;

                            var renderer = ShellBatWindow.GetPdfRenderer();
                            if (renderer != null)
                            {
                                var page = request.QueryString.GetValue<uint>("page", CultureInfo.InvariantCulture);
                                imageOptions = request.QueryString.GetValue("options", CultureInfo.InvariantCulture, imageOptions);

                                parsingName = segments[1].Replace('/', '\\');
                                var localid = request.QueryString.GetValue<int>("id", CultureInfo.InvariantCulture);
                                var pdfSize = request.QueryString.GetValue<uint>("size", CultureInfo.InvariantCulture);
                                var ignoreHighContrast = request.QueryString.GetValue<bool>("ignorehc", CultureInfo.InvariantCulture);
                                var pdfOptions = new PdfRenderOptions(pdfSize, pdfSize, page)
                                {
                                    IgnoreHighContrast = ignoreHighContrast
                                };

                                name = $"{parsingName}\0{page}\0{pdfOptions.Key}".ComputeGuidHash().ToString("N");
                                try
                                {
                                    cacheFileName = await EnsureCache(type, parsingName, name, () => GetPdfPageImage(renderer, parsingName, localid, pdfOptions), 0, imageOptions);
                                }
                                catch (Exception ex)
                                {
                                    if (ex.HResult != DirectN.Constants.ERROR_WRONG_PASSWORD || !pdfOptions.RenderPasswordProtectedImage)
                                        throw;

                                    // happens at least on PDFs
                                    cacheFileName = Path.Combine(Settings.WebRootPath, "Assets", "PasswordProtected.png");
                                }
                            }
                            break;

                        case UrlType.AssocHandler:
                            if (segments.Length < 3)
                                break;

                            var ext = UnescapeUrl(segments[1]);
                            var key = UnescapeUrl(segments[2]);
                            var shil = request.QueryString.GetValue("shil", CultureInfo.InvariantCulture, SHIL.SHIL_JUMBO);
                            var ild = request.QueryString.GetValue("ild", CultureInfo.InvariantCulture, IMAGE_LIST_DRAW_STYLE.ILD_NORMAL);

                            name = $"assochandler:{ext}\0{key}\0{(int)shil}\0{(int)ild}".ComputeGuidHash().ToString("N");
                            cacheFileName = await EnsureCache(type, null, name, () => GetAssocHandlerIcon(ext, key, shil, ild), 0, imageOptions);
                            break;
                    }

                    if (cacheFileName != null)
                    {
                        var fi = new FileInfo(cacheFileName);
                        if (IsFileOk(fi))
                        {
                            if (context.Request.Headers["If-Modified-Since"] is string ims && fi.LastWriteTimeUtc.ToString("r").EqualsIgnoreCase(ims))
                                return HttpStatusCode.NotModified;

                            AddAllowOrigin(context.Response);
                            AddContentTypeHeader(context.Response, Path.GetExtension(cacheFileName));
                            context.Response.Headers.Add("Last-Modified", fi.LastWriteTimeUtc.ToString("r"));
                            using var fs = new FileStream(cacheFileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                            await fs.CopyToAsync(outputStream).ConfigureAwait(false);
                            return HttpStatusCode.OK;
                        }
                    }
                }
            }
        }

        context.Response.Headers.Add("X-Powered-By", Program.AppId);
        return HttpStatusCode.NotFound;
    }

    public static string EscapeUrl(string str) => Uri.EscapeDataString(str);
    private static string UnescapeUrl(string str) => Uri.UnescapeDataString(str);
    private static void AddAllowOrigin(HttpListenerResponse response) => response.Headers.Add("Access-Control-Allow-Origin", "*");
    private static void AddContentTypeHeader(HttpListenerResponse response, string? extension)
    {
        if (string.IsNullOrEmpty(extension))
            return;

        var mime = FileExtension.GetMimeTypesByExtension(extension).FirstOrDefault();
        if (mime == null)
            return;

        response.ContentType = mime;
    }

    private static async Task<string?> EnsureCache(
        UrlType type,
        string? parsingName,
        string fileNameWithoutExtension,
        Func<Task<WicBitmapSource?>> getBitmap,
        int size,
        GetImageOptions options)
    {
        var cacheFileName = Path.Combine(Settings.HttpServerCachePath, type.ToString(), fileNameWithoutExtension + ".png");
        var noCache = options.HasFlag(GetImageOptions.DontCache);
        if (noCache || !IsFileOk(cacheFileName))
        {
            var bmp = await getBitmap();
            if (bmp != null)
            {
                FileSystemUtilities.WrapSharingViolations(() =>
                {
                    IOUtilities.FileEnsureDirectory(cacheFileName);
                    bmp.Save(cacheFileName);
                    File.SetLastWriteTime(cacheFileName, DateTime.Now);
                    bmp.Dispose();
                }, (error, count) =>
                {
                    ShellBatInstance.LogWarning($"Error saving cache file '{cacheFileName}' (attempt {count}): {error}");
                    return true;
                });
            }
            else
            {
                if ((type == UrlType.Image || type == UrlType.Icon) && parsingName != null)
                {
                    // may be a virtual item (like an entry in a .zip), try to extract to a 1st level cache
                    using var item = ShellItem.FromParsingName(parsingName, throwOnError: false);
                    if (item != null)
                    {
                        // keep extension if any
                        var extension = item.Extension.Nullify();
                        var name = parsingName.ComputeGuidHash().ToString("B");
                        var virtualCacheFileName = Path.Combine(Settings.HttpServerCachePath, UrlType.Image.ToString(), name + extension);
                        if (noCache || !IsFileOk(virtualCacheFileName))
                        {
                            IOUtilities.FileEnsureDirectory(virtualCacheFileName);
                            IOUtilities.FileDelete(virtualCacheFileName);
                            using var folder = ShellItem.FromParsingName(Path.GetDirectoryName(virtualCacheFileName)!, throwOnError: false);
                            if (folder != null)
                            {
                                using var fo = new FileOperation();
                                fo.SetOperationFlags(FILEOPERATION_FLAGS.FOF_NO_UI);
                                var hr = fo.CopyItem(item, folder, Path.GetFileName(virtualCacheFileName), false);
                                if (hr.IsError)
                                {
                                    ShellBatInstance.LogWarning($"CopyItems failed on '{parsingName}' => {virtualCacheFileName}: {hr}");
                                }
                                else
                                {
                                    hr = fo.PerformOperations(false);
                                    if (hr.IsError)
                                    {
                                        ShellBatInstance.LogWarning($"PerformOperations failed on '{parsingName}' => {virtualCacheFileName}: {hr}");
                                    }
                                }
                            }
                        }

                        if (IOUtilities.PathIsFile(virtualCacheFileName))
                        {
                            NormalizeSize(ref size);
                            var virtualBmp = await Entry.GetImage(virtualCacheFileName, size, options);
                            if (virtualBmp != null)
                            {
                                FileSystemUtilities.WrapSharingViolations(() =>
                                {
                                    IOUtilities.FileEnsureDirectory(cacheFileName);
                                    virtualBmp.Save(cacheFileName);
                                    File.SetLastWriteTime(cacheFileName, DateTime.Now);
                                    virtualBmp.Dispose();
                                });
                            }
                        }
                    }
                }
            }
        }
        return cacheFileName;
    }

    private static void NormalizeSize(ref int size)
    {
        if (size <= 0)
        {
            size = 16;
        }
    }

    private static async Task<WicBitmapSource?> GetImage(string parsingName, int size, GetImageOptions options)
    {
        NormalizeSize(ref size);
        var bmp = await Entry.GetImage(parsingName, size, options);
        if (bmp == null)
        {
            ShellBatInstance.LogWarning($"Could not get image for '{parsingName}'");
        }
        return bmp;
    }

    private static async Task<WicBitmapSource?> GetPdfPageImage(PdfRenderer renderer, string parsingName, int localId, PdfRenderOptions options)
    {
        Application.TraceInfo($"Serving PDF page image from local document for '{parsingName}' (local id {localId}) size:{options.Width}");
        if (localId > 0)
        {
            var doc = PdfViewer.GetDocument(localId);
            if (doc != null)
            {
                var bmpLocal = renderer.RenderPdfPage(doc, options);
                Application.TraceInfo($"bmp:{bmpLocal?.Width}");
                if (bmpLocal == null)
                {
                    ShellBatInstance.LogWarning($"Could not get PDF page image for '{parsingName}' (local id {localId})");
                }
                return bmpLocal;
            }
        }

        var bmp = await renderer.RenderPdfPage(parsingName, options);
        if (bmp == null)
        {
            ShellBatInstance.LogWarning($"Could not get PDF page image for '{parsingName}'");
        }
        return bmp;
    }

    private static Task<WicBitmapSource?> GetAssocHandlerIcon(string ext, string key, SHIL shil, IMAGE_LIST_DRAW_STYLE flags)
    {
        Application.TraceInfo($"Serving asso handler icon for ext '{ext}' key '{key}' shil {shil} flags {flags}");
        var fext = FileExtension.GetByExtension(ext);
        if (fext == null)
        {
            ShellBatInstance.LogWarning($"Could not find file extension for ext '{ext}' key '{key}'");
            return Task.FromResult<WicBitmapSource?>(null);
        }

        using var handler = fext.GetAssocHandler(h => h.Key.EqualsIgnoreCase(key));
        if (handler == null)
        {
            ShellBatInstance.LogWarning($"Could not find assoc handler for ext '{ext}' key '{key}'");
            return Task.FromResult<WicBitmapSource?>(null);
        }

        var loc = handler.IconLocation;
        if (loc == null)
        {
            ShellBatInstance.LogWarning($"Assoc handler for ext '{ext}' key '{key}' has no icon location");
            return Task.FromResult<WicBitmapSource?>(null);
        }

        if (loc.IsPathIndirect)
        {
            var path = loc.ResolvedPath;
            var pathExt = Path.GetExtension(path);
            if (WicImagingComponent.DecoderFileExtensions.Contains(pathExt))
                return GetImage(path, shil.GetSize(), GetImageOptions.None);
        }

        var icon = loc.GetIconHandle(shil, flags, throwOnError: false);
        if (icon == 0)
        {
            ShellBatInstance.LogWarning($"Could not get icon handle for progid '{key}'");
            return Task.FromResult<WicBitmapSource?>(null);
        }

        try
        {
            var bmp = WicBitmapSource.FromHICON(icon);
            if (bmp != null)
            {
                if (bmp.PixelFormat != WicPixelFormat.GUID_WICPixelFormat32bppPBGRA)
                {
                    bmp.ConvertTo(WicPixelFormat.GUID_WICPixelFormat32bppPBGRA);
                }
            }
            return Task.FromResult(bmp);
        }
        catch (Exception ex)
        {
            ShellBatInstance.LogWarning($"Could not create bitmap from icon for progid '{key}': {ex}");
            return Task.FromResult<WicBitmapSource?>(null);
        }
        finally
        {
            DirectN.Functions.DestroyIcon(icon);
        }
    }

    public virtual void Stop()
    {
        var listener = _listener;
        if (listener == null || !listener.IsListening)
            return;

        // we must do this or a race condition can occur and ProcessRequest can be called after Stop...
        _stopped = true;
        listener.Stop();
    }

    public virtual void Start()
    {
        ObjectDisposedException.ThrowIf(_listener == null, this);
        _listener.Start();
        _listener.BeginGetContext(ProcessRequest, null);
    }

    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            Stop();
            Interlocked.Exchange(ref _listener, null)?.SafeDispose();
        }
    }

    private async void ProcessRequest(IAsyncResult result)
    {
        var listener = _listener;
        if (listener == null || !listener.IsListening || _stopped)
            return;

        try
        {
            var ctx = listener.EndGetContext(result);
            listener.BeginGetContext(ProcessRequest, null);

            try
            {
                await ProcessRequest(ctx).ConfigureAwait(false);
            }
            catch (Exception e)
            {
                HandleError(ctx, e);
            }
        }
        catch (Exception e2)
        {
            ShellBatInstance.LogError("Error(" + e2.GetType().FullName + "): " + e2);
        }
    }

    protected virtual void HandleError(HttpListenerContext context, Exception error)
    {
        ArgumentNullException.ThrowIfNull(context);
        if (error == null)
            return;

        ShellBatInstance.LogError("Error(" + error.GetType().FullName + "): " + error);
        try
        {
            if (error is HttpRequestException && error.InnerException is WebException we)
            {
                context.Response.StatusCode = (int)HttpStatusCode.BadGateway;
                var text = we.GetAllMessagesWithDots();
                var bytes = Encoding.UTF8.GetBytes(text);
                context.Response.OutputStream.Write(bytes, 0, bytes.Length);
                return;
            }

            context.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
        }
        finally
        {
            context.Response.OutputStream.Close();
        }
    }

    protected virtual async Task ProcessRequest(HttpListenerContext context)
    {
        ArgumentNullException.ThrowIfNull(context);

        StringBuilder? log = null;
        if (LogRequests)
        {
            log = new StringBuilder();
            log.AppendLine($"<{context.Request.ProtocolVersion} {context.Request.HttpMethod} {context.Request.RawUrl}");
            foreach (var headerName in context.Request.Headers.OfType<string>())
            {
                var headerValue = context.Request.Headers[headerName];
                log.AppendLine($"{headerName}: {headerValue}");
            }
        }

        using var os = context.Response.OutputStream;
        var code = await Get(context.Request, os, context).ConfigureAwait(false);
        context.Response.StatusCode = (int)code;
        context.Response.StatusDescription = code.ToString();

        if (context.Response.StatusCode >= 400)
        {
            ShellBatInstance.LogVerbose($"{context.Response.StatusCode} {context.Request.HttpMethod} {context.Request.RawUrl}");
        }

        if (log != null)
        {
            log.AppendLine($">{context.Response.StatusCode} {context.Request.HttpMethod} {context.Request.RawUrl}");
            ShellBatInstance.LogInfo(log.ToString());
        }
    }

    internal static ShellBatHttpServer Get(ShellBatInstance instance)
    {
        var ports = (Settings.Current.LastHttpServerPorts.SplitToList<int>([',']) ?? []).ToHashSet();
        CreateListener(instance, ports, out var listener, out var chosenPort, out var remoteServer);
        if (listener == null)
            return remoteServer!; // was started by another ShellBat process

        var lastPorts = Settings.Current.LastHttpServerPorts.SplitToList<int>([',']) ?? [];
        if (!ports.SetEquals(lastPorts))
        {
            Task.Run(() =>
            {
                Settings.Current.LastHttpServerPorts = string.Join(',', ports).Nullify();
                Settings.Current.SerializeToConfiguration();
            });
        }

        var localServer = new ShellBatHttpLocalServer(listener, chosenPort);
        return new ShellBatHttpServer(localServer.Url, localServer);
    }

    private static void CreateListener(
        ShellBatInstance instance,
        HashSet<int> preferredPorts,
        out HttpListener? httpListener,
        out int chosenPort,
        out ShellBatHttpServer? remoteServer)
    {
        remoteServer = null;
        if (preferredPorts.Count > 0)
        {
            foreach (var port in preferredPorts.ToArray())
            {
                var listener = new HttpListener();
                listener.Prefixes.Add(GetUrl(port));
                try
                {
                    listener.Start();
                    chosenPort = port;
                    httpListener = listener;
                    preferredPorts.Add(port);
                    return;
                }
                catch
                {
                    foreach (var remote in instance.RemoteInstances)
                    {
                        var uri = remote.HttpServerUrl;
                        if (uri != null && uri.Port == port)
                        {
                            remoteServer = new ShellBatHttpServer(uri, null) { ProcessId = remote.ProcessId };
                            httpListener = null;
                            chosenPort = port;
                            return;
                        }
                    }
                    // continue (the listener disposes itself when Start throws)
                }
            }
        }

        // we keep max of 5 ports
        while (preferredPorts.Count > 5)
        {
            preferredPorts.Remove(preferredPorts.First());
        }

        chosenPort = FindFreePort();
        httpListener = new HttpListener();
        httpListener.Prefixes.Add(GetUrl(chosenPort));
        preferredPorts.Add(chosenPort);
        httpListener.Start();
    }

    private static bool IsFileOk(string filePath) => IOUtilities.FileGetSize(filePath) > 64;
    private static bool IsFileOk(FileInfo info) => info.Exists && info.Length > 64;
    private static string GetUrl(int port) => $"http://localhost:{port}/";
    private static int FindFreePort()
    {
        using var s = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        s.Bind(IPEndPoint.Parse("127.0.0.1:0"));
        return ((IPEndPoint)s.LocalEndPoint!).Port;
    }
}
