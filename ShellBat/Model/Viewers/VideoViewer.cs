namespace ShellBat.Model.Viewers;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class VideoViewer(Entry entry) : Viewer(entry)
{
    public static IReadOnlyList<string> VideoFileExtensions => _videoFileExtensions.Value;
    private static readonly Lazy<IReadOnlyList<string>> _videoFileExtensions = new(() => (string[])[.. FileExtension.GetMediaFileExtensions("video")]);

    public override int Priority => FileExtensionMatchPriority - 1;
    public override string? Icon => "fa-regular fa-file-video";
    public override bool IsSupported
    {
        get
        {
            // currently we only support this through the webview2 so we need a file system path
            if (Entry.FileSystemPath == null)
                return false;

            if (VideoFileExtensions.Contains(Entry.Extension.Name))
                return true;

            if (Entry.Extension.PerceivedType == PERCEIVED.PERCEIVED_TYPE_VIDEO)
                return true;

            return base.IsSupported;
        }
    }

    public virtual bool AutoPlay { get; set; } = ShellBatInstance.Current.Settings.VideoAutoPlay;

    public string? GetPageUrl()
    {
        if (!IsSupported)
            return null;

        return Entry.FileSystemPath;
    }

    public override Task<WebPropertyGrid?> GetProperties()
    {
        var grid = new WebPropertyGrid();

        grid.Properties.Add(new WebPropertyGridProperty()
        {
            Key = nameof(AutoPlay),
            DisplayName = Res.AutoPlay,
            Editor = "Boolean",
        });

        grid.Instance.PropertyChanged += (s, e) =>
        {
            switch (e.PropertyName)
            {
                case nameof(AutoPlay):
                    AutoPlay = grid.Instance.GetTypedValue<bool>(nameof(AutoPlay));
                    break;
            }
        };

        grid.Instance.SetValue(nameof(AutoPlay), AutoPlay);

        grid.Options.GroupByCategory = false;
        grid.Options.PropertyNamePostfix = ":";
        grid.Options.BaseClassName = "viewer-Video-pg";
        return Task.FromResult<WebPropertyGrid?>(grid);
    }
}
