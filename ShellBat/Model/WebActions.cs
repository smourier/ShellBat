namespace ShellBat.Model;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class WebActions(ShellBatWindow window, Entry entry) : DispatchObject, IActionExecutor
{
#pragma warning disable CA1822 // Mark members as static
    private Task<IDictionary<string, string>>? _openFromVisualStudioListTask;

    public bool IsNewFolderSupported => entry.IsNewFolderSupported;

    public IEnumerable<WebTerminal> GetTerminals() => Settings.Current.GetTerminals().Select(command => new WebTerminal(command, entry, null, null));
    public IDictionary<string, string> OpenFromExplorerList { get; } = new Dictionary<string, string>();
    public bool DetectsVisualStudioInstances { get; internal set; }

    public async Task<IDictionary<string, string>> GetOpenFromVisualStudioList()
    {
        RunScanVisualStudioSolutions();
        return await _openFromVisualStudioListTask;
    }

    [MemberNotNull(nameof(_openFromVisualStudioListTask))]
    public void RunScanVisualStudioSolutions() => _openFromVisualStudioListTask ??= Task.Run(ScanVisualStudioSolutions);
    private IDictionary<string, string> ScanVisualStudioSolutions()
    {
        var dic = new Dictionary<string, string>();
        foreach (var solution in VisualStudioSolution.Enumerate())
        {
            var dir = Path.GetDirectoryName(solution.FileName);
            if (!string.IsNullOrWhiteSpace(dir))
            {
                dic[dir] = $"<div class=vs-solution tooltip='{HttpUtility.HtmlAttributeEncode(dir)}'>{solution.Name}</div>";
            }

            foreach (var project in solution.Projects)
            {
                var projectDir = Path.GetDirectoryName(project.FileName);
                if (!string.IsNullOrWhiteSpace(projectDir))
                {
                    dic[projectDir] = $"<div class=vs-project tooltip='{HttpUtility.HtmlAttributeEncode(projectDir)}'>{project.Name}</div>";
                }

                if (project.OutputPath != null)
                {
                    dic[project.OutputPath] = $"<div class=vs-output tooltip='{HttpUtility.HtmlAttributeEncode(project.OutputPath)}'>{project.Name}</div>";
                }
            }
        }
        return dic;
    }

    public IReadOnlyList<WebMenuItem> GetEntryActions() => window.GetEntryActions(entry);
    public virtual void ExecuteAction(string id) => window.ExecuteAction((Entry[])[entry], id);

    protected override object? GetTaskResult(Task task)
    {
        if (task is Task<IDictionary<string, string>> dicTask)
            return dicTask.Result;

        return base.GetTaskResult(task);
    }
#pragma warning restore CA1822 // Mark members as static
}
