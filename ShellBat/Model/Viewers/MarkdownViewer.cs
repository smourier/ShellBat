using Markdig;

namespace ShellBat.Model.Viewers;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class MarkdownViewer(Entry entry) : Viewer(entry)
{
    private static readonly MarkdownPipeline _pipeline = new MarkdownPipelineBuilder().
        UseAdvancedExtensions().
        UseEmojiAndSmiley().
        Build();

    public override int Priority => FileExtensionMatchPriority;
    public override string? Icon => "fa-brands fa-markdown";
    public override bool IsSupported
    {
        get
        {
            if (Entry.Extension.Name == ".md")
                return true;

            return base.IsSupported;
        }
    }

    public virtual string? GetText()
    {
        var stream = Entry.OpenRead();
        if (stream == null)
            return null;

        using var sw = new StreamReader(stream);
        var text = sw.ReadToEnd();
        var result = Markdown.ToHtml(text, _pipeline);
        return result;
    }
}
