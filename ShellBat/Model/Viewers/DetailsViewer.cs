namespace ShellBat.Model.Viewers;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public partial class DetailsViewer(Entry entry) : Viewer(entry)
{
    public override int SortOrder => int.MinValue + 1;
    public override bool IsSupported => GetDetailsProperties().Any();

    public override Task<WebPropertyGrid?> GetProperties()
    {
        var grid = new WebPropertyGrid();

        foreach (var prop in GetDetailsProperties(GETPROPERTYSTOREFLAGS.GPS_DEFAULT))
        {
            grid.Properties.Add(prop.Property);
            grid.Instance.SetValue(prop.Property.Key, prop.Value);
        }

        grid.Options.IsReadOnly = true;
        grid.Options.GroupByCategory = true;
        grid.Options.BaseClassName = "fld-pg";

        grid.Options.Categories.Add(new WebPropertyGridCategory()
        {
            Key = "System",
            Collapsed = true
        });

        grid.Options.Categories.Add(new WebPropertyGridCategory()
        {
            Key = "System.Security",
            Collapsed = true
        });
        return Task.FromResult<WebPropertyGrid?>(grid);
    }

    protected virtual bool IncludeProperty(PropertyDescription description)
    {
        ArgumentNullException.ThrowIfNull(description);

        var name = description.Name;
        if (name == null)
            return true; // ok

        if (description.PropertyKey == ShellN.PropertyKeys.System.SFGAOFlags ||
            description.PropertyKey == ShellN.PropertyKeys.System.Security.AllowedEnterpriseDataProtectionIdentities ||
            description.PropertyKey == ShellN.PropertyKeys.System.Document.DateCreated ||
            description.PropertyKey == ShellN.PropertyKeys.System.Document.DateSaved ||
            description.PropertyKey == ShellN.PropertyKeys.System.Shell.SFGAOFlagsStrings)
            return false;

        // GPS and others produce both Numerator and Denominator properties for ratios; hide them
        if (endsWith("Numerator"))
            return false;

        if (endsWith("Denominator"))
            return false;

        return true;
        bool endsWith(string s) => name.Length > s.Length && name.EndsWith(s);
    }

    protected virtual string? ProcessName(string? name)
    {
        if (name == null)
            return null;

        return Conversions.Decamelize(name, DecamelizeOptions.ForceRestLower);
    }

    public virtual IEnumerable<PropertyValue> GetDetailsProperties(GETPROPERTYSTOREFLAGS flags = GETPROPERTYSTOREFLAGS.GPS_FASTPROPERTIESONLY)
    {
        using var item = Entry.GetItem(Entry.ParsingName, ShellItemParsingOptions.DontThrowOnError);
        if (item == null)
            yield break;

        foreach (var prop in item.EnumerateProperties(flags, false))
        {
            if (prop.Value == null)
                continue;

            if (prop.Value is string s && string.IsNullOrWhiteSpace(s))
                continue;

            if (prop.Value is Array arr && arr.Length == 0)
                continue;

            if (prop.Value is false)
                continue;

            var value = prop.Value;
            if (prop.Value is byte[] bytes)
            {
                value = bytes.ToHexa(16, true);
            }

            using var desc = prop.Key.ToDescription();
            if (desc is null)
                continue;

            var canonicalName = desc.CanonicalName;
            if (canonicalName == null)
                continue;

            if (!IncludeProperty(desc))
                continue;

            yield return new PropertyValue(new WebPropertyGridProperty()
            {
                Key = canonicalName,
                Tooltip = $"{canonicalName} ({prop.Key})",
                DisplayName = desc.DisplayName ?? ProcessName(desc.Name),
                CategoryName = desc.Namespace,
                IsReadOnly = true,
            }, value);
        }
    }

    public sealed class PropertyValue(WebPropertyGridProperty property, object? value)
    {
        public WebPropertyGridProperty Property { get; } = property;
        public object? Value { get; } = value;
    }
}
