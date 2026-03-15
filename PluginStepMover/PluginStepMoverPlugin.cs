using System;
using System.ComponentModel.Composition;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace PluginStepMover;

[Export(typeof(IXrmToolBoxPlugin))]
[ExportMetadata("Name", "Plugin Step Mover")]
[ExportMetadata("Description", "Move or copy plugin steps between plugin types, with cross-environment support.")]
[ExportMetadata("SmallImageBase64", PluginBranding.SmallImageBase64)]
[ExportMetadata("BigImageBase64", PluginBranding.BigImageBase64)]
[ExportMetadata("BackgroundColor", "White")]
[ExportMetadata("PrimaryFontColor", "Black")]
[ExportMetadata("SecondaryFontColor", "DimGray")]
public class PluginStepMoverPlugin : PluginBase, IGitHubPlugin, IHelpPlugin
{
    private static readonly Guid PluginId = new("8d487e02-acdb-4469-b177-9ca5688d38b3");

    public string RepositoryName => "PluginStepMover";
    public string UserName => "Gdynam";
    public string HelpUrl => "https://github.com/Gdynam/PluginStepMover#readme";

    public override IXrmToolBoxPluginControl GetControl()
    {
        return new PluginStepMoverControl();
    }

    public override Guid GetId()
    {
        return PluginId;
    }
}
