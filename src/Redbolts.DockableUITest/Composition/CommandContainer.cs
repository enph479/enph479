using System;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Linq;
using Autodesk.Revit.UI;
using Redbolts.UI.Common.Composition;
using Redbolts.UI.Common.Utility;

namespace Redbolts.DockableUITest.Composition
{
    public class CommandContainer
    {
        private readonly CompositionContainer _container;

        private static readonly CommandContainer INSTANCE = new CommandContainer();
        public static CommandContainer Instance()
        {
            return INSTANCE;
        }

        public bool Valid { get; set; }

        private CommandContainer()
        {
            try
            {
                var catalog = new AssemblyCatalog(typeof(CommandContainer).Assembly);
                _container = new CompositionContainer(catalog);
                _container.ComposeParts(this);
                Valid = true;
            }
            catch (CompositionContractMismatchException)
            {
                Valid = false;
            }
        }

        public void BuildRibbon(UIControlledApplication application)
        {
            foreach (var bc in _container.GetExports<IExternalCommand, IButtonMetaData>().OrderBy(l => l.Metadata.PanelIndex))
            {
                var md = bc.Metadata;
                var cmdType = bc.Value.GetType();
                var assembly = cmdType.Assembly;
                var panel = application.RibbonPanel(bc.Metadata.TabName, bc.Metadata.PanelName);
                var button = (PushButton)panel.AddItem(new PushButtonData(md.Name, md.Text, assembly.Location, cmdType.FullName));
                if (button == null) continue;

                if (!String.IsNullOrEmpty(md.LargeImage)) button.LargeImage = ImageUtil.GetEmbeddedImage(assembly,md.LargeImage);
                button.Enabled = md.Enabled;
                button.Visible = md.Visible;
            }
        }
    }
}