using Autodesk.Revit.UI;
using Redbolts.DockableUITest.Composition;
using Redbolts.DockableUITest.UI;

namespace Redbolts.DockableUITest
{
    public class ApplicationCommand:IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            var container = CommandContainer.Instance();
            if (container.Valid)
            {
                container.BuildRibbon(application);
            }

            var dPid = new DockablePaneId(DockConstants.Id);
            if (!DockablePane.PaneIsRegistered(dPid))
            {
                var state = new DockablePaneState {DockPosition = DockPosition.Right};
                var element = new DockPage();
                application.RegisterDockablePane(DockConstants.Id, DockConstants.Name, element, state);
            }


            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
