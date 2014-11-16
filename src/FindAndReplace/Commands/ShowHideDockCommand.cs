using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Redbolts.UI.Common.Composition;
using Redbolts.UI.Common.Utility;

namespace ElectricalToolSuite.FindAndReplace.Commands
{
    [Button(RibbonConstants.Tab, RibbonConstants.Panel, ElectricalToolSuite.FindAndReplace.CommandConstants.Name, ElectricalToolSuite.FindAndReplace.CommandConstants.Hide,
        LargeImage = "Redbolts.DockableUITest.Images.HideDock.png")]
    [Transaction(TransactionMode.Manual)]
    public class ShowHideDockCommand : IExternalCommand
    {
        private static bool _state = true;
        private static readonly Assembly Assembly = typeof(ShowHideDockCommand).Assembly;
        private static DockablePane _pane = null;
        private static PushButton _button = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (_button == null) _button = commandData.Application.GetRibbonItem<PushButton>(RibbonConstants.Tab, RibbonConstants.Panel, ElectricalToolSuite.FindAndReplace.CommandConstants.Name);
            if (_pane == null)
            {
                // if the pane isn't registered then disable show/hide
                if (!commandData.Application.PaneExists(DockConstants.Id, out _pane))
                {
                    _button.Enabled = false;
                    return Result.Succeeded;
                }
            }
            if (_state)
            {
                // showing so hide
                _pane.Hide();
                _button.ItemText = ElectricalToolSuite.FindAndReplace.CommandConstants.Show;
                _button.LargeImage = ImageUtil.GetEmbeddedImage(Assembly, "Redbolts.DockableUITest.Images.ShowDock.png");
            }
            else
            {
                //hidden so show
                _pane.Show();
                _button.ItemText = ElectricalToolSuite.FindAndReplace.CommandConstants.Hide;
                _button.LargeImage = ImageUtil.GetEmbeddedImage(Assembly, "Redbolts.DockableUITest.Images.HideDock.png");
            }
            _state = !_state;

            return Result.Succeeded;
        }
    }
}
