using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ElectricalToolSuite.MECoordination
{
    [Transaction(TransactionMode.Automatic)]
    public class ExternalCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var wnd = new UI.Window1();
            wnd.ShowDialog();

            var selected = wnd.SelectedItems.ToList();

            var sb = new StringBuilder();

            sb.AppendLine("You selected:");
            foreach (var item in selected)
            {
                sb.AppendLine(item);
            }

            TaskDialog.Show("Selected items", sb.ToString());

            return Result.Succeeded;
        }
    }
}
