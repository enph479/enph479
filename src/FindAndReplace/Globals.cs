using Autodesk.Revit.DB;

namespace ElectricalToolSuite.FindAndReplace
{
    public class Globals
    {
        public static ElementId SelectedElement { get; set; }
        public static ElementSet MatchingElementSet { get; set; }
    }
}
