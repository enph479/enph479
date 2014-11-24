using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.FindAndReplace
{
    public class Globals
    {
        public static ElementId SelectedElement { get; set; }
        public static List<ResultsDto>  MatchingElementSet { get; set; }
    }
}
