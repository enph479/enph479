using System.Collections.ObjectModel;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.FindAndReplace
{
    public class ResultsDto
    {
        public Element MatchingElement { get; set; }
        public ObservableCollection<MatchingParameterDto> MatchingParameters { get; set; }

        public ResultsDto(Element element, ObservableCollection<MatchingParameterDto> matchingParameters)
        {
            MatchingElement = element;
            MatchingParameters = matchingParameters;
        }
    }
}
