using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.FindAndReplace
{
    class TextFinder
    {
        private readonly String _searchText;
        private readonly RegexOptions _compareOptions;
        public TextFinder(String searchText, RegexOptions compareOptions, Document document)
        {
            _searchText = searchText;
            _compareOptions = compareOptions;
        }

        public List<ResultsDto> FindMatchingElements(FilteredElementCollector allElements,
            FilteredElementCollector allTextBoxes)
        {
            var matchingElements = new List<ResultsDto>();

            foreach (FamilyInstance elem in allElements.Cast<FamilyInstance>())
            {
                var matchingParamList = new List<MatchingParameterDto>();
                foreach (Parameter param in elem.Parameters)
                {
                    if (!String.IsNullOrEmpty(param.AsString())
                        && Regex.Match(param.AsString(), _searchText, _compareOptions).Success)
                    {
                        matchingParamList.Add(new MatchingParameterDto(param.ToString(), param.AsString()));
                    }   
                }
                if (matchingParamList.Count > 0)
                {
                    matchingElements.Add(new ResultsDto(elem, new ObservableCollection<MatchingParameterDto>(matchingParamList)));
                }
            }

            foreach (TextNote elem in allTextBoxes.Cast<TextNote>())
            {
                var matchingParamList = new List<MatchingParameterDto>();
                if (!String.IsNullOrEmpty(elem.Text)
                        && Regex.Match(elem.Text, _searchText, _compareOptions).Success)
                    {
                        matchingParamList.Add(new MatchingParameterDto("TextBox Text", elem.Text));
                    }
                if (matchingParamList.Count > 0)
                {
                    matchingElements.Add(new ResultsDto(elem, new ObservableCollection<MatchingParameterDto>(matchingParamList)));
                }
            }
            return matchingElements;
        }
    }
}
