using System;
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

        public ElementSet FindMatchingElements(FilteredElementCollector allElements)
        {
            var matchingElements = new ElementSet();
            foreach (FamilyInstance elem in allElements.Cast<FamilyInstance>())
            {
                foreach (Parameter param in elem.Parameters)
                {
                    if (!String.IsNullOrEmpty(param.AsString())
                        && Regex.Match(param.AsString(), _searchText, _compareOptions).Success)
                        {
                            matchingElements.Insert(elem);
                        }
                }
            }
            return matchingElements;
        }
    }
}
