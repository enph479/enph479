using System;
using System.Linq;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.FindAndReplace
{
    class TextFinder
    {
        private readonly String _searchText;
        public TextFinder(String searchText)
        {
            this._searchText = searchText;
        }

        public ElementSet FindMatchingElements(FilteredElementCollector allElements)
        {
            var matchingElements = new ElementSet();
            foreach (FamilyInstance elem in allElements.Cast<FamilyInstance>())
            {
                foreach (Parameter param in elem.Parameters)
                {
                    if (!String.IsNullOrEmpty(param.AsString()) && param.AsString().Contains(_searchText))
                    {
                        matchingElements.Insert(elem);
                    }
                }
            }
            return matchingElements;
        }

    }
}
