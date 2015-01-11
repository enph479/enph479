using System.Text.RegularExpressions;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.FindAndReplace
{
    class TextFinderBuilder
    {
        public static TextFinder BuildTextFinder(FinderSettings finderSettings, Document document)
        {
            //construct whether case sensitive or not
            var caseSensitive = RegexOptions.None;
            if (!finderSettings.CaseSensitive)
            {
                caseSensitive = RegexOptions.IgnoreCase;
            }
            //construct whether whole words matching or not
            var searchText = finderSettings.SearchText;
            if (finderSettings.WholeWords)
            {
                searchText = @"\b" + searchText + @"\b";
            }

            return new TextFinder(searchText, caseSensitive, finderSettings.HiddenElements);
        }
    }
}
