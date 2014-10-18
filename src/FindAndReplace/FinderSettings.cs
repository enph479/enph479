using System;

namespace ElectricalToolSuite.FindAndReplace
{
    public class FinderSettings
    {
        public Boolean _caseSensitive { get; set; }
        public Boolean _wholeWords { get; set; }
        public Boolean _hiddenElements { get; set; }
        public Boolean _dimensionText { get; set; }
        public Boolean _includeTableText { get; set; }
        public Boolean _familyProperties { get; set; }
        public SearchViewSettings _searchViewFilter { get; set; }
        public String _searchText { get; set; }
        public String _replaceText { get; set; }

        public FinderSettings()
        {
            _caseSensitive = false;
            _wholeWords = false;
            _hiddenElements = false;
            _dimensionText = false;
            _includeTableText = false;
            _familyProperties = false;
            _searchViewFilter = SearchViewSettings.CurrentView;
            _searchText = "";
            _replaceText = "";
        }

        public override String ToString()
        {
            String newString = "";
            newString += this._searchText + "\n";
            newString += this._replaceText + "\n";
            newString += this._caseSensitive.ToString() + this._wholeWords.ToString() + this._hiddenElements.ToString() +
                this._dimensionText.ToString() + this._includeTableText.ToString() + this._familyProperties.ToString() + "\n";
            newString += this._searchViewFilter.ToString();
            return newString;
        }
    }
}
