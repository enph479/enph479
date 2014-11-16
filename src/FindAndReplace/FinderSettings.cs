namespace ElectricalToolSuite.FindAndReplace
{
    public class FinderSettings
    {
        public bool CaseSensitive { get; set; }
        public bool WholeWords { get; set; }
        public bool HiddenElements { get; set; }
        public bool DimensionText { get; set; }
        public bool IncludeTableText { get; set; }
        public bool FamilyProperties { get; set; }
        public SearchViewSettings SearchViewFilter { get; set; }
        public string SearchText { get; set; }
        public string ReplaceText { get; set; }

        public FinderSettings()
        {
            SearchViewFilter = SearchViewSettings.CurrentView;
            SearchText = "";
            ReplaceText = "";
        }
    }
}
