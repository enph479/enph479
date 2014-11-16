namespace ElectricalToolSuite.FindAndReplace
{
    class TextFinderBuilder
    {
        public static TextFinder BuildTextFinder(FinderSettings finderSettings)
        {
            return new TextFinder(finderSettings.SearchText);
        }

    }
}
