using System.Collections.Generic;

namespace ElectricalToolSuite.FindAndReplace.UI
{
    /// <summary>
    /// Interaction logic for SelectedViewResults.xaml
    /// </summary>
    public partial class SelectedViewResults
    {
        public SelectedViewResults(List<ViewSelectorDto> allViews)
        {
            InitializeComponent();
            ListOfViews.ItemsSource = allViews;
        }
    }
}
