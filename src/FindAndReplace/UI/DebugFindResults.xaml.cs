using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.FindAndReplace
{
    /// <summary>
    /// Interaction logic for
    /// </summary>
    public partial class DebugFindResult
    {
        public Element SelectedElement { get; set; }
        public DebugFindResult()
        {
            InitializeComponent();
        }

        public void UpdateElements(List<ResultsDto> elements)
        {
            ListOfMatchingElements.ItemsSource = elements;
        }

        private void GetSelectedElementFromResults(object sender, RoutedEventArgs e)
        {
            var castedSender = ((ListBox) sender);
            Globals.SelectedElement = ((ResultsDto)castedSender.SelectedItem).MatchingElement.Id;
        } 
    }
}
