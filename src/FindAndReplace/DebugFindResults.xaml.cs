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
            var temp = new ElementSet();
            temp.Insert(null);
            UpdateElements(temp);
        }

        public void UpdateElements(ElementSet elements)
        {
            ListOfMatchingElements.ItemsSource = elements;
        }

        private void GetSelectedElementFromResults(object sender, RoutedEventArgs e)
        {
            var castedSender = ((ListBox) sender);
            Globals.SelectedElement = ((Element)castedSender.SelectedItem).Id;
        } 
    }
}
