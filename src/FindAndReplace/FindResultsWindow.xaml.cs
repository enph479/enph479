using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ElectricalToolSuite.FindAndReplace
{
    /// <summary>
    /// Interaction logic for FindResultsUI.xaml
    /// </summary>
    public partial class FindResultsWindow
    {
        public SelElementSet SelectedElements { get; set; }
        public FindResultsWindow()
        {
            InitializeComponent();
            ListOfMatchingElements.AddHandler(MouseDoubleClickEvent, new RoutedEventHandler(GetSelectedElementFromResults));
        }

        public void UpdateElements(ElementSet elements)
        {
            ListOfMatchingElements.ItemsSource = elements;
        }

        private void GetSelectedElementFromResults(object sender, RoutedEventArgs e)
        {
            var castedSender = ((ListBox) sender);
            var t = ((Element) castedSender.SelectedItem);
            SelectedElements.Clear();
            SelectedElements.Insert(t);
            Close();
        } 
    }
}
