using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.FindAndReplace
{
    /// <summary>
    /// Interaction logic for FindResultsUI.xaml
    /// </summary>
    public partial class FindResultsWindow
    {
        public Element SelectedElement { get; set; }
        public FindResultsWindow()
        {
            InitializeComponent();
            //ListOfMatchingElements.AddHandler(, new RoutedEventHandler(GetSelectedElementFromResults));
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
            //Close();
        } 
    }
}
