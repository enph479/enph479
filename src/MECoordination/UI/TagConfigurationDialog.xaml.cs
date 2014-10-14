using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.MECoordination.UI
{
    /// <summary>
    /// Interaction logic for TagConfigurationDialog.xaml
    /// </summary>
    public partial class TagConfigurationDialog : Window
    {
        public TagConfigurationDialog()
        {
            InitializeComponent();
        }

        class Dog
        {
            public string Name { get; set; }
            public int Size { get; set; }

            public Dog(string name, int size)
            {
                this.Name = name;
                this.Size = size;
            }
        }

        private void CloseCommandHandler(object sender, ExecutedRoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void CanExecuteHandler(object sender, CanExecuteRoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void FamilyCategoryTagDataGrid_Loaded(object sender, RoutedEventArgs e)
        {
            var items = new List<Dog>();
            items.Add(new Dog("Fido", 10));
            items.Add(new Dog("Spark", 20));
            items.Add(new Dog("Fluffy", 4));

            // ... Assign ItemsSource of DataGrid.
            var grid = sender as DataGrid;
            grid.ItemsSource = items;
        }
    }
}
