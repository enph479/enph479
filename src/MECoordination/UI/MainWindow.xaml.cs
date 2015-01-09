using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace ElectricalToolSuite.MECoordination.UI
{
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();

            MaxHeight = SystemParameters.MaximizedPrimaryScreenHeight;
            MaxWidth = SystemParameters.MaximizedPrimaryScreenWidth;
        }
        
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void CloseCommandHandler(object sender, ExecutedRoutedEventArgs e)
        {
        }

        private void CanExecuteHandler(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = MechanicalTree.Items.Cast<TreeViewItemWithCheckbox>().SelectMany(i => i.SelectedWithChildren).Any()
                && ElectricalTree.Items.Cast<TreeViewItemWithCheckbox>().SelectMany(i => i.SelectedWithChildren).Any()
                && (WorksetComboBox.SelectedItem != null || !WorksetComboBox.HasItems);
        }
    }
}