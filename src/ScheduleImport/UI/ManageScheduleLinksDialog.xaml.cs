using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    /// <summary>
    /// Interaction logic for ManageScheduleLinksDialog.xaml
    /// </summary>
    public partial class ManageScheduleLinksDialog
    {
        public bool PressedCreate { get; private set; }

        public ManageScheduleLinksDialog(ManageLinksViewModel dataContext)
        {
            DataContext = dataContext;
            InitializeComponent();
        }

        private void ManagedScheduleLinksDataGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            ((ManageLinksViewModel)DataContext).EditCommand.RaiseCanExecuteChanged();
            ((ManageLinksViewModel)DataContext).RemoveCommand.RaiseCanExecuteChanged();
            ((ManageLinksViewModel)DataContext).ReloadCommand.RaiseCanExecuteChanged();
        }

        private void CommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            PressedCreate = true;
            DialogResult = true;
        }
    }
}
