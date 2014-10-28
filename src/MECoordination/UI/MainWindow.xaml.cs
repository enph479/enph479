using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Autodesk.Revit.UI;

namespace ElectricalToolSuite.MECoordination.UI
{
    public partial class MainWindow
    {
        private readonly UIApplication _application;

        public MainWindow(UIApplication application)
        {
            if (application == null) throw new ArgumentNullException("application");
            _application = application;

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
                && WorksetComboBox.SelectedItem != null;
        }

        private void TagsButtonClick(object sender, RoutedEventArgs e)
        {
            const TaskDialogCommonButtons buttons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel;

            var dlgResult = TaskDialog.Show("Close to continue", "This window must close to configure tagging.  Continue?", buttons);

            if (dlgResult == TaskDialogResult.Ok)
            {
                _application.PostCommand(RevitCommandId.LookupPostableCommandId(PostableCommand.LoadedTags));
                Close();
            }
            else
            {
                Focus();
            }
        }
    }
}