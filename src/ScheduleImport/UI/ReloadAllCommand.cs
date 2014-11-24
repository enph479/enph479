using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    class ReloadAllCommand : ICommand
    {
        private ManageLinksViewModel _viewModel;

        public ReloadAllCommand(ManageLinksViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        public bool CanExecute(object parameter)
        {
            return _viewModel.HasAnyLinks;
        }

        public void Execute(object parameter)
        {
            foreach (var link in _viewModel.Links)
            {
                link.Reload();
            }

            _viewModel.OnPropertyChanged("Links");
        }

        public event EventHandler CanExecuteChanged;
    }
}
