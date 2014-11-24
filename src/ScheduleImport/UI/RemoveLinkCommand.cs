using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    class RemoveLinkCommand : ICommand
    {
        private ManageLinksViewModel _viewModel;

        public RemoveLinkCommand(ManageLinksViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        public bool CanExecute(object parameter)
        {
            var selectedLink = parameter as ManagedScheduleLink;
            
            if (selectedLink == null)
                return false;

            return true;
        }

        public void Execute(object parameter)
        {
            var selectedLink = parameter as ManagedScheduleLink;

            if (selectedLink == null)
                throw new ArgumentException("Parameter to RemoveCommand must be the selected ManagedScheduleLink");

            if (LinkGateway.IsLinked(selectedLink.GetSchedule()))
                LinkGateway.DeleteLink(selectedLink.GetSchedule());

            _viewModel.Document.Delete(selectedLink.GetSchedule().Id);
            _viewModel.Links.Remove(selectedLink);
            _viewModel.OnPropertyChanged("Links");
        }

        public event EventHandler CanExecuteChanged;
    }
}
