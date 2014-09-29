using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using ElectricalToolSuite.MECoordination.UI.Annotations;

namespace ElectricalToolSuite.MECoordination.UI
{
    public class CategoryItem : INotifyPropertyChanged
    {
        private bool? _checked;

        public CategoryItem(string name)
        {
            Name = name;
            Families = new ObservableCollection<FamilyItem>();
        }

        public string Name { get; set; }

        public bool? Checked
        {
            get
            {

            }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("value");

                var nonNullChecked = value.Value;
                foreach (var child in Families.Where(s => s.Checked != nonNullChecked))
                {
                    child.Checked = nonNullChecked;
                }

                OnPropertyChanged("Checked");
            }
        }

        private void UpdateIsChecked()
        {
            if (!Families.Any())
            {
                _checked = false;
                return;
            }

            bool? newState;
            var anyChildrenIndeterminate = Families.Any(s => !s.Checked.HasValue);
            var anyChildrenSelected = Families.Any(s => s.Checked.HasValue && s.Checked.Value);
            var anyChildrenUnselected = Families.Any(s => s.Checked.HasValue && !s.Checked.Value);

            if (anyChildrenIndeterminate || (anyChildrenSelected && anyChildrenUnselected))
                newState = null;
            else
                newState = anyChildrenSelected;

            _checked = newState;
        }
        
        public ObservableCollection<FamilyItem> Families { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
