using System;
using System.ComponentModel;
using System.Linq;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using ElectricalToolSuite.MECoordination.UI.Annotations;

namespace ElectricalToolSuite.MECoordination.UI
{
    public class FamilyItem : INotifyPropertyChanged
    {
        public FamilyItem()
        {
            Symbols = new ObservableCollection<FamilySymbolItem>();
        }

        public string Name { get { return Family.Name; } }

        [NotifyParentProperty(true)]
        public bool? Checked 
        {
            get
            {
                if (!Symbols.Any())
                    return false;

                var anyChildrenSelected = Symbols.Any(s => s.Checked);
                var anyChildrenUnselected = Symbols.Any(s => !s.Checked);

                if (anyChildrenSelected && anyChildrenUnselected)
                    return null;

                return anyChildrenSelected;
            }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("value");

                var nonNullChecked = value.Value;
                foreach (var child in Symbols.Where(s => s.Checked != nonNullChecked))
                {
                    child.Checked = nonNullChecked;
                }

                OnPropertyChanged("Checked");
            }
        }

        public Autodesk.Revit.DB.Family Family { get; set; }

        public ObservableCollection<FamilySymbolItem> Symbols { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
