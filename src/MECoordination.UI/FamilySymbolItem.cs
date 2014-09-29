using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ElectricalToolSuite.MECoordination.UI.Annotations;

namespace ElectricalToolSuite.MECoordination.UI
{
    public class FamilySymbolItem : INotifyPropertyChanged
    {
        private bool _checked;

        public FamilySymbolItem()
        {
            Instances = new ObservableCollection<FamilyInstanceItem>();
        }

        public string Name { get { return FamilySymbol.Name; } }

        [NotifyParentProperty(true)]
        public bool Checked
        {
            get { return _checked; }
            set
            {
                if (value.Equals(_checked)) return;
                _checked = value;
                OnPropertyChanged();
            }
        }

        public Autodesk.Revit.DB.FamilySymbol FamilySymbol { get; set; }

        public ObservableCollection<FamilyInstanceItem> Instances { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
