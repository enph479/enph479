using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ElectricalToolSuite.MECoordination.UI.Annotations;

namespace ElectricalToolSuite.MECoordination.UI
{
    public class TreeViewItemWithCheckbox : INotifyPropertyChanged
    {
        private bool? _checked = false;
        private TreeViewItemWithCheckbox _parent;

        public TreeViewItemWithCheckbox(string name)
        {
            Name = name;
            Children = new ObservableCollection<TreeViewItemWithCheckbox>();
        }

        public bool? Checked
        {
            get { return _checked; }
            set { SetIsChecked(value, true, true); }
        }
        
        void SetIsChecked(bool? value, bool updateChildren, bool updateParent)
        {
            if (value == _checked)
                return;

            _checked = value;

            if (updateChildren && _checked.HasValue)
            {
                foreach (var child in Children)
                {
                    child.SetIsChecked(_checked, true, false);
                }
            }

            if (updateParent && _parent != null)
                _parent.VerifyCheckState();

            OnPropertyChanged("Checked");
        }

        void VerifyCheckState()
        {
            bool? state = null;
            for (int i = 0; i < Children.Count; ++i)
            {
                bool? current = Children[i].Checked;
                if (i == 0)
                {
                    state = current;
                }
                else if (state != current)
                {
                    state = null;
                    break;
                }
            }
            SetIsChecked(state, false, true);
        }

        public void AddChild(TreeViewItemWithCheckbox newItem)
        {
            Children.Add(newItem);
            newItem._parent = this;
        }

        public string Name { get; set; }
        
        public ObservableCollection<TreeViewItemWithCheckbox> Children { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
