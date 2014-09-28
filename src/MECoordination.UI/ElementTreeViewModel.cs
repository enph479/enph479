using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.MECoordination.UI
{
    public class ElementTreeViewModel : INotifyPropertyChanged
    {
        bool? _isChecked = false;
        private ElementTreeViewModel _parent;
        private List<ElementTreeViewModel> _children = new List<ElementTreeViewModel>();

        public Element Element { get; set; }

        public string Name { get; set; }

        public bool IsInitiallySelected { get; set; }

        public List<ElementTreeViewModel> Children
        {
            get { return _children; }
            set { _children = value; }
        }
        
        public void AddChild(ElementTreeViewModel existing)
        {
            Children.Add(existing);
        }

        /// <summary>
        /// Gets/sets the state of the associated UI toggle (ex. CheckBox).
        /// The return value is calculated based on the check state of all
        /// children.  Setting this property to true or false
        /// will set all children to the same check state, and setting it 
        /// to any value will cause the parent to verify its check state.
        /// </summary>
        public bool? IsChecked
        {
            get { return _isChecked; }
            set { SetIsChecked(value, true, true); }
        }

        public ElementTreeViewModel Parent
        {
            get { return _parent; }
        }

        void SetIsChecked(bool? value, bool updateChildren, bool updateParent)
        {
            if (value == _isChecked)
                return;

            _isChecked = value;

            if (updateChildren && _isChecked.HasValue)
                Children.ForEach(c => c.SetIsChecked(_isChecked, true, false));

            if (updateParent && Parent != null)
                Parent.VerifyCheckState();

            OnPropertyChanged("IsChecked");
        }

        void VerifyCheckState()
        {
            bool? state = null;
            for (int i = 0; i < Children.Count; ++i)
            {
                bool? current = Children[i].IsChecked;
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

        void OnPropertyChanged(string prop)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }

        public IEnumerable<Element> SelectedItems
        {
            get
            {
                foreach (var child in Children)
                {
                    foreach (var item in child.SelectedItems)
                    {
                        yield return item;
                    }
                }

                if (IsChecked.HasValue && IsChecked.Value && Element != null)
                    yield return Element;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
