using System.Collections.Generic;
using System.ComponentModel;
using ElectricalToolSuite.MECoordination.UI.Annotations;

namespace ElectricalToolSuite.MECoordination.UI
{
    class SelectableTreeViewModel : INotifyPropertyChanged
    {
        #region Data

        bool? _isChecked = false;
        SelectableTreeViewModel _parent;

        #endregion // Data

        #region CreateFoos

        public static List<SelectableTreeViewModel> CreateFoos()
        {
            SelectableTreeViewModel root = new SelectableTreeViewModel("Weapons")
            {
                IsInitiallySelected = true,
                Children =
                {
                    new SelectableTreeViewModel("Blades")
                    {
                        Children =
                        {
                            new SelectableTreeViewModel("Dagger"),
                            new SelectableTreeViewModel("Machete"),
                            new SelectableTreeViewModel("Sword"),
                        }
                    },
                    new SelectableTreeViewModel("Vehicles")
                    {
                        Children =
                        {
                            new SelectableTreeViewModel("Apache Helicopter"),
                            new SelectableTreeViewModel("Submarine"),
                            new SelectableTreeViewModel("Tank"),                            
                        }
                    },
                    new SelectableTreeViewModel("Guns")
                    {
                        Children =
                        {
                            new SelectableTreeViewModel("AK 47"),
                            new SelectableTreeViewModel("Beretta"),
                            new SelectableTreeViewModel("Uzi"),
                        }
                    },
                }
            };

            root.Initialize();
            return new List<SelectableTreeViewModel> { root };
        }

        SelectableTreeViewModel(string name)
        {
            Name = name;
            Children = new List<SelectableTreeViewModel>();
        }

        void Initialize()
        {
            foreach (SelectableTreeViewModel child in Children)
            {
                child._parent = this;
                child.Initialize();
            }
        }

        #endregion // CreateFoos

        #region Properties

        public List<SelectableTreeViewModel> Children { get; private set; }

        public bool IsInitiallySelected { get; set; }

        public string Name { get; set; }

        #region IsChecked

        /// <summary>
        /// Gets/sets the state of the associated UI toggle (ex. CheckBox).
        /// The return value is calculated based on the check state of all
        /// child FooViewModels.  Setting this property to true or false
        /// will set all children to the same check state, and setting it 
        /// to any value will cause the parent to verify its check state.
        /// </summary>
        public bool? IsChecked
        {
            get { return _isChecked; }
            set { SetIsChecked(value, true, true); }
        }

        void SetIsChecked(bool? value, bool updateChildren, bool updateParent)
        {
            if (value == _isChecked)
                return;

            _isChecked = value;

            if (updateChildren && _isChecked.HasValue)
                Children.ForEach(c => c.SetIsChecked(_isChecked, true, false));

            if (updateParent && _parent != null)
                _parent.VerifyCheckState();

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

        #endregion // IsChecked

        #endregion // Properties

        #region INotifyPropertyChanged Members

        void OnPropertyChanged(string prop)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
}
