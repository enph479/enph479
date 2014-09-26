using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using MECoordination.UI;

namespace ElectricalToolSuite.MECoordination.UI
{
    public partial class Window1 : Window
    {
        public Window1()
        {
            InitializeComponent();

            SelectableTreeViewModel root = this.tree.Items[0] as SelectableTreeViewModel;

            base.CommandBindings.Add(
                new CommandBinding(
                    ApplicationCommands.Undo,
                    (sender, e) => // Execute
                    {
                        e.Handled = true;
                        root.IsChecked = false;
                        this.tree.Focus();
                    },
                    (sender, e) => // CanExecute
                    {
                        e.Handled = true;
                        e.CanExecute = (root.IsChecked != false);
                    }));

            this.tree.Focus();
        }

        public IEnumerable<string> SelectedItems
        {
            get
            {
                var s = new Stack<SelectableTreeViewModel>();

                foreach (var item in tree.Items)
                    s.Push((SelectableTreeViewModel) item);

                while (s.Any())
                {
                    var vm = s.Pop();

                    if (vm.Children.Any())
                    {
                        foreach (var c in vm.Children)
                            s.Push(c);
                        continue;
                    }

                    if ((vm.IsChecked.HasValue && vm.IsChecked.Value) || (!vm.IsChecked.HasValue && vm.IsInitiallySelected))
                        yield return vm.Name;
                }
            }
        }
    }
}
