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
            tree.Focus();
        }

        public IEnumerable<ElementTreeViewModel> Roots
        {
            get { return tree.Items.Cast<ElementTreeViewModel>(); }
        }

        public void AddRoot(ElementTreeViewModel root)
        {
            tree.Items.Add(root);

            CommandBindings.Add(
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
        }
    }
}
