using System.Collections;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace ElectricalToolSuite.MECoordination.UI
{
    public partial class TreeViewMultipleTemplatesSample : Window
    {
        public TreeViewMultipleTemplatesSample()
        {
            InitializeComponent();

//            List<CategoryItem> categories = new List<CategoryItem>();

//            CategoryItem category1 = new CategoryItem() { Name = "The Doe's" };
//            family1.Symbols.Add(new FamilySymbol() { Name = "John Doe", Age = 42 });
//            family1.Symbols.Add(new FamilySymbol() { Name = "Jane Doe", Age = 39 });
//            family1.Symbols.Add(new FamilySymbol() { Name = "Sammy Doe", Age = 13 });
//            families.Add(family1);

//            FamilyItem family2 = new FamilyItem() { Name = "The Moe's" };
//            family2.Symbols.Add(new FamilySymbol() { Name = "Mark Moe", Age = 31 });
//            family2.Symbols.Add(new FamilySymbol() { Name = "Norma Moe", Age = 28 });
//            families.Add(family2);

//            trvFamilies.ItemsSource = families;
        }

        public TreeView Categories { get { return trvFamilies; } }
    }
}
