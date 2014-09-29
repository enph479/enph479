using System.Collections.ObjectModel;

namespace ElectricalToolSuite.MECoordination.UI
{
    public class CategoryItem
    {
        public CategoryItem(string name)
        {
            Name = name;
            Families = new ObservableCollection<FamilyItem>();
        }

        public string Name { get; set; }

        public bool? Checked { get; set; }

        public ObservableCollection<FamilyItem> Families { get; set; }
    }
}
