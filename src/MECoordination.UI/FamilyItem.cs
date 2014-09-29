using System.Collections.ObjectModel;

namespace ElectricalToolSuite.MECoordination.UI
{
    public class FamilyItem
    {
        public FamilyItem()
        {
            Symbols = new ObservableCollection<FamilySymbolItem>();
        }

        public string Name { get { return Family.Name; } }

        public bool? Checked { get; set; }

        public Autodesk.Revit.DB.Family Family { get; set; }

        public ObservableCollection<FamilySymbolItem> Symbols { get; set; }
    }
}
