using System.Collections.ObjectModel;

namespace ElectricalToolSuite.MECoordination.UI
{
    public class FamilySymbolItem
    {
        public FamilySymbolItem()
        {
            Instances = new ObservableCollection<FamilyInstanceItem>();
        }

        public string Name { get { return FamilySymbol.Name; } }

        public bool Checked { get; set; }

        public Autodesk.Revit.DB.FamilySymbol FamilySymbol { get; set; }

        public ObservableCollection<FamilyInstanceItem> Instances { get; set; }
    }
}
