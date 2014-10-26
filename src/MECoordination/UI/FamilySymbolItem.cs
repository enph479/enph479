using Autodesk.Revit.DB;
using System;

namespace ElectricalToolSuite.MECoordination.UI
{
    public class FamilySymbolItem : TreeViewItemWithCheckbox
    {
        public FamilySymbolItem(string name, FamilySymbol familySymbol) : base(name, familySymbol.Id)
        {
            if (familySymbol == null) throw new ArgumentNullException("familySymbol");
            FamilySymbol = familySymbol;
        }

        public FamilySymbol FamilySymbol { get; private set; }
    }
}