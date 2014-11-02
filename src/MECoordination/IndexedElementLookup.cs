using Autodesk.Revit.DB;
using System.Linq;

namespace ElectricalToolSuite.MECoordination
{
    class IndexedElementLookup
    {
        public ILookup<Category, Family> FamilyLookup { get; private set; }
        public ILookup<Family, FamilySymbol> SymbolLookup { get; private set; }

        public IndexedElementLookup(DocumentAccess document)
        {
            var symbols = document.FamilySymbols.ToList();

            SymbolLookup = symbols.Where(s => s.Family != null)
                .ToLookup(s => s.Family, new FamilyIdEqualityComparer());

            FamilyLookup = symbols.Select(s => s.Family)
                .Where(f => (f.FamilyCategory ?? f.Category) != null)
                .Distinct(new FamilyIdEqualityComparer())
                .ToLookup(f => f.FamilyCategory ?? f.Category, new CategoryIdEqualityComparer());
        }
    }
}
