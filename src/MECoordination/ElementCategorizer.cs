using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace ElectricalToolSuite.MECoordination
{
    public class ElementCategorizer
    {
        public IEnumerable<IGrouping<Category, IGrouping<Family, FamilySymbol>>> GroupByFamilyAndCategory(IEnumerable<FamilySymbol> familySymbols)
        {
            familySymbols = familySymbols.Where(fs => fs.Category != null)
                .OrderBy(fs => fs.Name);

            var families = familySymbols.GroupBy(fs => fs.Family, new ElementNameEqualityComparer<Family>())
                .OrderBy(g => g.Key.Name);

            return families.GroupBy(g => g.Key.FamilyCategory, new CategoryNameEqualityComparer())
                .OrderBy(g => g.Key.Name);
        }

        public IEnumerable<IGrouping<string, IGrouping<string, FamilySymbol>>> GroupByFamilyAndCategoryNames(IEnumerable<FamilySymbol> familySymbols)
        {
            familySymbols = familySymbols.Where(fs => fs.Category != null)
                .OrderBy(fs => fs.Name);

            var families = familySymbols.GroupBy(fs => fs.Family.Name)
                .OrderBy(g => g.Key);

            return families.GroupBy(g => g.First().Category.Name)
                .OrderBy(g => g.Key);
        }
    }
}
