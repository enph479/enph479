using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace ElectricalToolSuite.MECoordination
{
    public class ElementCategorizer
    {
        public IEnumerable<IGrouping<string, IGrouping<string, FamilySymbol>>> GroupByFamilyAndCategoryNames(
            IEnumerable<FamilySymbol> familySymbols)
        {
            // TODO Collect these in dictionaries or some other persistent structure

            familySymbols = familySymbols.Where(fs => fs.Category != null)
                .OrderBy(fs => fs.Name);

            var families = familySymbols.GroupBy(fs => fs.Family.Name)
                .OrderBy(g => g.Key);

            return families.GroupBy(g => g.First().Category.Name)
                .OrderBy(g => g.Key);
        }
    }
}