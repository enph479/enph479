using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ElectricalToolSuite.MECoordination.UI;

namespace ElectricalToolSuite.MECoordination
{
    [Transaction(TransactionMode.Automatic)]
    public class ExternalCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            var collector = new FilteredElementCollector(doc);

            var familySymbols = collector.OfClass(typeof(FamilySymbol))
                                         .Cast<FamilySymbol>()
                                         .Where(fs => fs.Category != null);

            var families = familySymbols.GroupBy(fs => fs.Family.Id);
            var categories = families.GroupBy(g => doc.GetElement(g.First().Id).Category.Name).OrderBy(g => g.Key);

            var categoryItems = new List<CategoryItem>();

            foreach (var categoryGroup in categories)
            {
                var categoryItem = new CategoryItem(categoryGroup.Key);
                foreach (var familyGroup in categoryGroup)
                {
                    var familyItem = new FamilyItem { Family = doc.GetElement(familyGroup.Key) as Family };
                    foreach (var familySymbol in familyGroup)
                    {
                        var familySymbolItem = new FamilySymbolItem { FamilySymbol = familySymbol };
                        familyItem.Symbols.Add(familySymbolItem);
                    }
                    categoryItem.Families.Add(familyItem);
                }
                categoryItems.Add(categoryItem);
            }

            var wnd = new TreeViewMultipleTemplatesSample();
            wnd.Categories.ItemsSource = categoryItems;

            wnd.ShowDialog();

            return Result.Succeeded;
        }
    }
}
