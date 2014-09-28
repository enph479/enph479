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
//            var wnd = new UI.Window1();
//            wnd.ShowDialog();

            var app = commandData.Application;
            var doc = app.ActiveUIDocument.Document;

//            var categories = doc.Settings.Categories;

//            var builtInCategories = Enum.GetValues(typeof (BuiltInCategory));

            

//            var builtInCategoryNames = new List<string>();

//            foreach (BuiltInCategory builtInCategory in builtInCategories)
//            {
//                var category = categories.get_Item(builtInCategory);
//                builtInCategoryNames.Add(category.Name);
//            }

//            var mechanicalFilter = new ElementCategoryFilter(BuiltInCategory.OST_MechanicalEquipment | BuiltInCategory.OST_DuctFitting | BuiltInCategory.OST_FlexPipeCurves);
            var collector = new FilteredElementCollector(doc);

            var byFamily = collector.WherePasses(new ElementClassFilter(typeof (FamilySymbol)))
                .Cast<FamilySymbol>()
                .GroupBy(e => e.Family.Name);

            var byCategory = byFamily.GroupBy(g => g.First().Category.Name);

//            var families = collector.WherePasses(new ElementClassFilter(typeof(FamilySymbol)))
//    .Cast<FamilySymbol>()
//    .GroupBy(e => e.Category == null ? "<null>" : e.Category.Name)
//    .ToDictionary(e => e.Key, e => e.ToList());

            var wnd = new Window1();

            foreach (var categoryGroup in byCategory)
            {
                var categoryItem = new ElementTreeViewModel {Name = categoryGroup.Key};
                wnd.AddRoot(categoryItem);
                foreach (var familyGroup in categoryGroup)
                {
                    var familyItem = new ElementTreeViewModel {Name = familyGroup.Key};
                    categoryItem.AddChild(familyItem);
                    foreach (var familySymbol in familyGroup)
                    {
                        familyItem.AddChild(new ElementTreeViewModel {Name = familySymbol.Name, Element = familySymbol});
                    }
                }

//                Debug.Print("{0}:", categoryGroup.Key);
//                foreach (var familyGroup in categoryGroup)
//                {
//                    Debug.Print("\t{0}", familyGroup.Key);
//                    foreach (var familySymbol in familyGroup)
//                    {
//                        Debug.Print("\t\t{0}", familySymbol.Name);
//                    }
//                }
            }

            wnd.ShowDialog();

            foreach (var familySymbol in wnd.Roots.SelectMany(r => r.SelectedItems))
            {
                Debug.Print(familySymbol.Name);
            }

//            var mechanicalElements = collector.WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctFitting))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctInsulations))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctLinings))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctSystem))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctTerminal))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_FlexDuctCurves))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory))
//                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory));



//            var disciplines = ViewDiscipline

//            var sb = new StringBuilder();

//            sb.AppendLine("You selected:");
//            foreach (var item in selected)
//            {
//                sb.AppendLine(item);
//            }

//            TaskDialog.Show("Selected items", sb.ToString());

//            var categories =
//                familyInstances.Select(fi => fi == null ? null : fi.Category)
//                    .Where(c => c != null)
//                    .Distinct();

//            var wnd = new UI.Window1();
//            var root = wnd.Root;

//            foreach (var category in categories)
//            {
//            }
            
            
//            wnd.ShowDialog();

//            var sb = new StringBuilder();
//            sb.AppendLine("Categories with family instances:");
//
//            foreach (var name in categoryNames)
//                sb.AppendLine(name);
//
//            TaskDialog.Show("Category names", sb.ToString());

            return Result.Succeeded;
        }
    }
}
