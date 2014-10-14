using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using RevitExceptions = Autodesk.Revit.Exceptions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ElectricalToolSuite.MECoordination.UI;
using Document = Autodesk.Revit.DB.Document;
using Autodesk.Revit.DB.Structure;

namespace ElectricalToolSuite.MECoordination
{
    [Transaction(TransactionMode.Automatic)]
    public class ExternalCommand : IExternalCommand
    {
        private Document _document;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _document = commandData.Application.ActiveUIDocument.Document;

            var familySymbols =
                new FilteredElementCollector(_document)
                    .WherePasses(new ElementIsCurveDrivenFilter(inverted: true))
                    // Curve-driven elements do not have LocationPoints.
                    .OfClass(typeof (FamilySymbol))
                    .Cast<FamilySymbol>();

            var categories = new ElementCategorizer().GroupByFamilyAndCategoryNames(familySymbols).ToList();
            var mechanicalItems = GenerateTreeViewData(categories);
            var electricalItems = GenerateTreeViewData(categories);

            var mainWindow = new MainWindow(commandData.Application)
            {
                MechanicalTree = {ItemsSource = mechanicalItems},
                ElectricalTree = {ItemsSource = electricalItems}
            };

            var accepted = mainWindow.ShowDialog();

            if (accepted.HasValue && accepted.Value)
            {
                bool tagOnPlacement = mainWindow.TagOnPlacementCheckBox.IsChecked.ValueOr(false);
                Synchronize(GetSelectedFamilySymbols(mechanicalItems), GetSelectedFamilySymbols(electricalItems),
                    tagOnPlacement);
            }

            return Result.Succeeded;
        }

        public Result _Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _document = commandData.Application.ActiveUIDocument.Document;

            var categories = _document.Settings.Categories;

            var tagBics = typeof (BuiltInCategory).GetEnumNames()
                .Where(name => name.EndsWith("Tags"))
                .Select(name => (BuiltInCategory) Enum.Parse(typeof (BuiltInCategory), name));

            var tagFamilySymbols =
                tagBics.SelectMany(bic => GetElementsOfType(_document, typeof (FamilySymbol), bic))
                    .Cast<FamilySymbol>()
                    .ToList();

            File.WriteAllLines(@"C:\git\enph479\tag_family_symbols.txt", tagFamilySymbols.Select(tfs => String.Format("{0} - {1}", tfs.Family.Name, tfs.Name)));

//            new FilteredElementCollector(_document).Ofcl
//
//            var tagCategories =
//                new FilteredElementCollector(_document)
//                    .OfClass(typeof (IndependentTag))
//                    .Cast<IndependentTag>()
//                    .Select(t => t.Category)
//                    .Distinct(new CategoryNameEqualityComparer());
//
//            File.WriteAllLines(@"C:\git\enph479\tag_categories.txt", tagCategories.Select(tc => tc.Name));

            commandData.Application.PostCommand(RevitCommandId.LookupPostableCommandId(PostableCommand.LoadedTags));

            return Result.Succeeded;
        }

        private IEnumerable<FamilySymbolItem> GetSelectedFamilySymbols(IEnumerable<TreeViewItemWithCheckbox> treeView)
        {
            return treeView.SelectMany(tv => tv.SelectedWithChildren).OfType<FamilySymbolItem>();
        }

        private void Synchronize(IEnumerable<FamilySymbolItem> mechanicalItems,
            IEnumerable<FamilySymbolItem> electricalItems, bool tagOnPlacement)
        {
            var familyCreationData = BuildFamilyCreationData(mechanicalItems, electricalItems);
            var newItemCount = familyCreationData.Count();

            if (newItemCount == 0)
            {
                TaskDialog.Show("No instances found", "No instances of the selected families were found.");
            }
            else
            {
                CreateInstances(tagOnPlacement, newItemCount, familyCreationData);
            }
        }

        private List<FamilyInstanceCreationData> BuildFamilyCreationData(IEnumerable<FamilySymbolItem> mechanicalItems,
            IEnumerable<FamilySymbolItem> electricalItems)
        {
            var mechanicalInstances = GetAllInstances(mechanicalItems);
            var electricalSymbols = electricalItems.Select(fsi => fsi.FamilySymbol).ToList();

            var targetLocationDatas =
                from mechanicalElement in mechanicalInstances
                let mechanicalInstance = (FamilyInstance) _document.GetElement(mechanicalElement)
                select new
                {
                    ((LocationPoint) mechanicalInstance.Location).Point,
                    mechanicalInstance.Host,
                    mechanicalInstance.FacingOrientation
                };

            var familyCreationData =
                from electricalSymbol in electricalSymbols
                from targetLocationData in targetLocationDatas
                select
                    new FamilyInstanceCreationData(targetLocationData.Point, electricalSymbol, targetLocationData.Host,
                        StructuralType.NonStructural);

            return familyCreationData.ToList();
        }

        private void CreateInstances(bool tagOnPlacement, int newItemCount,
            List<FamilyInstanceCreationData> familyCreationDataList)
        {
            var message = String.Format("This operation will create {0} new {1}. Proceed?", newItemCount,
                newItemCount > 1 ? "instances" : "instance");

            const TaskDialogCommonButtons buttons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

            var confirmationResult = TaskDialog.Show("Confirm operation", message, buttons);

            if (confirmationResult == TaskDialogResult.Yes)
            {
                var instanceIds = _document.Create.NewFamilyInstances2(familyCreationDataList);

                if (tagOnPlacement)
                {
                    CreateTags(instanceIds);
                }
            }
        }

        private void CreateTags(ICollection<ElementId> instanceIds)
        {
            string failureMessage = "";

            var currentView = _document.ActiveView;
            foreach (var instanceId in instanceIds)
            {
                var instance = (FamilyInstance) _document.GetElement(instanceId);
                var locationPoint = (LocationPoint) instance.Location;

                try
                {
                    const bool addLeader = false;
                    _document.Create.NewTag(currentView, instance, addLeader, TagMode.TM_ADDBY_CATEGORY,
                        TagOrientation.Horizontal, locationPoint.Point);
                }
                catch (RevitExceptions.InvalidOperationException)
                {
                    failureMessage = String.Format("There is no tag available for family type \"{0} - {1}\"",
                        instance.Symbol.Family.Name, instance.Symbol.Name);
                }
            }

            if (!String.IsNullOrEmpty(failureMessage))
            {
                TaskDialog.Show("Failed to create tags",
                    String.Format("One or more tags could not be created. \nError message: \n{0}", failureMessage));
            }
        }

        private IEnumerable<ElementId> GetAllInstances(IEnumerable<FamilySymbolItem> symbols)
        {
            var filters =
                symbols.Select(fsi => new FamilyInstanceFilter(_document, fsi.FamilySymbol.Id))
                    .Cast<ElementFilter>()
                    .ToList();
            var unionFilter = new LogicalOrFilter(filters);
            return new FilteredElementCollector(_document).WherePasses(unionFilter).ToElementIds();
        }

        private static List<TreeViewItemWithCheckbox> GenerateTreeViewData(
            IEnumerable<IGrouping<string, IGrouping<string, FamilySymbol>>> categories)
        {
            var categoryTreeViewItems = new List<TreeViewItemWithCheckbox>();

            foreach (var categoryGroup in categories)
            {
                var categoryItem = new TreeViewItemWithCheckbox(categoryGroup.Key);
                foreach (var familyGroup in categoryGroup)
                {
                    var familyItem = new TreeViewItemWithCheckbox(familyGroup.Key);
                    foreach (var familySymbol in familyGroup)
                    {
                        var familySymbolItem = new FamilySymbolItem(familySymbol.Name, familySymbol);
                        familyItem.AddChild(familySymbolItem);
                    }
                    categoryItem.AddChild(familyItem);
                }
                categoryTreeViewItems.Add(categoryItem);
            }
            return categoryTreeViewItems;
        }

        private static FilteredElementCollector GetElementsOfType(Document doc, Type type, BuiltInCategory bic)
        {
            var collector = new FilteredElementCollector(doc);

            collector.OfCategory(bic);
            collector.OfClass(type);

            return collector;
        }

        private static FilteredElementCollector GetFamilySymbols(Document doc, BuiltInCategory bic)
        {
            return GetElementsOfType(doc, typeof (FamilySymbol), bic);
        }

        private static FamilySymbol GetFirstFamilySymbol(Document doc, BuiltInCategory bic)
        {
            var s = GetFamilySymbols(doc, bic).FirstElement() as FamilySymbol;

            Debug.Assert(null != s, string.Format("expected at least one {0} symbol in project", bic.ToString()));

            return s;
        }
    }
}