using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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
                    .WherePasses(new ElementIsCurveDrivenFilter(inverted: true)) // Curve-driven elements do not have LocationPoints.
                    .OfClass(typeof (FamilySymbol))
                    .Cast<FamilySymbol>();

            var categories = new ElementCategorizer().GroupByFamilyAndCategoryNames(familySymbols);
            var mechanicalItems = GenerateTreeViewData(categories);
            var electricalItems = GenerateTreeViewData(categories);

            var mainWindow = new MainWindow
            {
                MechanicalTree = {ItemsSource = mechanicalItems},
                ElectricalTree = {ItemsSource = electricalItems}
            };

            var accepted = mainWindow.ShowDialog();

            if (accepted.HasValue && accepted.Value)
            {
                Synchronize(GetSelectedFamilySymbols(mechanicalItems), GetSelectedFamilySymbols(electricalItems));
            }

            return Result.Succeeded;
        }

        private IEnumerable<FamilySymbolItem> GetSelectedFamilySymbols(IEnumerable<TreeViewItemWithCheckbox> treeView)
        {
            return treeView.SelectMany(tv => tv.SelectedWithChildren).OfType<FamilySymbolItem>();
        }

        private void Synchronize(IEnumerable<FamilySymbolItem> mechanicalItems,
            IEnumerable<FamilySymbolItem> electricalItems)
        {
            var mechanicalInstances = GetAllInstances(mechanicalItems);
            var electricalSymbols = electricalItems.Select(fsi => fsi.FamilySymbol).ToList();

            var targetLocationDatas =
                (from mechanicalElement in mechanicalInstances
                let mechanicalInstance = (FamilyInstance) _document.GetElement(mechanicalElement)
                select new
                {
                    ((LocationPoint) mechanicalInstance.Location).Point,
                    mechanicalInstance.Host,
                    mechanicalInstance.FacingOrientation
                }).ToList();

            var familyCreationData =
                from electricalSymbol in electricalSymbols
                from targetLocationData in targetLocationDatas
                select
                    new FamilyInstanceCreationData(targetLocationData.Point, electricalSymbol, StructuralType.NonStructural);                

            var familyCreationDataList = familyCreationData.ToList();
            var newItemCount = familyCreationDataList.Count();

            if (newItemCount == 0)
            {
                TaskDialog.Show("No instances found", "No instances of the selected families were found.");
            }
            else
            {
                var message = String.Format("This operation will create {0} new {1}. Proceed?", newItemCount,
                    newItemCount > 1 ? "instances" : "instance");

                const TaskDialogCommonButtons buttons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

                var confirmationResult = TaskDialog.Show("Confirm operation", message, buttons);

                if (confirmationResult == TaskDialogResult.Yes)
                {
                    var ids = _document.Create.NewFamilyInstances2(familyCreationDataList);
                    
                    foreach (var rotationInformation in ids.Zip(targetLocationDatas, (id, tld) => new {Id = id, tld.Rotation}))
                    {
                        var instance = ((FamilyInstance) _document.GetElement(rotationInformation.Id));
                        var locationPoint = (LocationPoint) instance.Location;
                        var axis = Line.CreateBound(locationPoint.Point,
                            new XYZ(locationPoint.Point.X, locationPoint.Point.Y, locationPoint.Point.Z + 10));
                        instance.Location.Rotate(axis, rotationInformation.Rotation);
                    }
                }
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
    }
}