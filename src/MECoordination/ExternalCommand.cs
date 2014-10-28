using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using ElectricalToolSuite.MECoordination.UI;
using RevitExceptions = Autodesk.Revit.Exceptions;
using Settings = ElectricalToolSuite.MECoordination.Properties.Settings;

namespace ElectricalToolSuite.MECoordination
{
    [Transaction(TransactionMode.Automatic)]
    public class ExternalCommand : IExternalCommand
    {
        private readonly Settings _defaultSettings;
        private Document _document;

        public ExternalCommand()
        {
            _defaultSettings = Settings.Default;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _document = commandData.Application.ActiveUIDocument.Document;

            var familySymbols =
                new FilteredElementCollector(_document)
                    .WherePasses(new ElementIsCurveDrivenFilter(true))
                    // Curve-driven elements do not have LocationPoints.
                    .OfClass(typeof (FamilySymbol))
                    .Cast<FamilySymbol>();

            var categories =
                new ElementCategorizer().GroupByFamilyAndCategory(familySymbols);
            var mechanicalItems = GetFilteredTreeView(categories,
                _defaultSettings.MechanicalCategories);
            var electricalItems = GetFilteredTreeView(categories,
                _defaultSettings.ElectricalCategories);

            if (!TrySetInitialCheckedState(mechanicalItems, _defaultSettings.SelectedMechanicalElements) ||
                !TrySetInitialCheckedState(electricalItems, _defaultSettings.SelectedElectricalElements))
                _defaultSettings.Reset();

            var userWorksets = GetAllUserWorksets();
            var activeWorksetId = _document.GetWorksetTable().GetActiveWorksetId();
            var currentWorksetIndex = userWorksets.FindIndex(workset => workset.Id == activeWorksetId);

            var mainWindow = new MainWindow(commandData.Application)
            {
                MechanicalTree = { ItemsSource = mechanicalItems },
                ElectricalTree = { ItemsSource = electricalItems },
                WorksetComboBox = { ItemsSource = userWorksets, SelectedIndex = currentWorksetIndex }
            };

            bool? accepted = mainWindow.ShowDialog();

            if (accepted.ValueOr(false))
            {
                bool tagOnPlacement = mainWindow.TagOnPlacementCheckBox.IsChecked.ValueOr(false);
                var selectedWorkset = (Workset) mainWindow.WorksetComboBox.SelectedItem;

                Synchronize(GetSelectedFamilySymbols(mechanicalItems), GetSelectedFamilySymbols(electricalItems),
                    tagOnPlacement, selectedWorkset);
            }

            SaveCheckedItems(mechanicalItems, electricalItems);

            return Result.Succeeded;
        }

        private List<Workset> GetAllUserWorksets()
        {
            return new FilteredWorksetCollector(_document)
                .OfKind(WorksetKind.UserWorkset)
                .ToList();
        }

        private void SaveCheckedItems(IEnumerable<TreeViewItemWithCheckbox> mechanicalItems,
            IEnumerable<TreeViewItemWithCheckbox> electricalItems)
        {
            _defaultSettings.SelectedMechanicalElements = String.Join(",",
                GetSelected(mechanicalItems).Select(i => Convert.ToString(i.IntegerValue)));

            _defaultSettings.SelectedElectricalElements = String.Join(",",
                GetSelected(electricalItems).Select(i => Convert.ToString(i.IntegerValue)));

            // Ensure the result is parsable
            if (TrySetInitialCheckedState(Enumerable.Empty<TreeViewItemWithCheckbox>(), _defaultSettings.SelectedMechanicalElements)
                && TrySetInitialCheckedState(Enumerable.Empty<TreeViewItemWithCheckbox>(), _defaultSettings.SelectedElectricalElements))
                _defaultSettings.Save();
        }

        private List<TreeViewItemWithCheckbox> GetFilteredTreeView(
            Dictionary<Category, Dictionary<ElementId, HashSet<ElementId>>> categories, StringCollection names)
        {
            var filteredCategories =
                categories.Where(kvp => names.Contains(kvp.Key.Name));
            return GetTreeView(_defaultSettings.DebugMode ? categories : filteredCategories);
        }

        private List<View> GetAllViews()
        {
            return
                new FilteredElementCollector(_document)
                    .OfClass(typeof (View))
                    .Cast<View>()
                    .OrderBy(view => view.ViewName)
                    .ToList();
        }

        private bool TrySetInitialCheckedState(IEnumerable<TreeViewItemWithCheckbox> treeView, string checkedItemIds)
        {
            try
            {
                var selectedItemIds =
                    new HashSet<ElementId>(checkedItemIds.Split(',').Select(Int32.Parse).Select(id => new ElementId(id)));
                if (!selectedItemIds.Any())
                    return true;
                foreach (TreeViewItemWithCheckbox item in treeView)
                    item.SetInitialCheckedState(selectedItemIds);
            }
            catch (FormatException)
            {
                return false;
            }

            return true;
        }

        private IEnumerable<ElementId> GetSelected(IEnumerable<TreeViewItemWithCheckbox> treeView)
        {
            return treeView.SelectMany(tv => tv.SelectedWithChildren).Select(item => item.ElementId);
        }

        private IEnumerable<FamilySymbolItem> GetSelectedFamilySymbols(IEnumerable<TreeViewItemWithCheckbox> treeView)
        {
            return treeView.SelectMany(tv => tv.SelectedWithChildren).OfType<FamilySymbolItem>();
        }

        private void Synchronize(IEnumerable<FamilySymbolItem> mechanicalItems,
            IEnumerable<FamilySymbolItem> electricalItems, bool tagOnPlacement, Workset workset)
        {
            var mechanicalInstanceIds = GetAllInstances(mechanicalItems).ToList();
            var electricalSymbols = electricalItems.Select(fsi => fsi.FamilySymbol).ToList();
            int newItemCount = mechanicalInstanceIds.Count()*electricalSymbols.Count;

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
                    CreateInstances(tagOnPlacement, mechanicalInstanceIds, electricalSymbols, workset);
                }
            }
        }

        private void CreateInstances(bool tagOnPlacement, IEnumerable<ElementId> mechanicalInstanceIds,
            IEnumerable<FamilySymbol> electricalSymbols, Workset workset)
        {
            var targetLocations =
                (from mechanicalElement in mechanicalInstanceIds
                 let mechanicalInstance = (FamilyInstance) _document.GetElement(mechanicalElement)
                 select new
                 {
                     ((LocationPoint) mechanicalInstance.Location).Point,
                     mechanicalInstance.Host,
                 }).ToList();

            var newInstanceLocations = new List<Tuple<FamilyInstance, XYZ>>();
            foreach (var electricalSymbol in electricalSymbols)
            {
                foreach (var location in targetLocations)
                {
                    FamilyInstance newInstance;

                    if (location.Host != null)
                        newInstance = _document.Create.NewFamilyInstance(location.Point, electricalSymbol,
                            location.Host,
                            StructuralType.NonStructural);
                    else
                        newInstance = _document.Create.NewFamilyInstance(location.Point, electricalSymbol,
                            StructuralType.NonStructural);

                    var worksetParameter = newInstance.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                    worksetParameter.Set(workset.Id.IntegerValue);

                    newInstanceLocations.Add(Tuple.Create(newInstance, location.Point));
                }
            }

//            if (tagOnPlacement)
//            {
//                CreateTags(newInstanceLocations, tagView);
//            }
        }

        private void CreateTags(IEnumerable<Tuple<FamilyInstance, XYZ>> instanceLocations, View view)
        {
            string failureMessage = "";
            bool addLeader = _defaultSettings.UseTagLeaders;

            foreach (var tuple in instanceLocations)
            {
                FamilyInstance instance = tuple.Item1;
                XYZ location = tuple.Item2;

                try
                {
                    _document.Create.NewTag(view, instance, addLeader, TagMode.TM_ADDBY_CATEGORY,
                        TagOrientation.Horizontal, location);
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

        private List<TreeViewItemWithCheckbox> GetTreeView(
            Dictionary<Category, Dictionary<ElementId, HashSet<ElementId>>> categories)
        {
            var categoryTreeViewItems = new List<TreeViewItemWithCheckbox>();

            foreach (Category category in categories.Keys.OrderBy(c => c.Name))
            {
                var categoryItem =
                    new TreeViewItemWithCheckbox(category.Name, category.Id);
                foreach (ElementId family in categories[category].Keys.OrderBy(GetElementName))
                {
                    var familyItem = new TreeViewItemWithCheckbox(_document.GetElement(family).Name, family);
                    foreach (ElementId familySymbol in categories[category][family].OrderBy(GetElementName))
                    {
                        var familySymbolItem = new FamilySymbolItem(_document.GetElement(familySymbol).Name,
                            (FamilySymbol) _document.GetElement(familySymbol));
                        familyItem.AddChild(familySymbolItem);
                    }
                    categoryItem.AddChild(familyItem);
                }
                categoryTreeViewItems.Add(categoryItem);
            }

            return categoryTreeViewItems;
        }

        private string GetElementName(ElementId id)
        {
            return _document.GetElement(id).Name;
        }
    }
}