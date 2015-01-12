using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ElectricalToolSuite.MECoordination.UI;
using Settings = ElectricalToolSuite.MECoordination.Properties.Settings;

// TODO Move Workset button to the right

namespace ElectricalToolSuite.MECoordination
{
    [Transaction(TransactionMode.Automatic)]
    public class ExternalCommand : IExternalCommand
    {
        private readonly Settings _defaultSettings;
        private DocumentHelper _document;

        public ExternalCommand()
        {
            _defaultSettings = Settings.Default;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _document = new DocumentHelper(commandData.Application.ActiveUIDocument.Document);

            var mechanicalTree = GetTreeView(_defaultSettings.MechanicalCategories);
            var electricalTree = GetTreeView(_defaultSettings.ElectricalCategories);

            SetInitialCheckedState(mechanicalTree, _defaultSettings.SelectedMechanicalElements);
            SetInitialCheckedState(electricalTree, _defaultSettings.SelectedElectricalElements);

            var userWorksets = _document.UserWorksets;
            var activeWorksetId = _document.ActiveWorksetId;
            var currentWorksetIndex = userWorksets.FindIndex(workset => workset.Id == activeWorksetId);

            var mainWindow = new MainWindow
            {
                MechanicalTree = { ItemsSource = mechanicalTree },
                ElectricalTree = { ItemsSource = electricalTree },
                WorksetComboBox = { ItemsSource = userWorksets, SelectedIndex = currentWorksetIndex }
            };

            bool? accepted = mainWindow.ShowDialog();
            if (accepted.ValueOr(false))
            {
                var selectedWorkset = (Workset)mainWindow.WorksetComboBox.SelectedItem;
                var selectedMechanical = GetSelectedFamilySymbols(mechanicalTree);
                var selectedElectrical = GetSelectedFamilySymbols(electricalTree);

                var mechanicalInstanceIds
                    = _document.GetInstancesOfFamilySymbols(selectedMechanical)
                               .ToList();

                if (ConfirmOperation(mechanicalInstanceIds, selectedElectrical))
                {
                    var coordinator = new MechanicalElectricalEquipmentCoordinator(_document, selectedWorkset);
                    coordinator.Coordinate(mechanicalInstanceIds, selectedElectrical);
                }
            }

            SaveCheckedItems(mechanicalTree, electricalTree);

            return Result.Succeeded;
        }

        private void SaveCheckedItems(IEnumerable<TreeViewItemWithCheckbox> mechanicalItems,
            IEnumerable<TreeViewItemWithCheckbox> electricalItems)
        {
            _defaultSettings.SelectedMechanicalElements = String.Join(",",
                GetSelected(mechanicalItems).Select(i => Convert.ToString(i.IntegerValue)));

            _defaultSettings.SelectedElectricalElements = String.Join(",",
               GetSelected(electricalItems).Select(i => Convert.ToString(i.IntegerValue)));
            
            _defaultSettings.Save();
        }

        private List<TreeViewItemWithCheckbox> GetTreeView(StringCollection names)
        {
            var lookup = new IndexedElementLookup(_document);
            var treeViews = new List<TreeViewItemWithCheckbox>();

            var categoryGroups = _defaultSettings.DebugMode
                ? lookup.FamilyLookup
                : lookup.FamilyLookup.Where(grp => names.Contains(grp.Key.Name));

            foreach (var categoryGroup in categoryGroups.OrderBy(grp => grp.Key.Name))
            {
                var category = categoryGroup.Key;
                var categoryTreeView = new TreeViewItemWithCheckbox(category.Name, category.Id);
                treeViews.Add(categoryTreeView);

                foreach (var family in categoryGroup.OrderBy(f => f.Name))
                {
                    var familyTreeView = new TreeViewItemWithCheckbox(family.Name, family.Id);
                    categoryTreeView.AddChild(familyTreeView);

                    foreach (var symbol in lookup.SymbolLookup[family].OrderBy(s => s.Name))
                    {
                        familyTreeView.AddChild(new FamilySymbolItem(symbol.Name, symbol));
                    }
                }
            }

            return treeViews;
        }

        private void SetInitialCheckedState(IEnumerable<TreeViewItemWithCheckbox> treeView, string checkedItemIds)
        {
            if (String.IsNullOrWhiteSpace(checkedItemIds))
                return;

            var selectedItemIds
                = new HashSet<ElementId>(checkedItemIds.Split(',').Select(Int32.Parse).Select(id => new ElementId(id)));

            foreach (TreeViewItemWithCheckbox item in treeView)
                item.SetInitialCheckedState(selectedItemIds);
        }

        private IEnumerable<ElementId> GetSelected(IEnumerable<TreeViewItemWithCheckbox> treeView)
        {
            return treeView.SelectMany(tv => tv.SelectedWithChildren).Select(item => item.ElementId);
        }

        private List<FamilySymbol> GetSelectedFamilySymbols(IEnumerable<TreeViewItemWithCheckbox> treeView)
        {
            return treeView.SelectMany(tv => tv.SelectedWithChildren)
                .OfType<FamilySymbolItem>()
                .Select(fsi => fsi.FamilySymbol)
                .ToList();
        }

        private bool ConfirmOperation(ICollection<ElementId> mechanicalInstances, ICollection<FamilySymbol> electricalSymbols)
        {
            if (!mechanicalInstances.Any())
            {
                TaskDialog.Show("No instances found", "No instances of the selected families were found.");
                return false;
            }

            var newItemCount = mechanicalInstances.Count() * electricalSymbols.Count();

            var message = String.Format("This operation will create {0} new {1}. Proceed?", newItemCount,
                newItemCount > 1 ? "instances" : "instance");

            const TaskDialogCommonButtons buttons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

            return TaskDialog.Show("Confirm operation", message, buttons) == TaskDialogResult.Yes;
        }
    }
}
