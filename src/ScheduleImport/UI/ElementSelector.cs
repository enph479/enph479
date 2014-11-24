using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    class ElementSelector
    {
        private readonly UIDocument _uiDocument;

        public ElementSelector(UIDocument uiDocument)
        {
            _uiDocument = uiDocument;
        }

        public Element SelectSingle()
        {
            var selectedPanelRef = _uiDocument.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
            return _uiDocument.Document.GetElement(selectedPanelRef);
        }
    }
}
