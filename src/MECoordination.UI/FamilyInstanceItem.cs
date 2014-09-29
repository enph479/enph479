namespace ElectricalToolSuite.MECoordination.UI
{
    public class FamilyInstanceItem
    {
        public Autodesk.Revit.DB.FamilyInstance FamilyInstance { get; set; }

        public string Name { get { return FamilyInstance.Name; } }
    }
}
