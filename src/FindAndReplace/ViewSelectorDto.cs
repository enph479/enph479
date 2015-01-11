using Autodesk.Revit.DB;

namespace ElectricalToolSuite.FindAndReplace
{
    public class ViewSelectorDto
    {
        public bool IsChecked { get; set; }
        public View View { get; set; }

        public ViewSelectorDto(bool isChecked, View view)
        {
            IsChecked = isChecked;
            View = view;
        }
    }
}
