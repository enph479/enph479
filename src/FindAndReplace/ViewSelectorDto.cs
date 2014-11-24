using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
