using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
namespace ElectricalToolSuite.ScheduleImport.UI
{
    public class ScheduleNameValidator : ValidationRule
    {
        public object RevitDocument { get; set; }

        public string ExemptedName { get; set; }

        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            if (RevitDocument == null)
                return new ValidationResult(false, null);

            var name = (string) value;

            if (String.IsNullOrWhiteSpace(name))
                return new ValidationResult(false, null);

            if (!String.IsNullOrWhiteSpace(ExemptedName) && ExemptedName == name)
                return new ValidationResult(true, null);

            if (new FilteredElementCollector((Document) RevitDocument).OfClass(typeof (PanelScheduleView))
                        .Cast<PanelScheduleView>().Any(s => s.Name == name))
                return new ValidationResult(false, "There is already a panel schedule with this name");

            return new ValidationResult(true, null);
        }
    }
}
