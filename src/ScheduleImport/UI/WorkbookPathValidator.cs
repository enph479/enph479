using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    public class WorkbookPathValidator : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            var path = (string) value;

            if (String.IsNullOrWhiteSpace(path))
                return new ValidationResult(false, null);

            if (!File.Exists(path))
                return new ValidationResult(false, "File does not exist");

            try
            {
                using (ExcelSingleton.OpenWorkbook(path)) ;
            }
            catch (Exception)
            {
                return new ValidationResult(false, "File is not a valid Excel workbook");
            }

            return new ValidationResult(true, null);
        }
    }
}
