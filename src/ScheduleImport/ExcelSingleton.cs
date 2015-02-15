using System;
using Excel = NetOffice.ExcelApi;

namespace ElectricalToolSuite.ScheduleImport
{
    public class ExcelSingleton : IDisposable
    {
        private static readonly ExcelSingleton _instance;

        static ExcelSingleton()
        {
            _instance = new ExcelSingleton();
        }

        public static Excel.Application Instance { get { return _instance._application; } }

        public static Excel.Workbook OpenWorkbook(string path)
        {
            return Instance.Workbooks.Open(path, updateLinks: false, readOnly: true);
        }

        private readonly Excel.Application _application;

        public ExcelSingleton()
        {
            _application = new Excel.Application { DisplayAlerts = false };
        }

        public void Dispose()
        {
            _application.Dispose();
        }
    }
}
