using System;
using Excel = NetOffice.ExcelApi;

namespace ElectricalToolSuite.ScheduleImport
{
    public class ExcelSingleton : IDisposable
    {
        private static ExcelSingleton _instance;

        public static void Initialize()
        {
            _instance = new ExcelSingleton();
        }

        public static void Close()
        {
            if (_instance != null)
            {
                Instance.Quit();
                Instance.Dispose();
                _instance = null;
            }
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
            if (_application != null)
            {
                _application.Dispose();
            }
        }
    }
}
