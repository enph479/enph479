using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ElectricalToolSuite.ExtensionWrapper
{
    [Transaction(TransactionMode.Automatic)]
    public class ExtensionWrapper : IExternalCommand
    {
        // Point this to whichever DLL you want to load
        private static readonly string ExtensionPath = @"C:\git\enph479\src\bin\Debug\MECoordination.dll";

        // Point this to the name of the command class
        private static readonly string ExtensionTypeName = "ElectricalToolSuite.MECoordination.ExternalCommand";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var assemblyBytes = File.ReadAllBytes(ExtensionPath);
                var assembly = Assembly.Load(assemblyBytes);
                var commandType = assembly.GetType(ExtensionTypeName);
                var command = assembly.CreateInstance(ExtensionTypeName);
                var args = new object[] {commandData, message, elements};
                const BindingFlags flags = BindingFlags.Default | BindingFlags.InvokeMethod;
                return (Result) commandType.InvokeMember("Execute", flags, null, command, args);
            }
            catch (TargetInvocationException e)
            {
                TaskDialog.Show("Failed to invoke extension", String.Format("{0} - {1}", e.InnerException.GetType(), e.InnerException.Message));
                return Result.Failed;
            }
            catch (Exception e)
            {
                TaskDialog.Show("Failed to invoke extension", String.Format("{0} - {1}", e.GetType(), e.Message));
                return Result.Failed;
            }
        }
    }
}
