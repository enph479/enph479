using System;
using System.Configuration;
using System.IO;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ElectricalToolSuite.ExtensionLoader
{
    [Transaction(TransactionMode.Automatic)]
    public class ExtensionLoader : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var thisAssemblyLocation = GetType().Assembly.Location;
                var config = ConfigurationManager.OpenExeConfiguration(thisAssemblyLocation);

                var targetAssembly = GetRequiredAppSetting(config, "TargetAssembly");
                var targetCommand = GetRequiredAppSetting(config, "TargetCommand");

                var assembly = LoadAssembly(thisAssemblyLocation, targetAssembly);
                var commandType = LoadCommandType(assembly, targetCommand);

                var command = assembly.CreateInstance(targetCommand);
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

        private static Type LoadCommandType(Assembly assembly, string targetCommand)
        {
            var commandType = assembly.GetType(targetCommand);

            if (commandType == null)
                throw new TargetException(String.Format("Command type {0} could not be found", targetCommand));
            return commandType;
        }

        private static Assembly LoadAssembly(string assemblyLocation, string targetAssembly)
        {
            var currentDirectory = Path.GetDirectoryName(assemblyLocation);
            var assemblyBytes = File.ReadAllBytes(Path.Combine(currentDirectory, targetAssembly));
            var assembly = Assembly.Load(assemblyBytes);
            return assembly;
        }

        private static string GetRequiredAppSetting(Configuration config, string key)
        {
            KeyValueConfigurationElement element = config.AppSettings.Settings[key];
            if (element != null)
            {
                string value = element.Value;
                if (!string.IsNullOrEmpty(value))
                    return value;
            }
            throw new SettingsPropertyNotFoundException(key);
        }
    }
}
