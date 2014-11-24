using System;
using System.Windows;
using Autodesk.Revit.UI;


namespace ElectricalToolSuite.FindAndReplace
{
    public static class DockExtensions
    {
        public static void RegisterDockablePane2(UIControlledApplication application,
            Guid id, string title, FrameworkElement dockElement, DockablePaneState state)
        {
            var dPid = new DockablePaneId(id);
            var dataProvider = new DockablePaneProviderData {FrameworkElement = dockElement, InitialState = state};
            IDockablePaneProvider provider = new DockPageProvider(dataProvider);
            application.RegisterDockablePane(dPid, title, provider);
        }

        public class DockPageProvider : IDockablePaneProvider
        {
            private readonly DockablePaneProviderData _data;

            public DockPageProvider(DockablePaneProviderData data)
            {
                _data = data;
            }

            public void SetupDockablePane(DockablePaneProviderData data)
            {
                data.FrameworkElement = _data.FrameworkElement;
                data.InitialState = _data.InitialState;
            }
        }
    }
}