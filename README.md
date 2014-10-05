enph479
=======

## How is DLL formed?

Put `RevitAPI.dll`, `RevitAPI.xml`, `RevitAPIUI.dll`, and `RevitAPIUI.xml` into a new root folder, `/lib` (`../enph479/lib`, next to `../enph479/src`).

Open `ElectricalToolSuite.sln` and build it.  Also right click on the ExtensionLoader project and build that too - it's not built normally because it will stay loaded in Revit and the file will be locked. Check the Output pane to see where it put the ExtensionLoader DLL (`../enph479/src/bin/Debug/ExtensionLoader.dll`) and copy that path.

Add a `.addin` manifest file for the Extension Loader in `ProgramData/AutoDesk/Revit/` pointing to the ExtensionLoader DLL.

Open a Revit project and run Extension Loader - it should load and run the M&E tool.  You can rebuild the tool without restarting Revit now.

If you want to use another tool, change the configuration file `ExtensionLoader.dll.config` in the build directory (`../enph479/src/bin/Debug/ExtensionLoader.dll.config`). You won't have to rebuild or restart Revit.
