using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using ElectricalToolSuite.ScheduleImport.CellFormatting;
using Microsoft.Win32;
using NetOffice.ExcelApi.Enums;
using Color = Autodesk.Revit.DB.Color;
using Orientation = ElectricalToolSuite.ScheduleImport.CellFormatting.Orientation;
using Excel = NetOffice.ExcelApi;

namespace ElectricalToolSuite.ScheduleImport
{
    [Transaction(TransactionMode.Automatic)]
    public class ExternalCommand : IExternalCommand
    {
        private Result TestScheduleCreation(Document doc, UIDocument uiDoc)
        {
//            var panelScheduleCategory = BuiltInCategory.OST_BranchPanelScheduleTemplates;

//            var vs = ViewSchedule.CreateSchedule(doc, new ElementId(panelScheduleCategory));

            var selectedPanelRef = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);

            var selectedPanel = doc.GetElement(selectedPanelRef) as FamilyInstance;

            if (selectedPanel == null)
            {
                TaskDialog.Show("Invalid selection", "Please select a panel.");
                return Result.Cancelled;
            }

            var existingSchedule =
                new FilteredElementCollector(doc).OfClass(typeof (PanelScheduleView))
                    .Cast<PanelScheduleView>()
                    .Where(sch => sch.GetPanel() == selectedPanel.Id)
                    .Take(1)
                    .ToList();

            PanelScheduleView schedule = null;

            if (existingSchedule.Any())
            {
                var confirmOperation = TaskDialog.Show("Overwrite existing schedule",
                    "The selected panel already has a schedule.  This operation will overwrite the existing schedule.  Proceed?",
                    TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel);
                if (confirmOperation == TaskDialogResult.Cancel)
                    return Result.Cancelled;

                schedule = existingSchedule.First();
            }
            else
            {
                schedule = PanelScheduleView.CreateInstanceView(doc, selectedPanel.Id);
//                schedule.ViewName = selectedPanel.Name;
            }

            var importPath = GetExcelFile();

            if (String.IsNullOrWhiteSpace(importPath))
                return Result.Cancelled;

//            ImportSchedule(schedule, importPath);

            return Result.Succeeded;

//            vs.ViewName = "SHOOP DA WHOOP";
        }

        private string GetExcelFile()
        {
            // Create an instance of the open file dialog box.
            OpenFileDialog fileDialog = new OpenFileDialog();

            // Set filter options and filter index.
            fileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            fileDialog.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            bool? userClickedOK = fileDialog.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == true)
            {
                return fileDialog.FileName;
            }

            return null;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {


            var watch = new Stopwatch();

            var uiDoc = commandData.Application.ActiveUIDocument;
            var doc = uiDoc.Document;


            var selectedPanelRef = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);

            var selectedPanel = doc.GetElement(selectedPanelRef) as FamilyInstance;

            if (selectedPanel == null)
            {
                TaskDialog.Show("Invalid selection", "Please select a panel.");
                return Result.Cancelled;
            }

            var existingSchedule =
                new FilteredElementCollector(doc).OfClass(typeof(PanelScheduleView))
                    .Cast<PanelScheduleView>()
                    .Where(s => s.GetPanel() == selectedPanel.Id)
                    .Take(1)
                    .ToList();

            PanelScheduleView schedule = null;

            if (existingSchedule.Any())
            {
                var confirmOperation = TaskDialog.Show("Overwrite existing schedule",
                    "The selected panel already has a schedule.  This operation will overwrite the existing schedule.  Proceed?",
                    TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel);
                if (confirmOperation == TaskDialogResult.Cancel)
                    return Result.Cancelled;

                schedule = existingSchedule.First();
            }
            else
            {
                schedule = PanelScheduleView.CreateInstanceView(doc, selectedPanel.Id);
                schedule.ViewName = selectedPanel.Name;
            }

            string workbookPath;
            string worksheetName;

            using (var excelApplication = new Excel.Application {DisplayAlerts = false})
            {
                var wnd = new UI.SheetSelectionDialog(excelApplication);

                if (wnd.ShowDialog() != true)
                    return Result.Cancelled;

                workbookPath = wnd.FilePathTextBox.Text;
                worksheetName = (string) wnd.SheetComboBox.SelectedItem;

                excelApplication.Quit();
            }

//            return TestScheduleCreation(doc, commandData.Application.ActiveUIDocument);

//            return Result.Cancelled;

//            var gfxStyles =
//                new FilteredElementCollector(doc).OfClass(typeof (GraphicsStyle)).Cast<GraphicsStyle>().ToList();

//            return Result.Succeeded;

//            var sch =
//                new FilteredElementCollector(doc).OfClass(typeof (PanelScheduleView))
//                    .Cast<PanelScheduleView>()
//                    .First(psv => !psv.IsTemplate);

//            var panelId = sch.GetPanel();

//            var panel = doc.GetElement(panelId);

//            return Result.Cancelled;


//            return Result.Cancelled;

//            if (sch == null)
//                throw new InvalidOperationException("No panel schedules found");

//            using (var trans = new Transaction(doc))
//            {
//                Debug.Assert(trans.Start(String.Format("Schedule Import - {0}", schedule.Name)) == TransactionStatus.Started);
                ImportSchedule(schedule, workbookPath, worksheetName);
//                trans.Commit();
//            }

            // TODO Put this back in
//            uiDoc.ActiveView = schedule;

            var elapsed = watch.Elapsed;

            TaskDialog.Show("Elapsed", elapsed.ToString());
            
            return Result.Succeeded;
        }

        private void ImportSchedule(PanelScheduleView schedule, string workbookPath, string worksheetName)
        {
            var tbl = schedule.GetTableData();

            tbl.GetSectionData(SectionType.Header).HideSection = true;
            tbl.GetSectionData(SectionType.Footer).HideSection = true;
            tbl.GetSectionData(SectionType.Summary).HideSection = true;
            var secData = tbl.GetSectionData(SectionType.Body);

            if (secData.NeedsRefresh)
                secData.RefreshData();

            //            var lineStyleIds = new HashSet<ElementId>();

            //            for (int i = secData.FirstRowNumber; i <= secData.LastRowNumber; ++i)
            //                for (int j = secData.FirstColumnNumber; j <= secData.LastColumnNumber; ++j)
            //                {
            //                    try
            //                    {
            //                        lineStyleIds.Add(secData.GetTableCellStyle(i, j).BorderTopLineStyle);
            //                        lineStyleIds.Add(secData.GetTableCellStyle(i, j).BorderBottomLineStyle);
            //                        lineStyleIds.Add(secData.GetTableCellStyle(i, j).BorderLeftLineStyle);
            //                        lineStyleIds.Add(secData.GetTableCellStyle(i, j).BorderRightLineStyle);
            //                    }
            //                    catch (Exception)
            //                    {
            //                    }
            //                }

            //            var sb = new StringBuilder();
            //            foreach (var id in lineStyleIds)
            //            {
            //                var elem = doc.GetElement(id);
            //                var name = "";
            //                if (elem != null) name = elem.Name;
            //                sb.AppendLine(String.Format("{0} - {1}",  name, id.IntegerValue.ToString()));
            //            }

            //            File.WriteAllText(@"C:\users\blake\desktop\linestyleids.txt", sb.ToString());

            //            return Result.Succeeded;

            Debug.WriteLine(schedule.ViewName);
            Debug.WriteLine(schedule.Name);

            // start excel and turn off msg boxes
            var excelApplication = new Excel.Application { DisplayAlerts = false };

            //            var linesCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            //            var lineStyleCategories = linesCategory.SubCategories;

//            var graphicsStyles = new FilteredElementCollector(doc).OfClass(typeof(GraphicsStyle))
//                    .Cast<GraphicsStyle>().ToList();

            var visibleLineStyleId =
                new ElementId(BuiltInCategory.OST_CurvesThinLines);

            //                graphicsStyles.First(gfx => gfx.Name == "Schedule Default").Id;
            //            var hiddenLineStyleId =
            //                graphicsStyles.First(gfx => gfx.Name == "<Invisible lines>").Id;

            //            var sb = new StringBuilder();

            //            foreach (var cat in cats)
            //            {
            //                sb.AppendLine(String.Format("{0}: {1}, {2}", cat.Name, (cat.GraphicsStyleCategory ?? cat.Category).Name, cat.GraphicsStyleType));
            //            }

            //            File.WriteAllText(@"C:\users\blake\desktop\gfxstyles.txt", sb.ToString());

            //            return Result.Cancelled;

            try
            {
                var wb = excelApplication.Workbooks.Open(
//                    @"C:\Users\Blake\Google Drive\ENPH 479 Revit Project\Electrical Panel Schedules\3690_Elec Panel Sch Working.xlsm",
                    workbookPath,
                    false, true);

                Debug.Assert(wb != null);

                var ws = wb.Worksheets.Cast<Excel.Worksheet>().First(s => s.Name == worksheetName);

                Debug.Assert(ws != null);

                var usedRange = ws.UsedRange;

                // TODO Fix off-by-one errors
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                var colWidths = new List<double>();
                var rowHeights = new List<double>();

                for (int i = 1; i <= rowCount; ++i)
                {
                    rowHeights.Add((double)usedRange[i, 1].RowHeight);
                }
                for (int j = 1; j <= colCount; ++j)
                {
                    colWidths.Add((double)usedRange[1, j].ColumnWidth);
                }

                while (secData.NumberOfRows < rowCount + 1)
                    secData.InsertRow(secData.LastRowNumber);

                while (secData.NumberOfRows > rowCount + 1)
                    secData.RemoveRow(secData.LastRowNumber);

                while (secData.NumberOfColumns < colCount + 1)
                    secData.InsertColumn(secData.LastColumnNumber);

                while (secData.NumberOfColumns > colCount + 1)
                    secData.RemoveColumn(secData.LastColumnNumber);

                for (int colIndex = 0; colIndex < colWidths.Count; ++colIndex)
                    secData.SetColumnWidthInPixels(colIndex + secData.FirstColumnNumber, (int)(colWidths[colIndex] * 4.0 / 4.0));

                for (int rowIndex = 0; rowIndex < rowHeights.Count; ++rowIndex)
                    secData.SetRowHeightInPixels(rowIndex + secData.FirstRowNumber, (int)(rowHeights[rowIndex] * 4.0 / 3.0));

                for (int colIndex = 0; colIndex < colWidths.Count; ++colIndex)
                    for (int rowIndex = 0; rowIndex < rowHeights.Count; ++rowIndex)
                        secData.ClearCell(rowIndex + secData.FirstRowNumber, colIndex + secData.FirstColumnNumber);

                if (secData.NeedsRefresh)
                    secData.RefreshData();

                var cells = CreateCells(usedRange);

                Debug.Assert(cells.Any());

                var overrideOptions = new TableCellStyleOverrideOptions
                {
                    FontSize = true,
                    Font = true,
                    HorizontalAlignment = true,
                    VerticalAlignment = true,
                    Bold = true,
                    Italics = true,
                    Underline = true,
                    BackgroundColor = true,
                    FontColor = true,
                    TextOrientation = true,
                    BorderBottomLineStyle = true,
                    BorderTopLineStyle = true,
                    BorderLeftLineStyle = true,
                    BorderRightLineStyle = true,
                    BorderLineStyle = true
                };

                for (int col = secData.FirstColumnNumber; col <= secData.LastColumnNumber; ++col)
                    secData.SetCellType(col, CellType.Text);

                if (secData.NeedsRefresh)
                    secData.RefreshData();

                foreach (var cell in cells)
                {
                    var mergedCell = new TableMergedCell(
                        cell.RowIndex + secData.FirstRowNumber,
                        cell.ColumnIndex + secData.FirstColumnNumber,
                        cell.RowIndex + secData.FirstRowNumber + cell.NumberOfRows - 1,
                        cell.ColumnIndex + secData.FirstColumnNumber + cell.NumberOfColumns - 1);

                    secData.MergeCells(mergedCell);
                }

                if (secData.NeedsRefresh)
                    secData.RefreshData();

                foreach (var cell in cells)
                {
                    secData.SetCellText(cell.RowIndex + secData.FirstRowNumber,
                        cell.ColumnIndex + secData.FirstColumnNumber,
                        cell.Text ?? "");

                    var fmt = new TableCellStyle();
                    fmt.ResetOverride();
                    fmt.SetCellStyleOverrideOptions(overrideOptions);

                    fmt.FontName = cell.FontName;
                    fmt.TextSize = cell.FontSize;
                    fmt.IsFontBold = cell.FontBold;
                    fmt.IsFontItalic = cell.FontItalic;
                    fmt.IsFontUnderline = cell.FontUnderline;
                    fmt.FontHorizontalAlignment = ConvertHorizontalAlignment(cell.HorizontalAlignment);
                    fmt.FontVerticalAlignment = ConvertVerticalAlignment(cell.VerticalAlignment);
                    fmt.BackgroundColor = ConvertColor(cell.BackgroundColor);
                    fmt.TextColor = ConvertColor(cell.FontColor);
                    fmt.TextOrientation = ConvertOrientation(cell.Orientation);

                    fmt.BorderBottomLineStyle = cell.BottomBorderLine == BorderLineStyle.Border
                        ? visibleLineStyleId
                        : ElementId.InvalidElementId;

                    fmt.BorderTopLineStyle = cell.TopBorderLine == BorderLineStyle.Border
                        ? visibleLineStyleId
                        : ElementId.InvalidElementId;

                    fmt.BorderLeftLineStyle = cell.LeftBorderLine == BorderLineStyle.Border
                        ? visibleLineStyleId
                        : ElementId.InvalidElementId;

                    fmt.BorderRightLineStyle = cell.RightBorderLine == BorderLineStyle.Border
                        ? visibleLineStyleId
                        : ElementId.InvalidElementId;

                    secData.SetCellStyle(cell.RowIndex + secData.FirstRowNumber,
                        cell.ColumnIndex + secData.FirstColumnNumber,
                        fmt);
                }
            }
            finally
            {
                // close excel and dispose reference
                excelApplication.Quit();
                excelApplication.Dispose();
            }

        }

        private Color ConvertColor(System.Drawing.Color color)
        {
            return new Color(color.R, color.G, color.B);
        }

        private HorizontalAlignment ConvertHorizontalAlignment(XlHAlign alignment)
        {
            switch (alignment)
            {
                case XlHAlign.xlHAlignCenter:
                case XlHAlign.xlHAlignCenterAcrossSelection:
                case XlHAlign.xlHAlignDistributed:
                case XlHAlign.xlHAlignFill:
                case XlHAlign.xlHAlignJustify:
                    return HorizontalAlignment.Center;
                case XlHAlign.xlHAlignLeft:
                case XlHAlign.xlHAlignGeneral:
                    return HorizontalAlignment.Left;
                case XlHAlign.xlHAlignRight:
                    return HorizontalAlignment.Right;
                default:
                    return HorizontalAlignment.Unknown;
            }
        }

        private VerticalAlignment ConvertVerticalAlignment(XlVAlign alignment)
        {
            switch (alignment)
            {
                case XlVAlign.xlVAlignDistributed:
                case XlVAlign.xlVAlignCenter:
                case XlVAlign.xlVAlignJustify:
                    return VerticalAlignment.Center;
                case XlVAlign.xlVAlignBottom:
                    return VerticalAlignment.Bottom;
                case XlVAlign.xlVAlignTop:
                    return VerticalAlignment.Top;
                default:
                    return VerticalAlignment.Unknown;
            }
        }

        private bool ConvertUnderline(XlUnderlineStyle underline)
        {
            switch (underline)
            {
                case XlUnderlineStyle.xlUnderlineStyleSingle:
                case XlUnderlineStyle.xlUnderlineStyleSingleAccounting:
                case XlUnderlineStyle.xlUnderlineStyleDouble:
                case XlUnderlineStyle.xlUnderlineStyleDoubleAccounting:
                    return true;
                default:
                    return false;
            }
        }

        private HorizontalAlignmentStyle ConvertHorizontalAlignment(HorizontalAlignment alignment)
        {
            switch (alignment)
            {
                case HorizontalAlignment.Center:
                    return HorizontalAlignmentStyle.Center;
                case HorizontalAlignment.Left:
                    return HorizontalAlignmentStyle.Left;
                case HorizontalAlignment.Right:
                    return HorizontalAlignmentStyle.Right;
                default:
                    throw new ArgumentException("alignment");
            }
        }

        private VerticalAlignmentStyle ConvertVerticalAlignment(VerticalAlignment alignment)
        {
            switch (alignment)
            {
                case VerticalAlignment.Center:
                    return VerticalAlignmentStyle.Middle;
                case VerticalAlignment.Bottom:
                    return VerticalAlignmentStyle.Bottom;
                case VerticalAlignment.Top  :
                    return VerticalAlignmentStyle.Top;
                default:
                    throw new ArgumentException("alignment");
            }
        }

        private List<Cell> CreateCells(Excel.Range range)
        {
            var cells = new List<Cell>();

            int numberOfColumns = range.Columns.Count;
            int numberOfRows = range.Rows.Count;

            for (int i = 1; i <= numberOfRows; ++i)
            {
                for (int j = 1; j <= numberOfColumns; ++j)
                {
                    var cell = range[i, j];
                    if (!((bool) cell.MergeCells && !cell.MergeArea.Address.StartsWith(cell.Address)))
                        cells.Add(CreateCell(cell, i, j));

                    Debug.Print("{0}, {1}", i, j);
                }
            }

            return cells;
        }

        private Orientation ConvertOrientation(XlOrientation excelOrientation)
        {
            switch (excelOrientation)
            {
                case XlOrientation.xlDownward:
                    return Orientation.Downward;
                case XlOrientation.xlHorizontal:
                    return Orientation.Horizontal;
                case XlOrientation.xlUpward:
                    return Orientation.Upward;
                case XlOrientation.xlVertical:
                    return Orientation.Vertical;
                default:
                    return Orientation.Unknown;
            }
        }

        private int ConvertOrientation(Orientation orientation)
        {
            switch (orientation)
            {
                case Orientation.Vertical:
                case Orientation.Upward:
                    return 900;
                case Orientation.Downward:
                    return -900;
                default:
                    return 0;
            }
        }

        private BorderLineStyle ConvertBorder(Excel.Border border)
        {
            var style = (XlLineStyle) border.LineStyle;

            switch (style)
            {
                case XlLineStyle.xlLineStyleNone:
                    return BorderLineStyle.NoBorder;
                default:
                    return BorderLineStyle.Border;
            }
        }

        private Cell CreateCell(Excel.Range interopCell, int rowIndex, int columnIndex)
        {
            var font = interopCell.Font;
            var borders = interopCell.Borders;
            var interior = interopCell.Interior;

            var cell = new Cell
            {
                RowIndex = rowIndex - 1,
                ColumnIndex = columnIndex - 1,
                Text = (interopCell.Text ?? interopCell.Value2 ?? "").ToString(),
                FontName = (string) font.Name,
                FontSize = (double) font.Size,
                FontBold = (bool) font.Bold,
                FontItalic = (bool) font.Italic,
                FontUnderline = ConvertUnderline((XlUnderlineStyle) font.Underline),
                BackgroundColor = ColorTranslator.FromOle(Convert.ToInt32((double) interior.Color)),
                FontColor = ColorTranslator.FromOle(Convert.ToInt32((double) font.Color)),
                HorizontalAlignment = ConvertHorizontalAlignment((XlHAlign) interopCell.HorizontalAlignment),
                VerticalAlignment = ConvertVerticalAlignment((XlVAlign) interopCell.VerticalAlignment),
                Orientation = ConvertOrientation((XlOrientation) interopCell.Orientation),

                BottomBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeBottom]),
                TopBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeTop]),
                LeftBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeLeft]),
                RightBorderLine = ConvertBorder(borders[XlBordersIndex.xlEdgeRight])
            };

            if (cell.BackgroundColor == System.Drawing.Color.White &&
                ((XlPattern) interior.Pattern) != XlPattern.xlPatternNone)
            {
                cell.BackgroundColor = System.Drawing.Color.LightGray;
            }

            if ((bool) interopCell.MergeCells)
            {
                var mergeArea = interopCell.MergeArea;
                cell.NumberOfColumns = mergeArea.Columns.Count;
                cell.NumberOfRows = mergeArea.Rows.Count;
            }
            else
            {
                cell.NumberOfColumns = 1;
                cell.NumberOfRows = 1;
            }

            Debug.Assert(cell.ColumnIndex >= 0);
            Debug.Assert(cell.RowIndex >= 0);

            return cell;
        }

    }
}
