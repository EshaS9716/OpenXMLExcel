using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using GemBox.Spreadsheet;
using AutoFilter = DocumentFormat.OpenXml.Spreadsheet.AutoFilter;
using SortState = DocumentFormat.OpenXml.Spreadsheet.SortState;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace ExcelAppOpenXML
{
    public class Export_Data
    {
        public static readonly string date = DateTime.Now.ToString("yyyy-MM-dd");
        public static string filePath = _Default.DesPath;

        public static void WriteToExcel()
        {
            try
            {
                DataTable dt = GetDataFromAPI.dataTable2;
                File.Copy(_Default.SourcePath, filePath, true);
                int columnLen1 = 0, columnLen2 = 0, columnLen3 = 0;

                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
                {
                    WorksheetPart worksheetPart1 = GetWorksheetPartByName(spreadSheet, "Database & Connectivity");
                    WorksheetPart worksheetPart2 = GetWorksheetPartByName(spreadSheet, "Power Systems");
                    WorksheetPart worksheetPart3 = GetWorksheetPartByName(spreadSheet, "Z Systems");

                    #region Populate & Merge
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        int sheetNum = int.Parse(dt.Rows[i].ItemArray[1].ToString());
                        uint rows = uint.Parse(dt.Rows[i].ItemArray[dt.Columns.Count - 1].ToString());
                        int columns = int.Parse(dt.Rows[i].ItemArray[dt.Columns.Count - 2].ToString());
                        string cellData = dt.Rows[i].ItemArray[dt.Columns.Count - 3].ToString();
                        string cellDataLevel2 = dt.Rows[i].ItemArray[dt.Columns.Count - 8].ToString();
                        string cellDataLevel3 = dt.Rows[i].ItemArray[dt.Columns.Count - 7].ToString();
                        string cellDataLevel4 = dt.Rows[i].ItemArray[dt.Columns.Count - 6].ToString();

                        if (sheetNum == 1)
                        {
                            InsertTextExistingExcel(spreadSheet, worksheetPart1, columns, rows, cellData, false);
                            InsertTextExistingExcel(spreadSheet, worksheetPart1, columns, 5, cellDataLevel2, false);
                            InsertTextExistingExcel(spreadSheet, worksheetPart1, columns, 6, cellDataLevel3, false);
                            InsertTextExistingExcel(spreadSheet, worksheetPart1, columns, 7, cellDataLevel4, false);
                            columnLen1 = columns;
                        }
                        else if (sheetNum == 2)
                        {
                            InsertTextExistingExcel(spreadSheet, worksheetPart2, columns, rows, cellData, false);
                            InsertTextExistingExcel(spreadSheet, worksheetPart2, columns, 5, cellDataLevel2, false);
                            InsertTextExistingExcel(spreadSheet, worksheetPart2, columns, 6, cellDataLevel3, false);
                            InsertTextExistingExcel(spreadSheet, worksheetPart2, columns, 7, cellDataLevel4, false);
                            columnLen2 = columns;
                        }
                        else if (sheetNum == 3)
                        {
                            InsertTextExistingExcel(spreadSheet, worksheetPart3, columns, rows, cellData, false);
                            InsertTextExistingExcel(spreadSheet, worksheetPart3, columns, 5, cellDataLevel2, false);
                            InsertTextExistingExcel(spreadSheet, worksheetPart3, columns, 6, cellDataLevel3, false);
                            InsertTextExistingExcel(spreadSheet, worksheetPart3, columns, 7, cellDataLevel4, false);
                            columnLen3 = columns;
                        }
                    }

                    MergeSheetData(spreadSheet, worksheetPart1, columnLen1, 2, 2);
                    MergeSheetData(spreadSheet, worksheetPart2, columnLen2, 2, 2);
                    MergeSheetData(spreadSheet, worksheetPart3, columnLen3, 2, 2);
                    #endregion
                }
            }
            catch (Exception ex)
            {
                ErrorLogging.SendErrorToText(ex);
                throw ex;
            }
            WriteToExcel1();
            Autofit();
        }

        private static void WriteToExcel1()
        {
            try
            {
                DataTable dt = GetDataFromAPI.dataTable1;
                DataTable dtCount = GetDataFromAPI.dataTable3;
                int rowSource, rowDestination, column, i, j;
                Hyperlinks hyperlinks1 = new Hyperlinks();

                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
                {
                    WorksheetPart worksheetPart1 = GetWorksheetPartByName(spreadSheet, "Hierarchy");
                    WorksheetPart worksheetPart5 = GetWorksheetPartByName(spreadSheet, "Product Summary");

                    Worksheet worksheet = worksheetPart1.Worksheet;

                    #region Populate
                    for (i = 0; i <= dt.Rows.Count - 1; i++)
                    {
                        for (j = 0; j <= dt.Columns.Count - 2; j += 2)
                        {
                            uint row = uint.Parse(dt.Rows[i].ItemArray[dt.Columns.Count - 1].ToString());
                            int columns = int.Parse(dt.Rows[i].ItemArray[j + 1].ToString());

                            string cellData = dt.Rows[i].ItemArray[j].ToString();
                            InsertTextExistingExcel(spreadSheet, worksheetPart1, columns, row, cellData, true);
                        }
                    }
                    #endregion

                    #region Populate Count Page

                    for (i = 0; i <= dtCount.Rows.Count - 1; i++)
                    {
                        for (j = 0; j <= dtCount.Columns.Count - 1; j++)
                        {
                            string cellData = dtCount.Rows[i].ItemArray[j].ToString();
                            bool buChanged = false;
                            if (dtCount.Rows.Count == i + 1)
                            {
                                buChanged = true;
                            }
                            else if (dtCount.Rows[i].ItemArray[0].ToString() != dtCount.Rows[i + 1].ItemArray[0].ToString())
                            {
                                buChanged = true;
                            }
                            InsertTextExistingExcel(spreadSheet, worksheetPart5, j + 1, (uint)(i + 2), cellData, buChanged);
                        }
                    }

                    #endregion

                    #region Merge and Format
                    for (i = 0; i < (dt.Columns.Count / 2) - 1; i += 2)
                    {
                        j = 0;
                        rowSource = int.Parse(dt.Rows[j].ItemArray[dt.Columns.Count - 1].ToString());
                        column = int.Parse(dt.Rows[j].ItemArray[i + 1].ToString());
                        rowDestination = rowSource;

                        for (j = 0; j < dt.Rows.Count; j++)
                        {
                            if (j == dt.Rows.Count - 1)
                            {
                                rowDestination = int.Parse(dt.Rows[j].ItemArray[dt.Columns.Count - 1].ToString());
                                if (i == 0)
                                {
                                    int colSrc = int.Parse(dt.Rows[j].ItemArray[dt.Columns.Count - 8].ToString());
                                    uint colDes = uint.Parse(dt.Rows[j].ItemArray[dt.Columns.Count - 2].ToString());
                                    for (int col = colSrc; col <= colDes; col++)
                                    {
                                        for (uint row = 6; row <= rowDestination; row++)
                                        {
                                            StylesSheet1.AddBold(spreadSheet, GetSpreadsheetCell(worksheet, ColumnLetter(col - 1), row), col == colSrc, false, false, false);
                                            if (col == colSrc)
                                            {
                                                hyperlinks1.Append(AddLink(worksheet, col, (int)row));
                                            }
                                        }
                                    }
                                }
                                if (GetRowNumber(column) == 1 && rowDestination > 22)
                                {
                                    for (uint val = 23; val <= rowDestination; val++)
                                    {
                                        DeleteTextFromCell(worksheetPart1, ColumnLetter(column - 1), val);
                                    }
                                    rowDestination = 22;
                                }
                                Merge(worksheet, ColumnLetter(column - 1) + rowSource, ColumnLetter(column - 1) + rowDestination);
                                hyperlinks1.Append(AddLink(worksheet, column, rowSource));
                                StylesSheet1.AddBold(spreadSheet, GetSpreadsheetCell(worksheet, ColumnLetter(column - 1), (uint)rowSource), false, GetRowNumber(column) == 3, GetRowNumber(column) == 2, GetRowNumber(column) == 1);
                            }
                            else if (dt.Rows[j].ItemArray[i].ToString() == dt.Rows[j + 1].ItemArray[i].ToString())
                            {
                                continue;
                            }
                            else
                            {
                                rowDestination = int.Parse(dt.Rows[j].ItemArray[dt.Columns.Count - 1].ToString());
                                if (i == 0)
                                {
                                    int colSrc = int.Parse(dt.Rows[j].ItemArray[dt.Columns.Count - 8].ToString());
                                    uint colDes = uint.Parse(dt.Rows[j].ItemArray[dt.Columns.Count - 2].ToString());
                                    for (int col = colSrc; col <= colDes; col++)
                                    {
                                        for (uint row = 6; row <= rowDestination; row++)
                                        {
                                            StylesSheet1.AddBold(spreadSheet, GetSpreadsheetCell(worksheet, ColumnLetter(col - 1), row), col == colSrc, false, false, false);
                                            if (col == colSrc)
                                            {
                                                hyperlinks1.Append(AddLink(worksheet, col, (int)row));
                                            }
                                        }
                                    }
                                }
                                if (GetRowNumber(column) == 1 && rowDestination > 22)
                                {
                                    for (uint val = 23; val <= rowDestination; val++)
                                    {
                                        DeleteTextFromCell(worksheetPart1, ColumnLetter(column - 1), val);
                                    }
                                    rowDestination = 22;
                                }
                                Merge(worksheet, ColumnLetter(column - 1) + rowSource, ColumnLetter(column - 1) + rowDestination);
                                hyperlinks1.Append(AddLink(worksheet, column, rowSource));
                                StylesSheet1.AddBold(spreadSheet, GetSpreadsheetCell(worksheet, ColumnLetter(column - 1), (uint)rowSource), false, GetRowNumber(column) == 3, GetRowNumber(column) == 2, GetRowNumber(column) == 1);
                                rowSource = int.Parse(dt.Rows[j + 1].ItemArray[dt.Columns.Count - 1].ToString());
                                column = int.Parse(dt.Rows[j + 1].ItemArray[i + 1].ToString());
                            }
                        }
                    }
                    #endregion

                    InsertHyperLinkInWorksheet(worksheet, hyperlinks1);
                }
            }
            catch (Exception ex)
            {
                ErrorLogging.SendErrorToText(ex);
                throw ex;
            }
        }

        public static void Autofit()
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

                var workbook = ExcelFile.Load(filePath);

                var worksheet = workbook.Worksheets[0];
                for (int j = 0; j < 5; j++)
                {
                    worksheet = workbook.Worksheets[j];
                    int columnCount = worksheet.CalculateMaxUsedColumns();
                    for (int i = 0; i < columnCount; i++)
                        worksheet.Columns[i].AutoFit(1, worksheet.Rows[0], worksheet.Rows[worksheet.Rows.Count - 1]);
                }

                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                ErrorLogging.SendErrorToText(ex);
                throw ex;
            }
        }

        #region HelperMethods

        static Hyperlink AddLink(Worksheet worksheet, int column, int rowSource)
        {
            string src = GetSpreadsheetCell(worksheet, ColumnLetter(column - 1), Convert.ToUInt32(rowSource)).CellReference;
            string des;
            string value = GetSpreadsheetCell(worksheet, ColumnLetter(column - 1), Convert.ToUInt32(rowSource)).InnerText;
            int desCol = FindStringColumn(GetRowNumber(column), value, GetDataFromAPI.dataTable2, GetBuId(column));
            des = GetWorksheetName(column) + "!" + ColumnLetter(desCol == 0 ? 1 : desCol - 1) + GetColumnNumber(column);
            Hyperlink hyperlink1 = new Hyperlink() { Reference = src, Location = des };
            return hyperlink1;
        }

        static int FindStringColumn(int col, string str, DataTable dt, int buId)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string str1 = dt.Rows[i].ItemArray[dt.Columns.Count - 9].ToString();
                if (int.Parse(str1) == buId)
                {
                    string str2 = dt.Rows[i].ItemArray[col].ToString();
                    if (str2 == str)
                    {
                        return int.Parse(dt.Rows[i].ItemArray[dt.Columns.Count - 2].ToString());
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            return 0;
        }

        static string GetWorksheetName(int i)
        {
            switch (i)
            {
                case 1: case 2: case 3: case 4: return "'Database & Connectivity'";
                case 10: case 11: case 12: case 13: return "'Power Systems'";
                default: return "'Z Systems'";
            }
        }

        private static int GetColumnNumber(int i)
        {
            if (i == 4 || i == 13 || i == 22)
            {
                return 7;
            }
            else if (i == 3 || i == 12 || i == 21)
            {
                return 6;
            }
            else if (i == 2 || i == 11 || i == 20)
            {
                return 5;
            }
            else if (i == 1 || i == 10 || i == 19)
            {
                return 4;
            }
            else
            {
                return 2;
            }
        }

        private static int GetBuId(int i)
        {
            switch (i)
            {
                case 1: case 2: case 3: case 4: return 1;
                case 10: case 11: case 12: case 13: return 2;
                default: return 3;
            }
        }

        private static void InsertHyperLinkInWorksheet(Worksheet worksheet, Hyperlinks hyperlinks1)
        {
            PageMargins pageMargins = worksheet.Descendants<PageMargins>().First();
            worksheet.InsertBefore<Hyperlinks>(hyperlinks1, pageMargins);
            worksheet.Save();
        }

        public static void DeleteTextFromCell(WorksheetPart worksheetPart, string colName, uint rowIndex)
        {
            Cell cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex);
            cell.Remove();
            worksheetPart.Worksheet.Save();
        }

        private static int GetRowNumber(int i)
        {
            if (i == 3 || i == 12 || i == 21)
            {
                return 3;
            }
            else if (i == 2 || i == 11 || i == 20)
            {
                return 2;
            }
            else if (i == 1 || i == 10 || i == 19)
            {
                return 1;
            }
            else
            {
                return 4;
            }
        }

        public static void Merge(Worksheet worksheet, string cell1Name, string cell2Name)
        {
            MergeCells mergeCells;

            if (worksheet.Elements<MergeCells>().Count() > 0)
            {
                mergeCells = worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();
                if (worksheet.Elements<CustomSheetView>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                }
                else if (worksheet.Elements<DataConsolidate>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                }
                else if (worksheet.Elements<SortState>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                }
                else if (worksheet.Elements<AutoFilter>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                }
                else if (worksheet.Elements<Scenarios>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                }
                else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                }
                else if (worksheet.Elements<SheetProtection>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                }
                else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                }
            }

            MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) };
            mergeCells.Append(mergeCell);
            worksheet.Save();
        }

        private static void MergeSheetData(SpreadsheetDocument spreadSheet, WorksheetPart worksheetPart, int columnLen, int column1SrcLevel3, int columnSrcLevel2)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            MergeCells mergeCells = new MergeCells();
            string str;
            for (int i = 2; i < columnLen + 1; i++)
            {
                if (GetSpreadsheetCell(worksheet, ColumnLetter(i), 6) == null)
                {
                    int columnDestLevel3 = i;
                    str = ColumnLetter(column1SrcLevel3 - 1) + "6" + ":" + ColumnLetter(columnDestLevel3 - 1) + "6";
                    mergeCells.Append(new MergeCell() { Reference = new StringValue(str) });
                    StylesSheet2.AddBold(spreadSheet, GetSpreadsheetCell(worksheet, ColumnLetter(column1SrcLevel3 - 1), 6), 6);
                    column1SrcLevel3 = columnDestLevel3 + 1;
                }
                else if (GetSpreadsheetCell(worksheet, ColumnLetter(i - 1), 6).InnerText == GetSpreadsheetCell(worksheet, ColumnLetter(i), 6).InnerText)
                {
                    continue;
                }
                else
                {
                    int columnDestLevel3 = i;
                    str = ColumnLetter(column1SrcLevel3 - 1) + "6" + ":" + ColumnLetter(columnDestLevel3 - 1) + "6";
                    mergeCells.Append(new MergeCell() { Reference = new StringValue(str) });
                    StylesSheet2.AddBold(spreadSheet, GetSpreadsheetCell(worksheet, ColumnLetter(column1SrcLevel3 - 1), 6), 6);
                    column1SrcLevel3 = columnDestLevel3 + 1;
                }
            }
            for (int i = 2; i < columnLen + 1; i++)
            {
                if (GetSpreadsheetCell(worksheet, ColumnLetter(i), 5) == null)
                {
                    int columnDestLevel2 = i;
                    str = ColumnLetter(columnSrcLevel2 - 1) + "5" + ":" + ColumnLetter(columnDestLevel2 - 1) + "5";
                    mergeCells.Append(new MergeCell() { Reference = new StringValue(str) });
                    StylesSheet2.AddBold(spreadSheet, GetSpreadsheetCell(worksheet, ColumnLetter(columnSrcLevel2 - 1), 5), 5);
                    columnSrcLevel2 = columnDestLevel2 + 1;
                }
                else if (GetSpreadsheetCell(worksheet, ColumnLetter(i - 1), 5).InnerText == GetSpreadsheetCell(worksheet, ColumnLetter(i), 5).InnerText)
                {
                    continue;
                }
                else
                {
                    int columnDestLevel2 = i;
                    str = ColumnLetter(columnSrcLevel2 - 1) + "5" + ":" + ColumnLetter(columnDestLevel2 - 1) + "5";
                    mergeCells.Append(new MergeCell() { Reference = new StringValue(str) });
                    StylesSheet2.AddBold(spreadSheet, GetSpreadsheetCell(worksheet, ColumnLetter(columnSrcLevel2 - 1), 5), 5);
                    columnSrcLevel2 = columnDestLevel2 + 1;
                }
            }

            str = ColumnLetter(1) + "4" + ":" + ColumnLetter(columnSrcLevel2 - 2) + "4";
            mergeCells.Append(new MergeCell() { Reference = new StringValue(str) });
            StylesSheet2.AddBold(spreadSheet, GetSpreadsheetCell(worksheet, ColumnLetter(1), 4), 4);

            worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
            worksheet.Save();
        }

        public static Cell GetSpreadsheetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex);
            if (rows.Count() == 0)
            {
                // A cell does not exist at the specified row.
                return null;
            }

            IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
            if (cells.Count() == 0)
            {
                // A cell does not exist at the specified column, in the specified row.
                return null;
            }

            return cells.First();
        }

        public static void InsertTextExistingExcel(SpreadsheetDocument spreadSheet, WorksheetPart worksheetPart, int columns, uint rows, string cellData, bool isPage1)
        {
            if (GetSpreadsheetCell(worksheetPart.Worksheet, ColumnLetter(columns - 1), rows) == null)
            {
                Cell cell = InsertCellInWorksheet(ColumnLetter(columns - 1), rows, worksheetPart);
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString() { Text = new Text(cellData) };
                if (worksheetPart == GetWorksheetPartByName(spreadSheet, "Product Summary"))
                {
                    StylesSheet5.AddBold(spreadSheet, cell, columns, isPage1);
                }
                else if (rows == 7 && !isPage1)
                {
                    StylesSheet2.AddBold(spreadSheet, cell, 7);
                }
                worksheetPart.Worksheet.Save();
            }
        }

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            Cell refCell = row.Descendants<Cell>().LastOrDefault();
            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertAfter(newCell, refCell);
            worksheet.Save();
            return newCell;
        }

        public static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
                            Elements<Sheet>().Where(s => s.Name == sheetName);

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }

        private static string ColumnLetter(int intCol)
        {
            var intFirstLetter = ((intCol) / 676) + 64;
            var intSecondLetter = ((intCol % 676) / 26) + 64;
            var intThirdLetter = (intCol % 26) + 65;

            var firstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
            var secondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : ' ';
            var thirdLetter = (char)intThirdLetter;

            return string.Concat(firstLetter, secondLetter, thirdLetter).Trim();
        }
        #endregion
    }
}