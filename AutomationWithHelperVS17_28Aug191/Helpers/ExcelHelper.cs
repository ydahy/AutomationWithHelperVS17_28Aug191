using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using spd = DocumentFormat.OpenXml.Spreadsheet;

namespace AutomationWithHelperVS17_28Aug191
{
    class ExcelHelper
    {
        enum ColumnName
        {
            A = 0,
            B = 1,
            C = 2,
            D = 3,
            E = 4,
            F = 5,
            G = 6,
            H = 7,
            I = 8
        }

        // Create Excel Sheet
        public void UpdateCells(string docName, IList<ExcelSheetInput> excelSheetInputs)
        {
            using (SpreadsheetDocument createdDocument = SpreadsheetDocument.Create(docName, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart createdWorkbookPart = createdDocument.AddWorkbookPart();
                createdWorkbookPart.Workbook = new spd.Workbook();
                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart createdWorksheetPart = createdWorkbookPart.AddNewPart<WorksheetPart>();
                createdWorksheetPart.Worksheet = new spd.Worksheet();

                spd.Sheets createdSheets = createdWorkbookPart.Workbook.AppendChild(new spd.Sheets());

                spd.Sheet createdSheet = new spd.Sheet()
                {
                    Id = createdWorkbookPart.GetIdOfPart(createdWorksheetPart),
                    SheetId = 1,
                    Name = "Sheet 1"
                };
                createdSheets.Append(createdSheet);

                createdWorkbookPart.Workbook.Save();

                spd.SheetData sheetData = createdWorksheetPart.Worksheet.AppendChild(new spd.SheetData());

                var rowsCount = excelSheetInputs.Last().RowNumber;

                for (var rowsCounter = 1; rowsCounter <= rowsCount; rowsCounter++)
                {
                    var rowItems = excelSheetInputs.Where(e => e.RowNumber == rowsCounter).ToList();

                    // Constructing header
                    spd.Row createdRow = new spd.Row();

                    foreach (ColumnName c in Enum.GetValues(typeof(ColumnName)))
                    {
                        try
                        {
                            var cellValue = rowItems.First(e => e.ColumnName.Equals(c.ToString()));
                            createdRow.Append(ConstructCell(cellValue.Text, spd.CellValues.String));
                        }
                        catch (Exception)
                        {
                            createdRow.Append(ConstructCell("", spd.CellValues.String));
                        }
                    }

                    sheetData.AppendChild(createdRow);
                }

                createdWorkbookPart.Workbook.Save();
            }
        }

        private static spd.Cell ConstructCell(string value, spd.CellValues dataType)
        {
            return new spd.Cell()
            {
                CellValue = new spd.CellValue(value),
                DataType = new EnumValue<spd.CellValues>(dataType)
            };
        }

        // Read Excel Sheet
        public List<string[]> GetExcelSheetData(out Dictionary<string, string> excelHeaders, out string windowWidth, out string windowHeight)
        {
            excelHeaders = null;
            var headers = new Dictionary<string, string>
            {
                {"0", "URL"},
                {"1", "Locator"},
                {"2", "Location X"},
                {"3", "Location Y"},
                {"4", "Scroll Direction"},
                {"5", "Is Homepage"},
                {"6", "Is Repeated"}
            };

            windowWidth = "";
            windowHeight = "";

            var rowsVals = new List<string[]>();

            using (var spreadsheetDocument = SpreadsheetDocument.Open(TestHelper.SetDir("UIElements.xlsx"), false))
            {
                var sheets = spreadsheetDocument.WorkbookPart.Workbook.Descendants<spd.Sheet>();

                foreach (var sheet in sheets)
                {
                    if (sheet.SheetId > 1) break;
                    uint rowIndex = 0;

                    var worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id);

                    var worksheet = worksheetPart.Worksheet;
                    var sheetData = worksheet.GetFirstChild<spd.SheetData>();
                    var rows = sheetData.Elements<spd.Row>().GetEnumerator();

                    var seqId = 1;
                    while (rows.MoveNext())
                    {
                        if (rowIndex == 0)
                        {
                            windowWidth = GetCellValue(spreadsheetDocument, GetCell(worksheet, GetColName(1), rowIndex)).Split(':')[1];
                            windowHeight = GetCellValue(spreadsheetDocument, GetCell(worksheet, GetColName(2), rowIndex)).Split(':')[1];
                        }
                        if (rowIndex == 1)
                        {
                            excelHeaders = GetExcelHeaders(spreadsheetDocument, rowIndex, worksheet);
                        }
                        if (rowIndex > 1)
                        {
                            var colVals = new string[headers.Count];
                            for (uint _iCols = 0; _iCols < headers.Count; _iCols++)
                            {
                                var hashColIndex = -1;
                                hashColIndex = GetHashColIndex(excelHeaders, headers[_iCols.ToString()]);
                                if (hashColIndex != -1)
                                {
                                    var colIndex = hashColIndex;
                                    string colValue = null;

                                    var colName = GetColName((uint)colIndex);
                                    var cell = GetCell(worksheet, colName, rowIndex);
                                    if (cell != null)
                                    {
                                        colValue = GetCellValue(spreadsheetDocument, cell);
                                    }
                                    colVals[_iCols] = colValue;
                                }
                            }

                            rowsVals.Add(colVals);

                            seqId++;
                        }
                        rowIndex++;
                    }
                }
            }

            return rowsVals;
        }

        private static int GetHashColIndex(Dictionary<string, string> headers, string headerVal)
        {
            var iIndex = -1;
            var curr = headers.FirstOrDefault(x => x.Value == headerVal);
            if (curr.Key != null)
                iIndex = int.Parse(curr.Key);
            return iIndex;
        }

        private static string GetCellValue(SpreadsheetDocument spreadsheetDocument, spd.Cell cell)
        {
            string colValue = null;
            if (cell.DataType != null &&
                cell.DataType == spd.CellValues.SharedString)
                colValue = GetSharedStringItemById(spreadsheetDocument.WorkbookPart, int.Parse(cell.CellValue.Text));
            else
            {
                if (cell.CellValue != null)
                    colValue = ((spd.CellValue)cell.FirstChild).Text;
            }
            return colValue;
        }

        private static Dictionary<string, string> GetExcelHeaders(SpreadsheetDocument spreadsheetDocument, uint rowIndex, spd.Worksheet worksheet)
        {
            uint colIndex = 1;
            var headers = new Dictionary<string, string>();

            while (true)
            {
                string colValue = null;
                var colName = GetColName(colIndex);
                var cell = GetCell(worksheet, colName, rowIndex);
                if (cell != null)
                {
                    colValue = GetCellValue(spreadsheetDocument, cell);
                }
                if (string.IsNullOrEmpty(colValue))
                {
                    break;
                }
                headers.Add(colIndex.ToString(), colValue);
                colIndex++;

            }
            return headers;
        }

        private static string GetColName(uint colIndex)
        {
            return ((char)('A' + colIndex - 1)).ToString();
        }

        private static spd.Cell GetCell(spd.Worksheet worksheet, string columnName, uint rowIndex)
        {
            var row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;
            //    if (row.Hidden!=null) 
            //        return null;
            spd.Cell cell = null;
            try
            {
                cell = row.Elements<spd.Cell>().ElementAt((int)(ColumnName)Enum.Parse(typeof(ColumnName), columnName));
            }
            catch (Exception)
            {
                // ignored
            }
            return cell;
        }

        private static spd.Row GetRow(spd.Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<spd.SheetData>().
                Elements<spd.Row>().ElementAt((int)rowIndex);
        }

        public static string GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<spd.SharedStringItem>()
                .ElementAt(id).InnerText;
        }

    }

    public class ExcelSheetInput
    {
        public string FileName { set; get; }
        public string Text { set; get; }
        public uint RowNumber { set; get; }
        public string ColumnName { set; get; }

        public ExcelSheetInput(string fileName, string text, uint rowNumber, string columnName)
        {
            FileName = fileName;
            Text = text;
            RowNumber = rowNumber;
            ColumnName = columnName;
        }
    }
}
