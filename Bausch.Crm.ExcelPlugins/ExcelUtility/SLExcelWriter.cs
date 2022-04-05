using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;

namespace Crm.ExcelPlugins
{
    public class SLExcelWriter
    {
        private string ColumnLetter(int intCol)
        {
            try
            {
                var intFirstLetter = ((intCol) / 676) + 64;
                var intSecondLetter = ((intCol % 676) / 26) + 64;
                var intThirdLetter = (intCol % 26) + 65;

                var firstLetter = (intFirstLetter > 64)
                    ? (char)intFirstLetter : ' ';
                var secondLetter = (intSecondLetter > 64)
                    ? (char)intSecondLetter : ' ';
                var thirdLetter = (char)intThirdLetter;

                return string.Concat(firstLetter, secondLetter,
                    thirdLetter).Trim();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private Cell CreateTextCell(string header, UInt32 index,string text)
        {
            try
            {
                var cell = new Cell
                {
                    DataType = CellValues.InlineString,
                    CellReference = header + index
                };

                var istring = new InlineString();
                var t = new Text { Text = text };
                istring.AppendChild(t);
                cell.AppendChild(istring);
                return cell;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public byte[] GenerateExcel(SLExcelData data)
        {
            try
            {
                var stream = new MemoryStream();
                var document = SpreadsheetDocument
                    .Create(stream, SpreadsheetDocumentType.Workbook);

                var workbookpart = document.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();

                worksheetPart.Worksheet = new Worksheet(sheetData);

                var sheets = document.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                var sheet = new Sheet()
                {
                    Id = document.WorkbookPart
                    .GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = data.SheetName ?? "Sheet 1"
                };
                sheets.AppendChild(sheet);

                // Add header
                UInt32 rowIdex = 0;
                var row = new Row { RowIndex = ++rowIdex };
                sheetData.AppendChild(row);
                var cellIdex = 0;

                if (data !=null && data.Headers !=null && data.Headers.Count !=0)
                {
                    foreach (var header in data.Headers)
                    {
                        string textHeader = !string.IsNullOrEmpty(header) ? header : string.Empty;
                        row.AppendChild(CreateTextCell(ColumnLetter(cellIdex++),
                            rowIdex, textHeader));
                    } 
                }

                if (data !=null )
                {
                    // Add the column configuration if available
                    if (data.ColumnConfigurations != null)
                    {
                        var columns = (Columns)data.ColumnConfigurations.Clone();
                        if (columns !=null)
                        {
                            worksheetPart.Worksheet
                           .InsertAfter(columns, worksheetPart
                           .Worksheet.SheetFormatProperties);
                        }
                    }
                }

                // Add sheet data
                foreach (System.Collections.Generic.List<string> rowData in data.DataRows)
                {
                    if (rowData != null && rowData.Count != 0)
                    {
                        cellIdex = 0;
                        row = new Row { RowIndex = ++rowIdex };
                        sheetData.AppendChild(row);
                        foreach (var callData in rowData)
                        {
                            string callDataText = !string.IsNullOrEmpty(callData) ? callData : string.Empty;

                            var cell = CreateTextCell(ColumnLetter(cellIdex++), rowIdex, callDataText);
                            row.AppendChild(cell);
                        }
                    }
                }

                workbookpart.Workbook.Save();
                document.Close();

                return stream.ToArray();
            }
            catch ( Exception ex)
            {
               
                throw ex;
            }
        }
    }
}
