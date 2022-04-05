using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Crm.ExcelPlugins
{
    public class SLExcelReader : ExcelProvider
    {
        private string GetColumnName(string cellReference)
        {
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);

            return match.Value;
        }

        private int ConvertColumnNameToNumber(string columnName)
        {
            var alpha = new Regex("^[A-Z]+$");
            if (!alpha.IsMatch(columnName)) throw new ArgumentException();

            char[] colLetters = columnName.ToCharArray();
            Array.Reverse(colLetters);

            var convertedValue = 0;
            for (int i = 0; i < colLetters.Length; i++)
            {
                char letter = colLetters[i];
                // ASCII 'A' = 65
                int current = i == 0 ? letter - 65 : letter - 64;
                convertedValue += current * (int)Math.Pow(26, i);
            }

            return convertedValue;
        }

        private IEnumerator<Cell> GetExcelCellEnumerator(Row row)
        {
            int currentCount = 0;
            foreach (Cell cell in row.Descendants<Cell>())
            {
                string columnName = GetColumnName(cell.CellReference);

                int currentColumnIndex = ConvertColumnNameToNumber(columnName);

                for (; currentCount < currentColumnIndex; currentCount++)
                {
                    var emptycell = new Cell()
                    {
                        DataType = null,
                        CellValue = new CellValue(string.Empty)
                    };
                    yield return emptycell;
                }

                yield return cell;
                currentCount++;
            }
        }

        private string ReadExcelCell(Cell cell, WorkbookPart workbookPart)
        {
            var cellValue = cell.CellValue;
            var text = (cellValue == null) ? cell.InnerText : cellValue.Text;
            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            {
                text = workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(
                        Convert.ToInt32(cell.CellValue.Text)).InnerText;
            }

            return (text ?? string.Empty).Trim();
        }

        public SLExcelData ReadExcel(MemoryStream stream)
        {
            var data = new SLExcelData();

            // Check if the file is excel
            //if (file.ContentLength <= 0)
            //{
            //    data.Status.Message = "You uploaded an empty file";
            //    return data;
            //}

            //if (file.ContentType
            //    != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            //{
            //    data.Status.Message
            //        = "Please upload a valid excel file of version 2007 and above";
            //    return data;
            //}

            // Open the excel document
            WorkbookPart workbookPart; List<Row> rows;
            try
            {
                var document = SpreadsheetDocument.Open(stream, false);
                workbookPart = document.WorkbookPart;

                var sheets = workbookPart.Workbook.Descendants<Sheet>();
                var sheet = sheets.First();
                data.SheetName = sheet.Name;

                var workSheet = ((WorksheetPart)workbookPart
                    .GetPartById(sheet.Id)).Worksheet;
                var columns = workSheet.Descendants<Columns>().FirstOrDefault();
                data.ColumnConfigurations = columns;

                var sheetData = workSheet.Elements<SheetData>().First();
                rows = sheetData.Elements<Row>().ToList();
            }
            catch (Exception e)
            {
                data.Status.Message = "Unable to open the file";
                return data;
            }

            // Read the header
            if (rows.Count > 0)
            {
                var row = rows[0];
                var cellEnumerator = GetExcelCellEnumerator(row);
                while (cellEnumerator.MoveNext())
                {
                    var cell = cellEnumerator.Current;
                    var text = ReadExcelCell(cell, workbookPart).Trim();
                    data.Headers.Add(text);
                }
            }

            // Read the sheet data
            if (rows.Count > 1)
            {
                for (var i = 1; i < rows.Count; i++)
                {
                    var dataRow = new List<string>();
                    data.DataRows.Add(dataRow);
                    var row = rows[i];
                    var cellEnumerator = GetExcelCellEnumerator(row);
                    while (cellEnumerator.MoveNext())
                    {
                        var cell = cellEnumerator.Current;
                        var text = ReadExcelCell(cell, workbookPart).Trim();
                        dataRow.Add(text);
                    }
                }
            }

            return data;
        }

        public List<SLExcelData> CustomReadExcel(IOrganizationService crmService, MemoryStream stream)
        {
            List<SLExcelData> returnList = new List<SLExcelData>();

            SpreadsheetDocument document;
            try
            {
                #region Local Parameters
                List<ExcelColumn> excelColumns = new List<ExcelColumn>();
                
                List<string> excelColumnsString = new List<string>();
                
                #endregion

                #region Open the excel document
                WorkbookPart workbookPart; List<Row> rows;

                document = SpreadsheetDocument.Open(stream, false);
                workbookPart = document.WorkbookPart;

                var sheets = workbookPart.Workbook.Descendants<Sheet>();
                //var sheet = sheets.First();

                #endregion

                foreach (var sheet in sheets)
                {
                   
                        var excelData = new SLExcelData();
                        // data.SheetName = sheet.Name;

                        var workSheet = ((WorksheetPart)workbookPart
                                       .GetPartById(sheet.Id)).Worksheet;
                        var columns = workSheet.Descendants<Columns>().FirstOrDefault();
                        excelData.ColumnConfigurations = columns;

                        var sheetData = workSheet.Elements<SheetData>().First();
                        rows = sheetData.Elements<Row>().ToList();

                        #region Read the header
                        if (rows.Count > 0)
                        {
                            var row = rows[0];

                            foreach (Cell cell in row.Descendants<Cell>())
                            {
                                string columnName = GetColumnName(cell.CellReference);
                                var text = ReadExcelCell(cell, workbookPart).Trim();
                                int currentColumnIndex = ConvertColumnNameToNumber(columnName);

                                if (excelColumnsString.Contains(text))
                                    excelData.Headers.Add(text);

                                ExcelColumn c = new ExcelColumn();
                                c.Name = columnName;
                                c.Text = text;
                                c.Index = currentColumnIndex;

                                excelColumns.Add(c);
                            }
                        }

                        #endregion

                        #region Read the sheet data
                        if (rows.Count > 1)
                        {
                            for (var i = 1; i < rows.Count; i++)
                            {
                                var row = rows[i];

                                #region ListMember

                                var dataRow = new List<string>();

                                int currentCount = 0;
                                foreach (Cell cell in row.Descendants<Cell>())
                                {
                                    string columnName = GetColumnName(cell.CellReference);
                                    var text = ReadExcelCell(cell, workbookPart).Trim();
                                    int currentColumnIndex = ConvertColumnNameToNumber(columnName);

                                    if (excelColumns.Exists(a => a.Name.Equals(columnName)))
                                    {
                                        dataRow.Add(text);
                                    }

                                    currentCount++;
                                }

                                if (dataRow != null && dataRow.Count != 0 && !string.IsNullOrEmpty(dataRow[0]))
                                {
                                    excelData.DataRows.Add(dataRow);
                                }
                                #endregion
                            }
                        }
                        #endregion

                        returnList.Add(excelData);
                    
                }
            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink"))
                {
                  
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString() + Environment.NewLine +
                                "public ExcelResult CustomReadExcel(IOrganizationService crmService, MemoryStream stream)");
            }

            return returnList;
        }

       
        private static Uri FixUri(string brokenUri)
        {
            return new Uri("http://broken-link/");
        }

        private static void ProcessDocument(SpreadsheetDocument wDoc)
        {
            //var elementCount = wDoc.MainDocumentPart.Document.Descendants().Count();
            //Console.WriteLine(elementCount);
        }
    }

    public static class UriFixer
    {
        public static void FixInvalidUri(Stream fs, Func<string, Uri> invalidUriHandler)
        {
            XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
            //using (ZipArchive za = new ZipArchive(fs, ZipArchiveMode.Update))
            //{
            //    foreach (var entry in za.Entries.ToList())
            //    {
            //        if (!entry.Name.EndsWith(".rels"))
            //            continue;
            //        bool replaceEntry = false;
            //        XDocument entryXDoc = null;
            //        using (var entryStream = entry.Open())
            //        {
            //            try
            //            {
            //                entryXDoc = XDocument.Load(entryStream);
            //                if (entryXDoc.Root != null && entryXDoc.Root.Name.Namespace == relNs)
            //                {
            //                    var urisToCheck = entryXDoc
            //                        .Descendants(relNs + "Relationship")
            //                        .Where(r => r.Attribute("TargetMode") != null && (string)r.Attribute("TargetMode") == "External");
            //                    foreach (var rel in urisToCheck)
            //                    {
            //                        var target = (string)rel.Attribute("Target");
            //                        if (target != null)
            //                        {
            //                            try
            //                            {
            //                                Uri uri = new Uri(target);
            //                            }
            //                            catch (UriFormatException)
            //                            {
            //                                Uri newUri = invalidUriHandler(target);
            //                                rel.Attribute("Target").Value = newUri.ToString();
            //                                replaceEntry = true;
            //                            }
            //                        }
            //                    }
            //                }
            //            }
            //            catch (XmlException)
            //            {
            //                continue;
            //            }
            //        }
            //        if (replaceEntry)
            //        {
            //            var fullName = entry.FullName;
            //            entry.Delete();
            //            var newEntry = za.CreateEntry(fullName);
            //            using (StreamWriter writer = new StreamWriter(newEntry.Open()))
            //            using (XmlWriter xmlWriter = XmlWriter.Create(writer))
            //            {
            //                entryXDoc.WriteTo(xmlWriter);
            //            }
            //        }
            //    }
            //}
        }
    }
}
