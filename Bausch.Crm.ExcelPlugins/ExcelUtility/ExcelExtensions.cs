using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Crm.ExcelPlugins
{
    public static class ExcelExtensions
    {
        public static string rowToString(this Row row, WorkbookPart workbookPart, string kolonAdi)
        {
            string returnString = string.Empty;

            foreach (Cell cell in row.Descendants<Cell>())
            {
                string columnName = ExcelExtensions.GetColumnName(cell.CellReference);
               
                if (kolonAdi == columnName)
                {
                    var text = ReadExcelCell(cell, workbookPart).Trim();
                    returnString = text;
                    break;
                }
            }


            return returnString;
        }

        public static bool? rowToBoolean(this Row row, WorkbookPart workbookPart, string kolonAdi)
        {
            bool? returnDonecek = null;

            foreach (Cell cell in row.Descendants<Cell>())
            {
                string columnName = ExcelExtensions.GetColumnName(cell.CellReference);
               
                if (kolonAdi == columnName)
                {
                    var text = ReadExcelCell(cell, workbookPart).Trim();
                    if (text == "Evet")
                        returnDonecek = true;
                    else if (text == "Hayır")
                        returnDonecek = false;

                    break;
                }
            }

            return returnDonecek;
        }

        private static string GetColumnName(string cellReference)
        {
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);

            return match.Value;
        }

        private static string ReadExcelCell(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.CellValue != null)
            {
                var cellValue = cell.CellValue;
                var text = (cellValue == null) ? cell.InnerText : cellValue.Text;
                if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                {
                    text = workbookPart.SharedStringTablePart.SharedStringTable
                        .Elements<SharedStringItem>().ElementAt(Convert.ToInt32(cell.CellValue.Text)).InnerText;
                }
                else
                {
                    if (cell.StyleIndex != null)
                    {
                        var cellFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[
                    int.Parse(cell.StyleIndex.InnerText)] as CellFormat;
                        // only focus on Date
                        if (cellFormat != null)
                        {
                            var dateFormat = GetDateTimeFormat(cellFormat.NumberFormatId);
                            if (!string.IsNullOrEmpty(dateFormat))
                                return DateTime.FromOADate(double.Parse(text)).ToString();
                        }

                    }
                    return text;
                }

                return (text ?? string.Empty).Trim();
            }
            else
                return string.Empty;
        }

        private static string GetDateTimeFormat(uint numberFormatId)
        {
            return DateFormatDictionary.ContainsKey(numberFormatId) ? DateFormatDictionary[numberFormatId] : string.Empty;
        }

        //// https://msdn.microsoft.com/en-GB/library/documentformat.openxml.spreadsheet.numberingformat(v=office.14).aspx
        private static readonly Dictionary<uint, string> DateFormatDictionary = new Dictionary<uint, string>()
        {
            [14] = "dd/MM/yyyy",
            [15] = "d-MMM-yy",
            [16] = "d-MMM",
            [17] = "MMM-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "M/d/yy h:mm",
            [30] = "M/d/yy",
            [34] = "yyyy-MM-dd",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mmss.0",
            [51] = "MM-dd",
            [52] = "yyyy-MM-dd",
            [53] = "yyyy-MM-dd",
            [55] = "yyyy-MM-dd",
            [56] = "yyyy-MM-dd",
            [58] = "MM-dd",
            [165] = "M/d/yy",
            [166] = "dd MMMM yyyy",
            [167] = "dd/MM/yyyy",
            [168] = "dd/MM/yy",
            [169] = "d.M.yy",
            [170] = "yyyy-MM-dd",
            [171] = "dd MMMM yyyy",
            [172] = "d MMMM yyyy",
            [173] = "M/d",
            [174] = "M/d/yy",
            [175] = "MM/dd/yy",
            [176] = "d-MMM",
            [177] = "d-MMM-yy",
            [178] = "dd-MMM-yy",
            [179] = "MMM-yy",
            [180] = "MMMM-yy",
            [181] = "MMMM d, yyyy",
            [182] = "M/d/yy hh:mm t",
            [183] = "M/d/y HH:mm",
            [184] = "MMM",
            [185] = "MMM-dd",
            [186] = "M/d/yyyy",
            [187] = "d-MMM-yyyy"
        };
    }
}
