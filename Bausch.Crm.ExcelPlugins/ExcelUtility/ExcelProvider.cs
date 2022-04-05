
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Crm.ExcelPlugins.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Crm.ExcelPlugins
{
    public class ExcelProvider : BaseExcelProviderFiller
    {
        protected delegate T FillerToList<T>(Row dr, WorkbookPart w);

        protected List<K> ExecuteReaderToList<K>(string sheetName, SpreadsheetDocument document, FillerToList<K> filler)
        {

            List<K> ret = new List<K>();
            List<Row> rows = null;

            try
            {
                WorkbookPart workbookPart = document.WorkbookPart;

                var sheets = workbookPart.Workbook.Descendants<Sheet>();

                var workSheet = ((WorksheetPart)workbookPart
                               .GetPartById(sheets.FirstOrDefault().Id)).Worksheet;
                var columns = workSheet.Descendants<Columns>().FirstOrDefault();

                var sheetData = workSheet.Elements<SheetData>().First();
                rows = sheetData.Elements<Row>().ToList();

                for (int i = 1; i < rows.Count; i++)
                {
                    ret.Add(filler(rows[i], workbookPart));

                    //if (i == 50)
                    //    break;
                }

                return ret;
            }
            catch (Exception ex)
            {
                throw new Exception("ExecuteReadertoList", ex);
            }
            finally
            {

            }
        }
    }

    public class BaseExcelProviderFiller
    {
    
        public ProjectDetail CRM_ProjectDetail_Filler(Row dr, WorkbookPart w)
        {
            try
            {
                var tmp = new ProjectDetail();

                tmp.Fullname = dr.rowToString(w, "D");
                tmp.Institution = dr.rowToString(w, "E");
                tmp.InstitutionType = dr.rowToString(w, "F");
                tmp.Speciality = dr.rowToString(w, "G");
                tmp.Country = dr.rowToString(w, "H");
                tmp.City = dr.rowToString(w, "I");
                tmp.Nationality = dr.rowToString(w, "J");

                string participantType = dr.rowToString(w, "K");
                if (!string.IsNullOrEmpty(participantType))
                {
                    tmp.ParticipantType = participantType.Split(';');
                }

                string sponsorshipType = dr.rowToString(w, "L");
                if (!string.IsNullOrEmpty(sponsorshipType))
                {
                    tmp.SponsorshipType = sponsorshipType.Split(';');
                }

                tmp.Note = dr.rowToString(w, "M");

                return tmp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
