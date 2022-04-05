using DocumentFormat.OpenXml.Packaging;
using Crm.ExcelPlugins.Model;
using System.Collections.Generic;
using System.IO;

namespace Crm.ExcelPlugins
{
    public class ExcelManager : ExcelProvider
    {
        
        public List<ProjectDetail> GetProjectDetails(byte[] file)
        {
            using (var templateStream = new MemoryStream())
            {
                templateStream.Write(file, 0, file.Length);
                SpreadsheetDocument document = SpreadsheetDocument.Open(templateStream, false);

                return GetProjectDetails(document);
            }
        }

        private List<ProjectDetail> GetProjectDetails(SpreadsheetDocument document)
        {
            return ExecuteReaderToList<ProjectDetail>("", document, CRM_ProjectDetail_Filler);
        }
       
    }
}
