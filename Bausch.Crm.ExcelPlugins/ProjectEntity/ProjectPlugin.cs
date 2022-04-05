using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Crm.ExcelPlugins.Model;
using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Query;
using System.Globalization;
using System.Collections;

namespace Crm.ExcelPlugins
{
    public class ProjectPlugin : Plugin
    {
        #region Global Parameters

        #endregion

        public ProjectPlugin()
            : base(typeof(ProjectPlugin))
        {
            base.RegisteredEvents.Add(new Tuple<int, string, string, Action<LocalPluginContext>>((int)Stages.PostOperation,
               "new_LoadExcelFile", "new_project", new Action<LocalPluginContext>(ExecuteImportProjectDetailsAction)));
        }

        protected void ExecuteImportProjectDetailsAction(LocalPluginContext localContext)
        {
            try
            {
                #region Plugin Tanımlamaları
                IPluginExecutionContext context = localContext.PluginExecutionContext;
                IOrganizationService crmService = localContext.OrganizationService;

                if (!context.InputParameters.Contains("Target"))
                    return;

                EntityReference activityRef = context.InputParameters["Target"] as EntityReference;

                string sponsorshipType = "";

                if (context.InputParameters != null)
                {
                    if (context.InputParameters["SponsorshipType"] != null)
                    {
                        sponsorshipType = context.InputParameters["SponsorshipType"].ToString();
                    }
                }
                #endregion

                if (activityRef != null && activityRef.Id != Guid.Empty)
                {

                    var details = GetProjectDetails(crmService, activityRef);
                    if (details != null && details.Count != 0)
                    {
                        foreach (var item in details)
                        {

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void ActivityUpdateStatus(IOrganizationService crmService, Guid activityid, int statecode, int statuscode)
        {
            Entity activity = new Entity("new_activity");
            activity.Id = activityid;
            activity["statecode"] = new OptionSetValue(statecode);
            activity["statuscode"] = new OptionSetValue(statuscode);
            crmService.Update(activity);
        }

        #region  Methods
        private List<ProjectDetail> GetProjectDetails(IOrganizationService crmService, EntityReference activityRef)
        {
            try
            {
                InitializeFileBlocksDownloadRequest ifbdRequest = new InitializeFileBlocksDownloadRequest
                {
                    Target = activityRef,
                    FileAttributeName = "new_excelfile"
                };

                InitializeFileBlocksDownloadResponse ifbdResponse = (InitializeFileBlocksDownloadResponse)crmService.Execute(ifbdRequest);

                DownloadBlockRequest dbRequest = new DownloadBlockRequest
                {
                    FileContinuationToken = ifbdResponse.FileContinuationToken
                };

                var resdbResponsep = (DownloadBlockResponse)crmService.Execute(dbRequest);

                var excelManager = new ExcelManager();

                return excelManager.GetProjectDetails(resdbResponsep.Data);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
    }
}
