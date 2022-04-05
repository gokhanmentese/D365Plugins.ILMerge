using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;
using System.Linq;

namespace Crm.ExcelPlugins
{
    public static class OrganizationBaseService
    {
       
        #region Share , UnShare
        public static void Share(this IOrganizationService serv, EntityReference principal, EntityReference target)
        {
            try
            {
                GrantAccessRequest grantAccessRequest = new GrantAccessRequest
                {
                    PrincipalAccess = new PrincipalAccess
                    {
                        Principal = principal,
                        AccessMask = AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.ShareAccess | AccessRights.AssignAccess | AccessRights.AppendAccess | AccessRights.AppendToAccess
                    },
                    Target = target
                };
                serv.Execute(grantAccessRequest);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void UnShare(this IOrganizationService serv, EntityReference revokee, EntityReference target)
        {
            try
            {
                RevokeAccessRequest revokeAccessRequest = new RevokeAccessRequest
                {
                    Revokee = revokee,
                    Target = target
                };
                RevokeAccessResponse revokeaccessresponse = (RevokeAccessResponse)serv.Execute(revokeAccessRequest);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void RetrieveAndDisplayPrincipalAccess(IOrganizationService service, string entityname, Guid entityid, string principalname, Guid principalid, String additionalIdentifier)
        {
            try
            {
                var principalAccessReq = new RetrievePrincipalAccessRequest
                {
                    Principal = new EntityReference(principalname, principalid),
                    Target = new EntityReference(entityname, entityid)
                };
                var principalAccessRes = (RetrievePrincipalAccessResponse)service.Execute(principalAccessReq);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static RetrieveSharedPrincipalsAndAccessResponse RetrieveAndDisplayAccess(IOrganizationService service, string entityname, Guid entityid)
        {
            try
            {
                var accessRequest = new RetrieveSharedPrincipalsAndAccessRequest
                {
                    Target = new EntityReference(entityname, entityid)
                };

                // The RetrieveSharedPrincipalsAndAccessResponse returns an entity reference
                // that has a LogicalName of "user" when returning access information for a
                // "team."
                return (RetrieveSharedPrincipalsAndAccessResponse)service.Execute(accessRequest);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static bool IsIndustryShared(IOrganizationService service, string entityname, Guid entityid, Guid teamid)
        {
            try
            {
                var isReturn = false;
                var accessResponse = RetrieveAndDisplayAccess(service, entityname, entityid);
                if (accessResponse != null && accessResponse.PrincipalAccesses != null && accessResponse.PrincipalAccesses.Length != 0)
                {
                    isReturn = accessResponse.PrincipalAccesses.Where(a => a.Principal.Id == teamid).Any();
                    //foreach (var principalAccess in accessResponse.PrincipalAccesses)
                    //{
                    //    if (principalAccess.Principal.LogicalName.ToLower().Equals("team"))
                    //    {
                    //        Team team = (Team)service.Retrieve("team", principalAccess.Principal.Id, new ColumnSet("teamid", "name"));
                    //        if (team != null && team.Id != Guid.Empty)
                    //        {
                    //            if (team.Name.ToLower().Equals(teamname.ToLower()))
                    //            {
                    //                isReturn = true; break;
                    //            }
                    //        }
                    //    }
                    //}
                }
                return isReturn;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        public static bool SetStateRequestFunc(this IOrganizationService crmService, EntityReference entityMoniker, int state, int statusreason)
        {
            try
            {
                SetStateRequest open = new SetStateRequest
                {
                    EntityMoniker = entityMoniker,
                    State = new OptionSetValue(state),
                    Status = new OptionSetValue(statusreason)
                };

                SetStateResponse resp = (SetStateResponse)crmService.Execute(open);

                return true; 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void AssignTeamorUser(this IOrganizationService crmService, EntityReference assignee, EntityReference target)
        {
            try
            {
                AssignRequest assign = new AssignRequest
                {
                    Assignee = assignee,
                    Target = target
                };
                crmService.Execute(assign);
            }
            catch (Exception ex)
            {
                
            }
        }
    }
}
