namespace Crm.ExcelPlugins.Model
{
    class Factories
    {
    }

    public class PluginMessages
    {
        #region Supported messages for custom entities
        internal static string Assign
        {
            get { return "Assign"; }
        }

        internal static string Execute
        {
            get { return "Execute"; }
        }

        internal static string Cancel // for salesorder
        {
            get { return "Cancel"; }
        }

        internal static string Create
        {
            get { return "Create"; }
        }

        internal static string Delete
        {
            get { return "Delete"; }
        }

        internal static string Merge
        {
            get { return "Merge"; }
        }

        internal static string GrantAccess
        {
            get { return "GrantAccess"; }
        }

        internal static string ModifyAccess
        {
            get { return "ModifyAccess"; }
        }

        internal static string Retrieve
        {
            get { return "Retrieve"; }
        }

        internal static string RetrieveMultiple
        {
            get { return "RetrieveMultiple"; }
        }

        internal static string RetrievePrincipalAccess
        {
            get { return "RetrievePrincipalAccess"; }
        }

        internal static string RetrieveSharedPrincipalsAndAccess
        {
            get { return "RetrieveSharedPrincipalsAndAccess"; }
        }

        internal static string RevokeAccess
        {
            get { return "RevokeAccess"; }
        }

        internal static string SetState
        {
            get { return "SetState"; }
        }

        internal static string SetStateDynamicEntity
        {
            get { return "SetStateDynamicEntity"; }
        }

        internal static string Update
        {
            get { return "Update"; }
        }

        internal static string Associate
        {
            get { return "Associate"; }
        }

        internal static string Disassociate
        {
            get { return "Disassociate"; }
        }

        internal static string QualifyLead
        {
            get { return "QualifyLead"; }
        }

        internal static string ConvertActivity
        {
            get { return "ConvertActivity"; }
        }
        #endregion
    }

    public class OptionSetList
    {
        public string Text { get; set; }
        public string Value { get; set; }
        public int Val { get; set; }
    }

    public class ProjectDetail
    {
        public string Fullname { get; set; }
        public string Speciality { get; set; }
        public string Institution { get; set; }
        public string InstitutionType { get; set; }
        public string Country { get; set; }
        public string City { get; set; }
        public string Nationality { get; set; }
        public string[] ParticipantType { get; set; }
        public string[] SponsorshipType { get; set; }
        public string Note { get; set; }
    }

    
}
