using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Crm.ExcelPlugins
{
    class Enumaration
    {
    }

    internal enum Stages
    {
        PreValidation = 10,
        PreOpertion = 20,
        MainOperation = 30,
        PostOperation = 40
        // PostOpertionDeprecated = 50 only MS CRM 4.0
    }

    internal enum CrmLanguageCode
    {
        Turkish = 1055,
        English = 1033
    }

    internal enum ActivityStates
    {
        Active = 0,
        InActive = 1
    }

    internal enum ActivityStatus
    {
        Approved = 121220005,
        PendingApproval=121220006,
        InActive=2,
        Active=1
    }
}
