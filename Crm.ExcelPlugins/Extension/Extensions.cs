using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Xrm.Sdk;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace Crm.ExcelPlugins
{
    internal static class Extensions
    {
        internal static bool IsGuid(this string s)
        {
            if (s == null)
                throw new ArgumentNullException("s");

            Regex format = new Regex(
                "^[A-Fa-f0-9]{32}$|" +
                "^({|\\()?[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}(}|\\))?$|" +
                "^({)?[0xA-Fa-f0-9]{3,10}(, {0,1}[0xA-Fa-f0-9]{3,6}){2}, {0,1}({)([0xA-Fa-f0-9]{3,4}, {0,1}){7}[0xA-Fa-f0-9]{3,4}(}})$");
            Match match = format.Match(s);

            return match.Success;
        }

        internal static string ToStringFromGuid(this Guid id)
        {
            return id.ToString().Replace("{", "").Replace("}", "").ToUpper();
        }

        internal static string ToStringFormatGuid(this string id)
        {
            return id.Replace("{", "").Replace("}", "").ToUpper();
        }

        internal static string ToGetTime(this string time)
        {
            DateTime date = Convert.ToDateTime(time);
            string hour = date.TimeOfDay.Hours.ToString();
            string minute = date.TimeOfDay.Minutes.ToString();
            string timeString = hour + ":" + minute;
            return timeString;
        }

        #region Project Extensions
        internal static OptionSetValue ToUnitD365(this string unit)
        {
            var returnValue = new OptionSetValue();

            if (!string.IsNullOrEmpty(unit))
            {
                if (unit == "Private")
                {
                    returnValue = new OptionSetValue(1);
                }
                else if (unit == "Public")
                {
                    returnValue = new OptionSetValue(2);
                }
            }
            else
                returnValue = null;

            return returnValue;
        }

        internal static OptionSetValueCollection ToParticipantTypeD365(this string[] participantTypes)
        {
            var returnValue = new OptionSetValueCollection();

            if (participantTypes != null && participantTypes.Length != 0)
            {
                for (int i = 0; i < participantTypes.Length; i++)
                {
                    if (participantTypes[i] == "Participant")
                    {
                        returnValue.Add(new OptionSetValue(1));
                    }
                    else if (participantTypes[i] == "Speaker")
                    {
                        returnValue.Add(new OptionSetValue(2));
                    }
                    else if (participantTypes[i] == "Moderator")
                    {
                        returnValue.Add(new OptionSetValue(3));
                    }
                }
            }
            else
                returnValue = null;

            return returnValue;
        }
        internal static OptionSetValueCollection ToSponsorshipTypeD365(this string[] sponsorshipTypes)
        {
            var returnValue = new OptionSetValueCollection();

            if (sponsorshipTypes != null && sponsorshipTypes.Length != 0)
            {
                for (int i = 0; i < sponsorshipTypes.Length; i++)
                {
                    if (sponsorshipTypes[i] == "Registration")
                    {
                        returnValue.Add(new OptionSetValue(1));
                    }
                    else if (sponsorshipTypes[i] == "Accommodation")
                    {
                        returnValue.Add(new OptionSetValue(2));
                    }
                    else if (sponsorshipTypes[i] == "Transfer")
                    {
                        returnValue.Add(new OptionSetValue(3));
                    }
                    else if (sponsorshipTypes[i] == "Meals")
                    {
                        returnValue.Add(new OptionSetValue(4));
                    }
                    else if (sponsorshipTypes[i] == "Honorarium")
                    {
                        returnValue.Add(new OptionSetValue(5));
                    }
                    else if (sponsorshipTypes[i] == "Transportation")
                    {
                        returnValue.Add(new OptionSetValue(6));
                    }
                }
            }
            else
                returnValue = null;

            return returnValue;
        }

        #endregion
    }
}
