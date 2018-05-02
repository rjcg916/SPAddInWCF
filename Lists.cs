using System;
using Microsoft.SharePoint.Client;

namespace WebcorLessonsLearnedAppWeb.Utilities
{
    public class Lists
    {

        public static string GetCreatorEmail(ListItem oListItem, ClientContext ctx)
        {
            string email = "";

            FieldUserValue oValue = oListItem["Author"] as FieldUserValue;
            String strAuthorName = oValue.LookupValue;
            if (oValue.LookupId != -1)
            {
                User oUser = ctx.Web.EnsureUser(strAuthorName);
                ctx.Load(oUser);
                ctx.ExecuteQuery();
                email = oUser.Email;
            }

            return email;
        }

        public static string GetCreatorName(ListItem oListItem)
        {
            FieldUserValue oValue = oListItem["Author"] as FieldUserValue;

            return oValue.LookupValue;
        }

        public static string GetLookupFieldSingle(ListItem oListItem, string fieldName)
        {
            string value = "";

            try
            {
                FieldLookupValue field = (FieldLookupValue)oListItem[fieldName];
                value = field.LookupValue;
            }
            catch { }

            return value;
        }

        public static string GetLookupFieldMulti(ListItem oListItem, string fieldName)
        {
            string fieldVal = "";

            try
            {
                FieldLookupValue[] field = (FieldLookupValue[])oListItem[fieldName];

                foreach (FieldLookupValue lookupVal in field)
                {
                    fieldVal += lookupVal.LookupValue + ", ";
                }
                if (fieldVal != "")
                {
                    fieldVal = fieldVal.Substring(0, fieldVal.Length - 2);
                }
                else
                {
                    fieldVal = "(none)";
                }
            }
            catch { }

            return fieldVal;
        }

        public static string GetDateField(ListItem oListItem, string fieldName)
        {
            string value = "";

            try
            {
                value = oListItem[fieldName].ToString().Substring(0, oListItem[fieldName].ToString().IndexOf(" "));
            }
            catch { }

            return value;
        }

        public static string GetStringField(ListItem oListItem, string fieldName)
        {
            string field = "";
            if (oListItem[fieldName] != null)
            {
                field = oListItem[fieldName].ToString();
            }
            return field;
        }

        public static string GetApprovalStatusField(ListItem oListItem)
        {
            string status = "";
            switch (oListItem["_ModerationStatus"].ToString())
            {
                case "0":
                    status = "Approved";
                    break;
                case "1":
                    status = "Rejected";
                    break;
                case "2":
                    status = "Pending";
                    break;

            }
            return status;
        }

    }
}