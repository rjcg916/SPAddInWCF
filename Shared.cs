using Microsoft.SharePoint.Client;
using System.Net;
using System.Security;
using System.Configuration;

namespace WebcorLessonsLearnedAppWeb.Utilities
{
    public class Shared
    {

        public static string GetAppSettingValue(string key)
        {
            string value = "";

            try
            {
                value = ConfigurationManager.AppSettings[key] ?? string.Empty;
            } catch { };

            return value;

        }

        public static string GetRootUrl(ClientContext ctx)
        {
            string root = "";
            Web web = ctx.Web;
            ctx.Load(web);
            ctx.ExecuteQuery();

            root = ctx.Web.Url;
            if (ctx.Web.ServerRelativeUrl != "")
            {
                root = root.Substring(0, root.IndexOf(ctx.Web.ServerRelativeUrl));
            }
            return root;
        }

        public static WebClient GetWebClient()
        {
            string webClientUsername = GetAppSettingValue("webClientUserName");
            string webClientSecurePassword = GetAppSettingValue("webClientSecurePassword");

            return GetWebClient(webClientUsername, webClientSecurePassword);            
        }

        public static WebClient GetWebClient(string webClientUsername, string webClientSecurePassword)
        {
            WebClient webClient = new WebClient();
            var securePassword = new SecureString();
            
            foreach (var c in webClientSecurePassword) { securePassword.AppendChar(c); }
            webClient.Credentials = new SharePointOnlineCredentials(webClientUsername, securePassword);
            webClient.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            return webClient;
        }

        public static string ProcessStringHTML(string html)
        {
            if (html == "")
            {
                return "(none)";
            }
            return html;
        }


    }
}
