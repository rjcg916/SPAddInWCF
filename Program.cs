using System.Web;
using WebcorLessonsLearnedAppWeb.Utilities;
using System.Security;
using Microsoft.SharePoint.Client;
using System.ComponentModel;
using System;

namespace WebcorLessonsLearnedAppWebConsole
{
    class Program
    {


        static public string FetchEventReceivers()
        {

            string webUrl = Shared.GetAppSettingValue("webUrl");

            string userName = Shared.GetAppSettingValue("webClientUsername");
            SecureString securePassword = new SecureString();
            foreach (var c in Shared.GetAppSettingValue("webClientSecurePassword")) { securePassword.AppendChar(c); }

            using (var context = new ClientContext(webUrl))
            {

                context.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                context.Load(context.Web);

                Web web = context.Web;
                context.Load(web.Lists);

                ListCollection allLists = web.Lists;
                context.Load(allLists);

                List lessonsLearnedList = allLists.GetByTitle("Lessons Learned");
                context.Load(lessonsLearnedList);
                context.ExecuteQuery();

                return RemoteEventReceivers.DisplayEventReceivers(lessonsLearnedList, "LessonsLearnedRemoteEventReceiver");
                   

            }
        }

        static public void FetchLessonsLearnedItem()
        {

            string webUrl = Shared.GetAppSettingValue("webUrl");

            string userName = Shared.GetAppSettingValue("webClientUsername");
            SecureString securePassword = new SecureString();
            foreach (var c in Shared.GetAppSettingValue("webClientSecurePassword")) { securePassword.AppendChar(c); }

            string fromAddress = Shared.GetAppSettingValue("fromAddress");
            string[] reviewerAddresses = Shared.GetAppSettingValue("reviewerAddresses").Split(',');
            string[] publishAddresses = Shared.GetAppSettingValue("publishAddresses").Split(',');

            ListItem oListItem;

            using (var context = new ClientContext(webUrl))
            {

                context.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                context.Load(context.Web);

                Web web = context.Web;
                context.Load(web.Lists);

                ListCollection allLists = web.Lists;
                context.Load(allLists);

                List lessonsLearnedList = allLists.GetByTitle("Lessons Learned");
                context.Load(lessonsLearnedList);

                oListItem = lessonsLearnedList.GetItemById(1);
                context.Load(oListItem);
                context.ExecuteQuery();

                if (Lists.GetStringField(oListItem, "Status") != "Draft")
                {
                    SharePointLessonsLearnedEmail email = new SharePointLessonsLearnedEmail(oListItem, fromAddress, publishAddresses, reviewerAddresses);
                    email.ToHtmlString();
                    email.Send();
                }


            }
        }

        static public void SendLessonsLearnedItem()
        {

            LessonsLearnedItem item = new LessonsLearnedItem();
            item.Title = "Title";
            item.Project = "Project";
            item.LessonType = "Lesson Type";
            item.Date = "Date";
            item.Keywords = "Keyword1, Keyword2";
            item.Issue = "Issue";
            item.Resolution = "Resolution";
            item.Innovator = "Innovator";
            item.clickToLink = "https://webcor.sharepoint.com/sites/intranet";
            item.clickToText = "Click Here to Access Intranet";
            item.attachmentImagesHTML = "Images HTML";
            item.attachmentOthersHTML = "Others HTML";

            LessonsLearnedEmail email = new LessonsLearnedEmail(LessonsLearnedType.Review, item, "Reviewer Comment");

            email.Subject = "Email Subject";
            email.To = "admin@rjcgraham.onmicrosoft.com";
            email.To = "admin@rjcgraham.onmicrosoft.com";
            email.From = "admin@rjcgraham.onmicrosoft.com";
            email.Bcc = "admin@rjcgraham.onmicrosoft.com";
            email.Bcc = "admin@rjcgraham.onmicrosoft.com";

            HtmlString emailString = email.ToHtmlString();
            try
            {
                email.Send("This is an Important Message"); //send with comment
                email.Send(); // send w/o comment
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }

        }
    
        static public void CreateLists()
        {

            string webUrl = Shared.GetAppSettingValue("webUrl");

            string userName = Shared.GetAppSettingValue("webClientUsername");
            SecureString securePassword = new SecureString();
            foreach (var c in Shared.GetAppSettingValue("webClientSecurePassword")) { securePassword.AppendChar(c); }

            using (var context = new ClientContext(webUrl))
            {

                context.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                context.Load(context.Web);

                LessonsLearnedLists.CreateLists(context.Web);

            }
        }
        static public void DeleteLists()
        {

            string webUrl = Shared.GetAppSettingValue("webUrl");

            string userName = Shared.GetAppSettingValue("webClientUsername");
            SecureString securePassword = new SecureString();
            foreach (var c in Shared.GetAppSettingValue("webClientSecurePassword")) { securePassword.AppendChar(c); }

            using (var context = new ClientContext(webUrl))
            {

                context.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                context.Load(context.Web);
                Web web = context.Web;

                LessonsLearnedLists.DeleteLists(web);

            }
        }

        static public void AddCustomAction()
        {

            string webUrl = Shared.GetAppSettingValue("webUrl");

            string userName = Shared.GetAppSettingValue("webClientUsername");
            SecureString securePassword = new SecureString();
            foreach (var c in Shared.GetAppSettingValue("webClientSecurePassword")) { securePassword.AppendChar(c); }

            using (var context = new ClientContext(webUrl))
            {

                context.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                context.Load(context.Web);
                Web web = context.Web;

                var lessonslearned = web.Lists.GetByTitle(LessonsLearnedLists.LESSONSLEARNEDLISTTITLE);
                context.Load(lessonslearned);
                context.ExecuteQuery();

                CustomActions.AddCustomAction(lessonslearned, Shared.GetAppSettingValue("ActionUrl"));


            }
        }

        static void Main(string[] args)
        {
            //AddCustomAction();
//            DeleteLists(); // need to test list population format
//            CreateLists();
            //string er = FetchEventReceivers();
            //FetchLessonsLearnedItem();
//            SendLessonsLearnedItem(); // need to test publish comment
        }
    }
}
