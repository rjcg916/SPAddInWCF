using Microsoft.SharePoint.Client;
using System.Net;
using System.IO;
using System.Web;
using System.Collections.Generic;

namespace WebcorLessonsLearnedAppWeb.Utilities
{
    public class SharePointLessonsLearnedEmail
    {
        string formPath = @"/_layouts/15/";
        string libraryName = "/" + LessonsLearnedLists.LESSONSLEARNEDLISTURL + "/";

        LessonsLearnedEmail email;

        List<System.Net.Mail.Attachment> Attachments = new List<System.Net.Mail.Attachment>();

        string attachmentImagesHTML = "";
        string attachmentOthersHTML = "";


        public SharePointLessonsLearnedEmail(ListItem oListItem, string fromAddress, string[] publishAddresses, string[] reviewerAddresses)
        {
            ClientContext ctx = (ClientContext)oListItem.Context;

            LessonsLearnedType type = LessonsLearnedType.Published;

            string clickToLink = "";
            string clickToText = "";
            string subject = "";
            string comments = "";

            // build subject and links as appropriate for message


            //Get parent list for link to pages
            List parentList = oListItem.ParentList;
            ctx.Load(parentList, l => l.DefaultDisplayFormUrl, l => l.DefaultEditFormUrl);
            ctx.ExecuteQuery();


            switch (Lists.GetApprovalStatusField(oListItem))
            {
                case "Pending":
                    type = LessonsLearnedType.Review;

                    subject = "Lessons Learned - Item for Review";

                    clickToLink = oListItem.Context.Url + formPath + "approve.aspx?List={" + parentList.Id.ToString() + "}&ID=" + oListItem.Id.ToString();
                    clickToText = "Click here to approve / reject this item";

                    break;

                case "Approved":
                    type = LessonsLearnedType.Approved;

                    subject = "Lessons Learned - Item Approved";
                    comments = Lists.GetStringField(oListItem, "_ModerationComments");

                    clickToLink = oListItem.Context.Url + libraryName + "DispForm.aspx?ID=" + oListItem.Id.ToString();
                    clickToText = "Click here to view your item";

                    break;

                case "Rejected":
                    type = LessonsLearnedType.Rejected;

                    subject = "Lessons Learned - Item Rejected";
                    comments = Lists.GetStringField(oListItem, "_ModerationComments");


                    clickToLink = oListItem.Context.Url + libraryName + "EditForm.aspx?ID=" + oListItem.Id.ToString();

                    clickToText = "Click here to edit your item";


                    break;

                case "Published":
                    type = LessonsLearnedType.Published;

                    subject = "Lessons Learned";


                    clickToLink = oListItem.Context.Url + libraryName + "DispForm.aspx?ID=" + oListItem.Id.ToString();
                    clickToText = "Click here to view this item";

                    break;
            }

            //process E-mail attachments
            AddAttachments(oListItem);

            LessonsLearnedItem item = new LessonsLearnedItem();
            item.Title = oListItem["Title"].ToString();
            item.Project = Lists.GetLookupFieldSingle(oListItem, LessonsLearnedLists.PROJECTFIELDNAME);
            item.LessonType = Lists.GetStringField(oListItem, LessonsLearnedLists.LESSONTYPEFIELDNAME);
            item.Date = Lists.GetDateField(oListItem, LessonsLearnedLists.DATEFIELDNAME);
            item.Keywords = Lists.GetLookupFieldMulti(oListItem, LessonsLearnedLists.KEYWORDSFIELDNAME);
            item.Issue = Lists.GetStringField(oListItem, LessonsLearnedLists.ISSUEFIELDNAME);
            item.Resolution = Lists.GetStringField(oListItem, LessonsLearnedLists.RESOLUTIONFIELDNAME);
            item.Innovator = Lists.GetStringField(oListItem, LessonsLearnedLists.INNOVATORFIELDNAME);
            item.clickToLink = clickToLink;
            item.clickToText = clickToText;
            item.attachmentImagesHTML = attachmentImagesHTML;
            item.attachmentOthersHTML = attachmentOthersHTML;

            email = new LessonsLearnedEmail(type, item, comments);


            // add attachments 
            foreach (System.Net.Mail.Attachment a in Attachments)
            {
                email.Attachment = a;
            }

            // fill in Subject, To, From 

            email.Subject = subject;
            email.From = fromAddress;

            switch (type)
            {
                case LessonsLearnedType.Review:
                    {
                        foreach (string reviewerAddress in reviewerAddresses)
                        {
                            email.To = reviewerAddress;
                        }

                        break;
                    }

                case LessonsLearnedType.Published:
                    {
                        foreach (string publishAddress in publishAddresses)
                        {
                            email.Bcc = publishAddress;
                        }
                        email.To = Lists.GetCreatorEmail(oListItem, ctx);

                        break;
                    }
                case LessonsLearnedType.Approved:
                case LessonsLearnedType.Rejected:
                    {
                        email.To = Lists.GetCreatorEmail(oListItem, ctx);

                        break;
                    }

            }

        }

        public SharePointLessonsLearnedEmail(ListItem oListItem) :
            this(
                oListItem,
                Shared.GetAppSettingValue("fromAddress"),
                Shared.GetAppSettingValue("publishAddresses").Split(','),
                Shared.GetAppSettingValue("reviewerAddresses").Split(',')
                )
        {

        }

        public void addAttachment(string fileSharePointPath, byte[] fileDataArray, Microsoft.SharePoint.Client.Attachment fileMetadata)
        {
            System.Net.Mail.Attachment emailAttachment = new System.Net.Mail.Attachment(new MemoryStream(fileDataArray), fileMetadata.FileName, System.Web.MimeMapping.GetMimeMapping(fileMetadata.FileName));

            if (emailAttachment.ContentType.MediaType.Substring(0, emailAttachment.ContentType.MediaType.IndexOf("/")) == "image")
            {

                Attachments.Add(emailAttachment);
                attachmentImagesHTML += "<a href='" + fileSharePointPath + "'><img class='emailImage' style='display:block' width='75%'  src='cid:" + fileMetadata.FileName + "'></a><hr>";
            }
            else
            {
                attachmentOthersHTML += "<div><a href='" + fileSharePointPath + "'>" + fileMetadata.FileName + "</a></div>";
            }
        }
        public void AddAttachments(ListItem oListItem)
        {
            ClientContext ctx = (ClientContext)oListItem.Context;

            //Get list item attachments

            Microsoft.SharePoint.Client.AttachmentCollection attachments = oListItem.AttachmentFiles;
            ctx.Load(attachments);
            ctx.ExecuteQuery();

            //Get webclient for downloading attachment file from SharePoint site

            WebClient webclient = Shared.GetWebClient();

            //Loop through item attachments and add to email
            foreach (Microsoft.SharePoint.Client.Attachment attachment in attachments)
            {
                string fileAbsolutePath = Shared.GetRootUrl(ctx) + attachment.ServerRelativeUrl; //location of file on SharePoint site
                byte[] fileDataArray = webclient.DownloadData(fileAbsolutePath); //contents of file downloaded to byte array

                addAttachment(fileAbsolutePath, fileDataArray, attachment);
            }
        }
        public HtmlString ToHtmlString()
        {
            return email.ToHtmlString();
        }

        public void Send(string publishComment = "")
        {
            email.Send(publishComment);
        }
    }
}