using System.Web;
using System.Net;
using System.Net.Mail;
using System.ComponentModel;
using System;

namespace WebcorLessonsLearnedAppWeb.Utilities
{
    public enum LessonsLearnedType
    {
        Review, Approved, Rejected, Published
    }

    public struct LessonsLearnedItem
    {
        public string Title;
        public string Project;
        public string LessonType;
        public string Date;
        public string Keywords;
        public string Issue;
        public string Resolution;
        public string Innovator;
        public string clickToLink;
        public string clickToText;
        public string attachmentImagesHTML;
        public string attachmentOthersHTML;
    }

    public class LessonsLearnedEmail
    {

        string _subject;
        public string Subject
        {
            set { _subject = value; }
        }

        string _From;
        public string From
        {
            set { _From = value; }
        }

        MailAddressCollection _To;
        public string To
        {
            set { _To.Add(value); }
        }

        MailAddressCollection _Bcc;
        public string Bcc
        {
            set { _Bcc.Add(value); }
        }

        string _Body;

        System.Collections.Generic.List<System.Net.Mail.Attachment> _Attachments;
        public System.Net.Mail.Attachment Attachment
        {
            set { _Attachments.Add(value); }
        }

        public LessonsLearnedEmail(LessonsLearnedType type, LessonsLearnedItem item, string comments = "")
        {
            _Bcc = new MailAddressCollection();
            _To = new MailAddressCollection();
            _Attachments = new System.Collections.Generic.List<Attachment>();

            InitializeBodyWithHead();
            BuildBodyHeaderText(type, item.Title, comments);
            BuildSharedBodyText(item);

        }

        private void InitializeBodyWithHead()
        {

            _Body = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'><html xmlns='http://www.w3.org/1999/xhtml'>" +
                              "<head>" +
                              "<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />" +
                              "<style type='text/css'>" +
                                  "body {margin:0; padding:0; min-width:100% !important; font-family:Calibri; font-size:12pt}" +
                                  ".content {width: 100%; max-width: 600px;}" +
                                  ".mainTable {width:100%; border-collapse:collapse; border:1px solid #efefef}" +
                                  ".mainTable td{border-bottom:1px solid #efefef;border-right:1px solid #efefef;padding:10px 15px}" +
                                  "h3{font-weight:500;color:#262626;}" +
                                  "h2{font-weight:600;color:#262626;font-size:16pt}" +
                                  "h1{font-size:18pt;font-weight:bold;background-color:#45865b;padding:5px;text-align:center;color:#ffffff}" +
                              "</style>" +
                          "</head>" +
                          "<body>";
        }



        private void BuildBodyHeaderText(LessonsLearnedType type, string title, string comments = "")
        {
            switch (type)
            {
                case LessonsLearnedType.Review:
                    _Body += "<p><span style='font-weight:bold'>" + title + "</span> is being submitted for your approval, please review and click the link at the bottom to approve or reject</p>";
                    break;
                case LessonsLearnedType.Approved:
                    _Body += "<p><span style='font-weight:bold'>" + title + "</span> has been approved by OB Review. Please look below for any comments</p>" +
                                  "<table><tr><td><h3>Comments:</h3></td><td><p>" + Shared.ProcessStringHTML(comments) + "</p></td></tr></table><hr>";
                    break;
                case LessonsLearnedType.Rejected:
                    _Body += "<p><span style='font-weight:bold'>" + title + "</span> has been rejected by OB Review. Please look below for any comments</p>" +
                                  "<table><tr><td><h3>Comments:</h3></td><td><p>" + Shared.ProcessStringHTML(comments) + "</p></td></tr></table><hr>";
                    break;
                case LessonsLearnedType.Published:
                    _Body += "<p>This item <span style='font-weight:bold'>" + title + "</span> has been published by OB Review</p>";
                    break;
            }
        }

        private void BuildSharedBodyText(LessonsLearnedItem item)
        {
            _Body += "<div class='content'>" +
                "<table class='mainTable'>" +
                    "<tr>" +
                        "<th colspan='2'>" +
                            "<h1>" +
                                "Lessons Learned" +
                             "</h1>" +
                         "</th>" +
                    "</tr>" +
                    "<tr>" +
                        "<td>" +
                            "<h3>Title</h3>" +
                        "</td>" +
                        "<td>" +
                            "<p>" + item.Title + "</p>" +
                        "</td>" +
                    "</tr>" +
                    "<tr>" +
                        "<td>" +
                            "<h3>Project</h3>" +
                        "</td>" +
                        "<td>" +
                            "<p>" + item.Project + "</p>" +
                        "</td>" +
                    "</tr>" +
                    "<tr>" +
                        "<td>" +
                            "<h3>Lesson Type</h3>" +
                        "</td>" +
                        "<td>" +
                            "<p>" + item.LessonType + "</p>" +
                        "</td>" +
                    "</tr>" +
                    "<tr>" +
                        "<td>" +
                            "<h3>Date</h3>" +
                        "</td>" +
                        "<td>" +
                            "<p>" + item.Date + "</p>" +
                        "</td>" +
                    "</tr>" +
                    "<tr>" +
                        "<td>" +
                            "<h3>Keyword(s)</h3>" +
                        "</td>" +
                        "<td>" +
                            "<p>" + item.Keywords + "</p>" +
                        "</td>" +
                    "</tr>" +
                    "<tr>" +
                        "<td>" +
                            "<h3>Issue</h3>" +
                        "</td>" +
                        "<td>" +
                            "<p>" + item.Issue + "</p>" +
                        "</td>" +
                    "</tr>" +
                    "<tr>" +
                        "<td>" +
                            "<h3>Resolution</h3>" +
                        "</td>" +
                        "<td>" +
                            "<p>" + item.Resolution + "</p>" +
                        "</td>" +
                    "</tr>" +
                    "<tr>" +
                        "<td>" +
                            "<h3>Innovator</h3>" +
                        "</td>" +
                        "<td>" +
                            "<p>" + item.Innovator + "</p>" +
                        "</td>" +
                    "</tr>" +
                    "<tr>" +
                        "<td>" +
                            "<h3>Attachment(s)</h3>" +
                        "</td>" +
                        "<td>" +
                            "<p>" + Shared.ProcessStringHTML(item.attachmentOthersHTML) + "</p>" +
                         "</td>" +
                    "</tr>" +
                    "<tr>" +
                        "<td colspan='2' style='text-align:center'>" +
                            "<a href='" + item.clickToLink + "' target='_blank'>" + item.clickToText + "</a> " +
                        "</td>" +
                    "</tr>" +
                "</table>" +
                "<h2>Images</h2>" + Shared.ProcessStringHTML(item.attachmentImagesHTML) +
                "</div><p><i>This is an auto-generated email, please do not reply</i></p></body></html>";
        }

        public HtmlString ToHtmlString()
        {
            string email = "";

            email += "<br/>From: " + _From;
            email += " To: ";
            foreach (MailAddress recipient in _To)
            {
                email += recipient + " ";
            }

            email += "<br /> Bcc: ";
            foreach (MailAddress recipient in _Bcc)
            {
                email += recipient + " ";
            }

            email += "<br /> Subject: " + _subject;

            email += "<br /><br />" + _Body;

            return new HtmlString(email);

        }


        private MailMessage CreateMessage(string publishComment = "")
        {
            MailMessage Email = new MailMessage();

            Email.Subject = _subject;

            Email.From = new MailAddress(_From);

            foreach (MailAddress recipient in _To)
            {
                Email.To.Add(recipient);
            }

            foreach (MailAddress recipient in _Bcc)
            {
                Email.Bcc.Add(recipient);
            }


            Email.IsBodyHtml = true;

            if (!String.IsNullOrWhiteSpace(publishComment))
            {
                Email.Body = "<p><strong>" + publishComment + "</strong></p>" + _Body;
            }
            else
                Email.Body = _Body;

            foreach (Attachment a in _Attachments)
            {
                Email.Attachments.Add(a);
            }

            return Email;
        }

        private SmtpClient GetMailClient()
        {
            string smtpHost = Shared.GetAppSettingValue("smtpHost");
            int smtpPort = int.Parse(Shared.GetAppSettingValue("smtpPort"));
            string smtpCredentialUser = Shared.GetAppSettingValue("smtpCredentialUser");
            string smtpCredentialPassword = Shared.GetAppSettingValue("smtpCredentialPassword");

            return Utilities.Email.GetSMTPClient(smtpHost, smtpPort, smtpCredentialUser, smtpCredentialPassword);

        }

        public void Send(string publishComment = "")
        {

            MailMessage msg = CreateMessage(publishComment);
            SmtpClient mailClient = GetMailClient();
            mailClient.Send(msg);
        }

    }

}