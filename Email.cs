using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net.Mail;
using System.Net.Mime;
using System.Net.Http;
using System.Net;

namespace WebcorLessonsLearnedAppWeb.Utilities
{
    public class Email
    {

        static public SmtpClient GetSMTPClient(string smtpHost, int smtpPort, string smtpCredentialUser, string smtpCredentialPassword)
        {

            SmtpClient smtp = new SmtpClient();
            smtp.Host = smtpHost;
            smtp.Port = smtpPort;
            smtp.EnableSsl = true;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential(smtpCredentialUser, smtpCredentialPassword);
            return smtp;
        }

      
    }
    }