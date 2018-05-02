using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebcorLessonsLearnedAppWeb.Utilities
{
    public static class LesssonsLearnedContentType
    {
        public static void CreateTermSets()
        {
            // create Projects Term Set
            // create Keywords Term Set


        }
        public static void CreateSiteColumns(Web web)
        {

            ClientContext context = (ClientContext)web.Context;

            context.Load(web, w => w.ContentTypes);

            FieldCollection fields = web.Fields;


            //Create Projects Column (using Term Set)
            //Create Keywords Column (using Term Set)

            //Create Lessons Learned Field
            string lessonTypeSchemaTextField = String.Format("<Field  Required='TRUE' Type='Choice' Format='Dropdown' Name = '{0}' StaticName = '{0}' DisplayName='Lesson Type'><Default> Lessons Learned</Default><CHOICES><CHOICE>Lessons Learned</CHOICE><CHOICE>Continuous Improvement</CHOICE></CHOICES></Field>", LESSONTYPEFIELDNAME);
            fields.AddFieldAsXml(lessonTypeSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);

            //Create Date-Time Field
            string dateSchemaTextField = String.Format("<Field  Required='TRUE' Type='DateTime' Format='DateOnly' Name='{0}' StaticName='{0}' DisplayName='{0}' />", DATEFIELDNAME);
            fields.AddFieldAsXml(dateSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);

            //Create Innovator Field
            string innovatorSchemaTextField = String.Format("<Field  Required='TRUE' Type='Text' Name='{0}' StaticName='{0}' DisplayName='{0}' />", INNOVATORFIELDNAME);
            fields.AddFieldAsXml(innovatorSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);

            //Create Issue Field
            string issueSchemaTextField = String.Format("<Field  Required='TRUE' Type='Note' NumLines='6' Name='{0}' StaticName='{0}' DisplayName='{0}' />", ISSUEFIELDNAME);
            fields.AddFieldAsXml(issueSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);
            //Create Resolution Field
            string resolutionSchemaTextField = String.Format("<Field Required='TRUE' Type='Note' NumLines='6' Name='{0}' StaticName='{0}' DisplayName='{0}' />", RESOLUTIONFIELDNAME);
            fields.AddFieldAsXml(resolutionSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);

            //Create Status Field
            string statusSchemaTextField = String.Format("<Field  Required='TRUE'  Type='Choice' Format='Dropdown' Name='{0}' StaticName='{0}' DisplayName='{0}'><Default>Draft</Default><CHOICES><CHOICE>Draft</CHOICE><CHOICE>Submit for Review</CHOICE></CHOICES></Field>", STATUSFIELDNAME);
            fields.AddFieldAsXml(statusSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);

        }

        public static void CreateContentType(Web web)
        {
            ClientContext context = (ClientContext)web.Context;

            context.Load(web, w => w.ContentTypes);            
            ContentTypeCollection contentTypes = web.ContentTypes;
            var itemContentType = contentTypes.GetById("0x01");

            ContentTypeCreationInformation lessonsLearnedInfo = new ContentTypeCreationInformation();
            lessonsLearnedInfo.ParentContentType = itemContentType;
            lessonsLearnedInfo.Name = "Lessons Learned";
            lessonsLearnedInfo.Group = "Webcor Custom";


            ContentType lessonsLearnedContentType = contentTypes.Add(lessonsLearnedInfo);

            //add Projects Column (using Term Set)
            //add Keywords Column (using Term Set)
            //add Lessons Learned Field
            //add Date-Time Field
            //add Innovator Field
            //add Issue Field
            //add Resolution Field
            //add Status Field

        }
    }
}