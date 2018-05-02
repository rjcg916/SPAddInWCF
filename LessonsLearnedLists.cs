using Microsoft.SharePoint.Client;
using System;

namespace WebcorLessonsLearnedAppWeb.Utilities
{
    public class LessonsLearnedLists
    {
        static string PROJECTSLISTURL = "Projects";
        static string PROJECTSLISTTITLE = PROJECTSLISTURL;
        public static string LESSONSLEARNEDLISTURL = "LessonsLearned";
        public static string LESSONSLEARNEDLISTTITLE = "Lessons Learned";
        static string KEYWORDSLISTURL = "LessonsLearnedKeywords";
        static string KEYWORDSLISTTITLE = "Lessons Learned Keywords";
        static string EVENTRECEIVERNAME = "LessonsLearnedRemoteEventReceiver";

        public static string PROJECTFIELDNAME = "Project";
        public static string LESSONTYPEFIELDNAME = "LessonType";
        public static string DATEFIELDNAME = "Date";
        public static string KEYWORDSFIELDNAME = "Keywords";
        public static string ISSUEFIELDNAME = "Issue";
        public static string RESOLUTIONFIELDNAME = "Resolution";
        public static string INNOVATORFIELDNAME = "Innovator";
        public static string STATUSFIELDNAME = "Status";

        private static void PopulateProjects(List projectsList)
        {
            string[] projects = new string[] {
    "Project 1",
    "zOther"
            };

            foreach (string project in projects)
            {
                ListItemCreationInformation keywordInfo = new ListItemCreationInformation();
                ListItem keywordItem = projectsList.AddItem(keywordInfo);
                keywordItem["Title"] = project;
                keywordItem.Update();
            }

            ClientContext context = (ClientContext)projectsList.Context;
            context.ExecuteQuery();

        }

        private static void PopulateKeywords(List keywordsList)
        {
            string[] keywords =  new string[] {
    "Columns",
    "Contractural",
    "Curbs",
    "Deck Shoring-Flyers",
    "Deck Shoring-Handset",
    "Demo/Chipping",
    "Design",
    "Dewatering",
    "Earth Retention(Shoring)",
    "Earthwork",
    "Embeds",
    "Equipment",
    "FormWork",
    "Housekeeping",
    "Logistics",
    "Management Process",
    "Mat",
    "Material",
    "Office",
    "P&F",
    "Permits/Inspection",
    "Pile Caps/Spread Footing",
    "Post Tensioning",
    "Pouring",
    "Pumping",
    "Rebar",
    "Reshore",
    "Safety",
    "Schedule Improvement",
    "Screens",
    "Shortcrete",
    "SOG",
    "Stairs",
    "Tools",
    "Truck-Load/Offload",
    "Walls",
    "Walls-Climber",
    "zOther"
            };

          foreach (string keyword in keywords) {
                ListItemCreationInformation keywordInfo = new ListItemCreationInformation();
                ListItem keywordItem = keywordsList.AddItem(keywordInfo);
                keywordItem["Title"] = keyword;
                keywordItem.Update();
            }

            ClientContext context = (ClientContext) keywordsList.Context;
            context.ExecuteQuery();

        }

        public static void CreateLists(Web web)
        {

            ClientContext context = (ClientContext)web.Context;

            // create lookup lists

            // projects list
            ListCreationInformation projectsParameters = new ListCreationInformation();
            projectsParameters.Url = PROJECTSLISTURL;
            projectsParameters.Title = PROJECTSLISTURL;
            projectsParameters.Description = "Contains a list of relevant projects";
            projectsParameters.TemplateType = (int)ListTemplateType.GenericList;

            List projects = web.Lists.Add(projectsParameters);
            context.Load(projects);
            projects.EnableAttachments = false;
            projects.Update();

            context.ExecuteQuery();


            //initialize known projects
            PopulateProjects(projects);


            // keywords list
            ListCreationInformation keywordsParameters = new ListCreationInformation();
            keywordsParameters.Url = KEYWORDSLISTURL;
            keywordsParameters.Title = "Lessons Learned Keywords";
            keywordsParameters.Description = "Contains a list of Lessons Learned keywords";
            keywordsParameters.TemplateType = (int)ListTemplateType.GenericList;
            List keywords = web.Lists.Add(keywordsParameters);
            keywords.EnableAttachments = false;
            keywords.Update();
            context.ExecuteQuery();

            //initialize known keywords
            PopulateKeywords(keywords);


            // create main list
            ListCreationInformation lessonslearnedParameters = new ListCreationInformation();
            lessonslearnedParameters.Url = LESSONSLEARNEDLISTURL;
            lessonslearnedParameters.Title = "Lessons Learned";
            lessonslearnedParameters.Description = "Lessons Learned compiled from past projects";
            lessonslearnedParameters.TemplateType = (int)ListTemplateType.GenericList;
            List lessonslearned = web.Lists.Add(lessonslearnedParameters);
            context.Load(lessonslearned);

            lessonslearned.EnableModeration = true;
            lessonslearned.DraftVersionVisibility = DraftVisibilityType.Author;
            lessonslearned.ReadSecurity = 1; //All users can read item
            lessonslearned.WriteSecurity = 2; //Users can only modify items they create
            lessonslearned.Update();
            context.ExecuteQuery();


            //// add other fields
            string lessonTypeSchemaTextField = String.Format("<Field  Required='TRUE' Type='Choice' Format='Dropdown' Name = '{0}' StaticName = '{0}' DisplayName='Lesson Type'><Default> Lessons Learned</Default><CHOICES><CHOICE>Lessons Learned</CHOICE><CHOICE>Continuous Improvement</CHOICE></CHOICES></Field>", LESSONTYPEFIELDNAME);
            lessonslearned.Fields.AddFieldAsXml(lessonTypeSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);


            string dateSchemaTextField = String.Format("<Field  Required='TRUE' Type='DateTime' Format='DateOnly' Name='{0}' StaticName='{0}' DisplayName='{0}' />", DATEFIELDNAME);
            lessonslearned.Fields.AddFieldAsXml(dateSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);


            string innovatorSchemaTextField = String.Format("<Field  Required='TRUE' Type='Text' Name='{0}' StaticName='{0}' DisplayName='{0}' />",INNOVATORFIELDNAME);
            lessonslearned.Fields.AddFieldAsXml(innovatorSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);


            string issueSchemaTextField = String.Format("<Field  Required='TRUE' Type='Note' NumLines='6' Name='{0}' StaticName='{0}' DisplayName='{0}' />", ISSUEFIELDNAME);
            lessonslearned.Fields.AddFieldAsXml(issueSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);


            string resolutionSchemaTextField = String.Format("<Field Required='TRUE' Type='Note' NumLines='6' Name='{0}' StaticName='{0}' DisplayName='{0}' />", RESOLUTIONFIELDNAME);
            lessonslearned.Fields.AddFieldAsXml(resolutionSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);


            string statusSchemaTextField = String.Format("<Field  Required='TRUE'  Type='Choice' Format='Dropdown' Name='{0}' StaticName='{0}' DisplayName='{0}'><Default>Draft</Default><CHOICES><CHOICE>Draft</CHOICE><CHOICE>Submit for Review</CHOICE></CHOICES></Field>", STATUSFIELDNAME );
            lessonslearned.Fields.AddFieldAsXml(statusSchemaTextField, true, AddFieldOptions.AddFieldInternalNameHint);

            context.ExecuteQuery();

            // add lookup fields  Project, Keywords

            context.Load(projects, p => p.Id);
            context.Load(keywords, k => k.Id);
            context.ExecuteQuery();

            string projectSchemaLookupField = String.Format("<Field  Required='TRUE' Type='Lookup' Name='{0}' DisplayName='{0}' List='{" + projects.Id + "}' ShowField='Title' />", PROJECTFIELDNAME);
            lessonslearned.Fields.AddFieldAsXml(projectSchemaLookupField, true, AddFieldOptions.AddFieldInternalNameHint);

            string keywordsSchemaLookupField = String.Format("<Field  Required='TRUE'  Type='LookupMulti' Mult='TRUE' Name='{0}' DisplayName='{0}' List='{" + keywords.Id + "}' ShowField='Title' />", KEYWORDSFIELDNAME);
            lessonslearned.Fields.AddFieldAsXml(keywordsSchemaLookupField, true, AddFieldOptions.AddFieldInternalNameHint);

            context.ExecuteQuery();


            // add custom action
            string actionUrl = Shared.GetAppSettingValue("ActionUrl");

            CustomActions.AddCustomAction(lessonslearned, actionUrl);

            // add event receivers

            string AppWebRemoteServiceUrl = Shared.GetAppSettingValue("AppWebRemoteServiceUrl");
            RemoteEventReceivers.DeleteEventReceivers(lessonslearned, EVENTRECEIVERNAME);
            RemoteEventReceivers.AddEventReceivers(lessonslearned, EVENTRECEIVERNAME, AppWebRemoteServiceUrl);

        }

        public static void DeleteLists(Web web)
        {
            ClientRuntimeContext context = web.Context;

            try
            {
                var lessonslearned = web.Lists.GetByTitle(LESSONSLEARNEDLISTTITLE);
                context.Load(lessonslearned);
                lessonslearned.DeleteObject();
                context.ExecuteQuery();
            }
            catch { }



            try
            {
                var projects = web.Lists.GetByTitle(PROJECTSLISTTITLE);
                context.Load(projects);
                projects.DeleteObject();
                context.ExecuteQuery();
            }
            catch { }


            try
            {
                var keywords = web.Lists.GetByTitle(KEYWORDSLISTTITLE);
                context.Load(keywords);
                keywords.DeleteObject();
                context.ExecuteQuery();
            }
            catch { }

        }
    }
}