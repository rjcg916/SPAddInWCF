using Microsoft.SharePoint.Client;

namespace WebcorLessonsLearnedAppWeb.Utilities
{
    public class CustomActions
    {
        public static void AddCustomAction(List list, string actionUrl)
        {
            ClientContext context = (ClientContext)list.Context;

            context.Load(list.UserCustomActions);
            context.ExecuteQuery();

            foreach (UserCustomAction uca in list.UserCustomActions)
            {
                uca.DeleteObject();
            }

            context.ExecuteQuery();

            UserCustomAction action = list.UserCustomActions.Add();
            action.Location = "EditControlBlock";
            action.Sequence = 10001;
            action.Title = "Publish this Lessons Learned";
            var permissions = new BasePermissions();
            permissions.Set(PermissionKind.EditListItems);
            action.Rights = permissions;
            action.Url = actionUrl;
    
            action.Update();

            context.ExecuteQuery();

        }
    }
}