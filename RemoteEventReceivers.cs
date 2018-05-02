using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using System.Reflection;

namespace WebcorLessonsLearnedAppWeb.Utilities
{
    public class RemoteEventReceivers

    {  
        static string RECEIVERCLASSPREFIX = "WebcorLessonsLearnedAppWeb.Services.";

        public static string DisplayEventReceivers(List list, string receiverName)
        {
            string eventReceivers = "";

            ClientContext context = (ClientContext)list.Context;
            context.Load(list);
            context.ExecuteQuery();

            EventReceiverDefinitionCollection erdc = list.EventReceivers;
            context.Load(erdc);
            context.ExecuteQuery();

            List<EventReceiverDefinition> toDelete = new List<EventReceiverDefinition>();
            foreach (EventReceiverDefinition erd in erdc)
            {
                if (erd.ReceiverName == receiverName)
                {
                    toDelete.Add(erd);

                    eventReceivers += "Name: " + erd.ReceiverName + " Class: " + erd.ReceiverClass + " Event Type:" + erd.EventType + " url:" + erd.ReceiverUrl + " Assembly:" + erd.ReceiverAssembly + "\n";
                }
            }

            return eventReceivers;
        }

        public static void DeleteEventReceivers(List list, string receiverName)
        {
            ClientContext context = (ClientContext)list.Context;
            context.Load(list);
            context.ExecuteQuery();

            EventReceiverDefinitionCollection erdc = list.EventReceivers;
            context.Load(erdc);
            context.ExecuteQuery();

            List<EventReceiverDefinition> toDelete = new List<EventReceiverDefinition>();
            foreach (EventReceiverDefinition erd in erdc)
            {
                if (erd.ReceiverName == receiverName)
                {
                    toDelete.Add(erd);
                }
            }

            //Delete the remote event receiver from the list
            foreach (EventReceiverDefinition item in toDelete)
            {
                item.DeleteObject();
                context.ExecuteQuery();
            }
        }

        public static void AddEventReceivers(List list, string receiverName, string receiverUrl)
        {

            ClientContext context = (ClientContext)list.Context;

            string receiverAssembly = Assembly.GetExecutingAssembly().FullName;
            EventReceiverDefinitionCollection eventReceivers = list.EventReceivers;
            string receiverClass = RECEIVERCLASSPREFIX + receiverName;

            EventReceiverDefinitionCreationInformation updatedEventReceiver = new EventReceiverDefinitionCreationInformation()
            {
                ReceiverAssembly = receiverAssembly,
                ReceiverName = receiverName,
                ReceiverClass = receiverClass,
                ReceiverUrl = receiverUrl,
                SequenceNumber = 1000,
                EventType = EventReceiverType.ItemUpdated
            };

            context.Load(eventReceivers);
            context.ExecuteQuery();
            list.EventReceivers.Add(updatedEventReceiver);
            context.ExecuteQuery();

            EventReceiverDefinitionCreationInformation addedEventReceiver = new EventReceiverDefinitionCreationInformation()
            {
                ReceiverAssembly = receiverAssembly,
                ReceiverName = receiverName,
                ReceiverClass = receiverClass,
                ReceiverUrl = receiverUrl,
                SequenceNumber = 1000,
                EventType = EventReceiverType.ItemAdded
            };

            context.Load(eventReceivers);
            context.ExecuteQuery();
            list.EventReceivers.Add(addedEventReceiver);
            context.ExecuteQuery();
        }


    }
}