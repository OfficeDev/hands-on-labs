using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using Microsoft.AspNet.SignalR;
using Microsoft.Graph;


namespace GraphWebhooks.SignalR
{
    public class NotificationService : PersistentConnection
    {
      public void SendNotificationToClient(List<Message> messages)
        {
            var hubContext = GlobalHost.ConnectionManager.GetHubContext<NotificationHub>();
            if (hubContext != null)
            {
                hubContext.Clients.All.showNotification(messages);
            }
        }
    }
}