using Microsoft.AspNet.SignalR;
using InboxSync.Helpers;
using System.Threading.Tasks;

namespace InboxSync.Hubs
{
  public class NotificationHub : Hub
  {
    public async Task RegisterUserConnection(string userId)
    {
      await UserManager.UpdateUserSignalRConnection(userId, Context.ConnectionId);
    }
  }
}