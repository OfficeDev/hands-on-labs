using Newtonsoft.Json;

namespace InboxSync.Helpers
{
  public class OutlookNotification
  {
    public string SubscriptionId { get; set; }
  }

  public class NotificationPayload
  {
    [JsonProperty("value")]
    public OutlookNotification[] Notifications { get; set; }
  }
}