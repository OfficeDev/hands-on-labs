using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace GraphWebhooks.Models
{
    // A webhooks subscription.
    public class Subscription
    {
        // The type of change in the subscribed resource that raises a notification.
        [JsonProperty(PropertyName = "changeType")]
        public string ChangeType { get; set; }

        // The string that MS Graph should send with each notification. Maximum length is 255 characters. 
        // To verify that the notification is from MS Graph, compare the value received with the notification to the value you sent with the subscription request.
        [JsonProperty(PropertyName = "clientState")]
        public string ClientState { get; set; }

        // The URL of the endpoint that receives the subscription response and notifications. Requires https.
        [JsonProperty(PropertyName = "notificationUrl")]
        public string NotificationUrl { get; set; }

        // The resource to monitor for changes.
        [JsonProperty(PropertyName = "resource")]
        public string Resource { get; set; }

        // The date and time when the webhooks subscription expires.
        // The time is in UTC, and can be up to three days from the time of subscription creation.
        [JsonProperty(PropertyName = "subscriptionExpirationDateTime")]
        public DateTimeOffset SubscriptionExpirationDateTime { get; set; }

        // The unique identifier for the webhooks subscription.
        [JsonProperty(PropertyName = "subscriptionId")]
        public string SubscriptionId { get; set; }
    }

    // The data that displays in the Subscription view.
    public class SubscriptionViewModel
    {
        public Subscription Subscription { get; set; }
    }
}
