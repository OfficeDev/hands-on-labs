using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web;
using System.Collections.Specialized;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

using MongoDB.Bson;
using MongoDB.Driver;

using InboxSync.Models;
using InboxSync.Auth;

namespace InboxSync.Helpers
{
  // This class manages all of the user data stored in the database.
  public class UserManager
  {
    // Configuration information for the database connection
    private static string mongoUri = System.Configuration.ConfigurationManager.AppSettings["mongoUri"];
    private static string mongoDB = System.Configuration.ConfigurationManager.AppSettings["mongoDB"];

    private static MongoClient mongo = new MongoClient(mongoUri);
    private static IMongoDatabase database = mongo.GetDatabase(mongoDB);
    private static IMongoCollection<User> userCollection = database.GetCollection<User>("users");
    private static IMongoCollection<Message> messageCollection = database.GetCollection<Message>("messages");

    // Gets a user by ID
    public static async Task<User> GetUserById(string userId)
    {
      var idFilter = Builders<User>.Filter.Eq("Id", new ObjectId(userId));
      return await userCollection.Find(idFilter).FirstOrDefaultAsync();
    }

    // Processes a token response object and updates the user's record.
    // If the user doesn't exist in the database, it will create a record.
    public static async Task<User> AddOrUpdateUser(string email, TokenRequestSuccessResponse tokenResponse)
    {
      // Create a filter to find the record with this email
      var emailFilter = Builders<User>.Filter.Eq("Email", email);

      // See if the user exists
      User user = await userCollection.Find(emailFilter).FirstOrDefaultAsync();
      if (null != user)
      {
        // User exists, update tokens and expire time
        var update = Builders<User>.Update.Set("AccessToken", tokenResponse.AccessToken)
          .Set("RefreshToken", tokenResponse.RefreshToken)
          .Set("TokenExpires", DateTime.UtcNow.AddSeconds(Int32.Parse(tokenResponse.ExpiresIn) - 300));

        var result = await userCollection.UpdateOneAsync(emailFilter, update);
      }
      else
      {
        // User does not exist, create a new user and add to the database
        user = new User()
        {
          Email = email,
          AccessToken = tokenResponse.AccessToken,
          RefreshToken = tokenResponse.RefreshToken,
          TokenExpires = DateTime.UtcNow.AddSeconds(Int32.Parse(tokenResponse.ExpiresIn) - 300)
        };

        await userCollection.InsertOneAsync(user);
      }

      return user;
    }

    // Gets the user's total count of messages in the database
    // This is used to enable the PagedList to calculate the appropriate number
    // of pages without retrieving all of the messages each time it loads a new page.
    public static async Task<long> GetUsersMessageCount(string userId)
    {
      var ownerFilter = Builders<Message>.Filter.Eq("Owner", new ObjectId(userId));
      return await messageCollection.Find(ownerFilter).CountAsync();
    }

    // Gets the users's current "page" of messages from the database
    public static async Task<List<Message>> GetUsersMessages(string userId, int pageSize, int pageNum)
    {
      var ownerFilter = Builders<Message>.Filter.Eq("Owner", new ObjectId(userId));
      var sort = Builders<Message>.Sort.Descending("ReceivedDateTime");

      return await messageCollection.Find(ownerFilter)
                                    .Sort(sort)
                                    .Limit(pageSize)
                                    .Skip((pageNum - 1) * pageSize)
                                    .ToListAsync();
    }

    public static async Task SyncUsersInbox(string userId)
    {
      User user = await GetUserById(userId);

      if (null != user)
      {
        OutlookHelper outlook = new OutlookHelper();
        outlook.AnchorMailbox = user.Email;

        // If we have no prior sync state, do an intial sync. This sync always 
        // returns a deltatoken, even if there are more results on the server. 
        // You must always do another sync after the initial sync.
        bool isInitialSync = string.IsNullOrEmpty(user.SyncState);
        bool syncComplete = false;
        string newSyncState = user.SyncState;
        string newSkipToken = string.Empty;

        while (!syncComplete)
        {
          var syncResults = await outlook.SyncInbox(user.Email, user.AccessToken, newSyncState, newSkipToken);

          // Do updates (add/delete/update)
          await ParseSyncItems(user.Id, (JArray)syncResults["value"]);

          // Is there a skip token?
          if (null != syncResults["@odata.nextLink"])
          {
            // There are more results, need to call again
            // nextLink is a URL which has a $skiptoken parameter
            string nextLink = (string)syncResults["@odata.nextLink"];
            string query = new UriBuilder(nextLink).Query;
            NameValueCollection queryParams = HttpUtility.ParseQueryString(query);

            if (string.IsNullOrEmpty(queryParams["$skiptoken"]))
            {
              throw new Exception(
                "Failed to find $skiptoken in nextLink value in sync response. Try resetting sync state.");
            }

            newSyncState = string.Empty;
            newSkipToken = queryParams["$skiptoken"];
          }

          else if (null != syncResults["@odata.deltaLink"])
          {
            // Delta link is a URL which has a $deltatoken parameter
            string deltaLink = (string)syncResults["@odata.deltaLink"];
            string query = new UriBuilder(deltaLink).Query;
            NameValueCollection queryParams = HttpUtility.ParseQueryString(query);

            if (string.IsNullOrEmpty(queryParams["$deltatoken"]))
            {
              throw new Exception(
                "Failed to find $deltatoken in deltaLink value in sync response. Try resetting sync state.");
            }

            newSkipToken = string.Empty;
            newSyncState = queryParams["$deltatoken"];
            if (isInitialSync)
            {
              isInitialSync = false;
            }
            else
            {
              syncComplete = true;
            }
          }
          else
          {
            // No deltaLink or nextLink, something's wrong
            throw new Exception("Failed to find either a deltaLink or nextLink value in sync response. Try resetting sync state.");
          }
        }

        // Save new sync state to the user's record in database
        var idFilter = Builders<User>.Filter.Eq("Id", new ObjectId(userId));
        var userUpdate = Builders<User>.Update.Set("SyncState", newSyncState);

        var result = await userCollection.UpdateOneAsync(idFilter, userUpdate);
      }
    }

    private static async Task ParseSyncItems(ObjectId userId, JArray syncItems)
    {
      List<string> deleteIds = new List<string>();
      List<Message> newMessages = new List<Message>();

      foreach (JObject syncItem in syncItems)
      {
        // First check if this is a delete
        // Deletes look like:
        // {
        //   "@odata.context": "https://...",
        //   "id": "Messages('AAMk...')"
        //   "reason": "deleted"
        // }

        if (null != syncItem["reason"] &&
            ((string)syncItem["reason"]).ToLower().Equals("deleted"))
        {
          string rawid = (string)syncItem["id"];
          string outlookId = Regex.Matches(rawid, @"'([^' ]+)'")[0].Groups[1].Value;
          // Add the ID to the list of delete IDs
          deleteIds.Add(outlookId);
        }
        else
        {
          // If there's no "reason" in the item, it's
          // either an update or a new item.
          // There is no indicator in the JSON that tells us if it's an update or
          // not, so we need to figure that out ourselves.

          string outlookId = (string)syncItem["Id"];
          var outlookIdFilter = Builders<Message>.Filter.Eq("OutlookId", outlookId);

          // See if there's an item with the outlook ID
          var existingMsg = await messageCollection.Find(outlookIdFilter).FirstOrDefaultAsync();
          if (null != existingMsg)
          {
            // UPDATE (only contains updated fields)
            // For example, marking a message read:
            // {
            //   "@odata.id": "https://...",
            //   "Id": "AAMk..."
            //   "IsRead": "true"
            // }
            UpdateDefinition<Message> update = null;
            foreach (var val in syncItem)
            {
              if (!val.Key.Equals("@odata.id") && !val.Key.Equals("Id"))
              {
                string newVal = val.Value.ToString();
                update = (null == update) ? Builders<Message>.Update.Set(val.Key, newVal) : update.Set(val.Key, newVal);
              }
            }

            var updateResult = await messageCollection.UpdateOneAsync(outlookIdFilter, update);
          }
          else
          {
            // No existing record, create a new message
            // and add it to our list to bulk insert
            Message newMessage = new Message()
            {
              BodyPreview = (string)syncItem["BodyPreview"],
              // Draft messages can have a null From field
              From = (JTokenType.Null != syncItem["From"].Type) ? new FromField()
              {
                EmailAddress = new EmailAddress()
                {
                  Name = (string)syncItem["From"]["EmailAddress"]["Name"],
                  Address = (string)syncItem["From"]["EmailAddress"]["Address"]
                }
              } : null,
              IsRead = (bool)syncItem["IsRead"],
              OutlookId = outlookId,
              Owner = userId,
              ReceivedDateTime = DateTime.Parse((string)syncItem["ReceivedDateTime"]),
              Subject = (string)syncItem["Subject"]
            };

            newMessages.Add(newMessage);
          }
        }
      }

      if (newMessages.Count > 0)
      {
        await messageCollection.InsertManyAsync(newMessages);
      }

      if (deleteIds.Count > 0)
      {
        var deleteFilter = Builders<Message>.Filter.In("OutlookId", deleteIds);
        var deleteResult = await messageCollection.DeleteManyAsync(deleteFilter);
      }
    }

    public static async Task ResetSyncState(string userId)
    {
      // Remove all messages owned by the user (so resync doesn't duplicate)
      var deleteFilter = Builders<Message>.Filter.Eq("Owner", new ObjectId(userId));
      var deleteResult = await messageCollection.DeleteManyAsync(deleteFilter);

      // Clear the sync state on the user
      var idFilter = Builders<User>.Filter.Eq("Id", new ObjectId(userId));
      var userUpdate = Builders<User>.Update.Set("SyncState", string.Empty);

      var result = await userCollection.UpdateOneAsync(idFilter, userUpdate);
    }

    // Creates a notification subscription on the user's inbox and
    // saves the subscription ID on the user in the database
    public static async Task SubscribeForInboxUpdates(string userId, string notificationUrl)
    {
      var idFilter = Builders<User>.Filter.Eq("Id", new ObjectId(userId));
      var user = await userCollection.Find(idFilter).FirstOrDefaultAsync();

      if (null != user && string.IsNullOrEmpty(user.SubscriptionId))
      {
        OutlookHelper outlook = new OutlookHelper();
        outlook.AnchorMailbox = user.Email;

        JObject subscription = await outlook.CreateInboxSubscription(user.Email, user.AccessToken, notificationUrl);

        var update = Builders<User>.Update.Set("SubscriptionId", (string)subscription["Id"])
          .Set("SubscriptionExpires", DateTime.Parse((string)subscription["SubscriptionExpirationDateTime"]));

        var result = await userCollection.UpdateOneAsync(idFilter, update);
      }
    }

    // Deletes the notification subscription for the user
    public static async Task UnsubscribeForInboxUpdates(string userId)
    {
      var idFilter = Builders<User>.Filter.Eq("Id", new ObjectId(userId));
      var user = await userCollection.Find(idFilter).FirstOrDefaultAsync();

      if (null != user && !string.IsNullOrEmpty(user.SubscriptionId))
      {
        OutlookHelper outlook = new OutlookHelper();
        outlook.AnchorMailbox = user.Email;

        await outlook.DeleteSubscription(user.Email, user.AccessToken, user.SubscriptionId);

        var update = Builders<User>.Update.Set("SubscriptionId", string.Empty)
          .Set("SubscriptionExpires", DateTime.UtcNow);

        var result = await userCollection.UpdateOneAsync(idFilter, update);
      }
    }

    public static async Task<bool> IsUserSubscribed(string userId)
    {
      var user = await GetUserById(userId);
      return !string.IsNullOrEmpty(user.SubscriptionId);
    }

    // Looks up the user by the subscription ID and does a sync
    // on that user's inbox
    public static async Task<string> UpdateInboxBySubscription(string subscription)
    {
      if (!string.IsNullOrEmpty(subscription))
      {
        var subscriptionFilter = Builders<User>.Filter.Eq("SubscriptionId", subscription);
        var user = await userCollection.Find(subscriptionFilter).FirstOrDefaultAsync();

        if (null != user)
        {
          await SyncUsersInbox(user.Id.ToString());
          return user.Id.ToString();
        }
      }
      return string.Empty;
    }

    // Saves a SignalR connection ID to the user in the database
    public static async Task UpdateUserSignalRConnection(string userId, string connectionId)
    {
      var idFilter = Builders<User>.Filter.Eq("Id", new ObjectId(userId));
      var update = Builders<User>.Update.Set("SignalRConnection", connectionId);
      var result = await userCollection.UpdateOneAsync(idFilter, update);
    }

    // Gets the user's SignalR connection ID
    public static async Task<string> GetUserSignalRConnection(string userId)
    {
      var user = await GetUserById(userId);

      if (null != user)
      {
        return user.SignalRConnection;
      }

      return null;
    }
  }
}