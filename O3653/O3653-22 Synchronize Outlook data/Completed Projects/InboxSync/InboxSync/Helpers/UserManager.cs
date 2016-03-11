using System;
using System.Collections.Generic;
using System.Threading.Tasks;

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
  }
}