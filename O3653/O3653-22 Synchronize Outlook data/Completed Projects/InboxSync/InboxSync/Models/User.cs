using System;
using MongoDB.Bson;

namespace InboxSync.Models
{
  public class User
  {
    public ObjectId Id { get; set; }
    public string Email { get; set; }
    public DateTime TokenExpires { get; set; }
    public string AccessToken { get; set; }
    public string RefreshToken { get; set; }
    public string IdToken { get; set; }
    public string SyncState { get; set; }
    public string SubscriptionId { get; set; }
    public DateTime SubscriptionExpires { get; set; }
    public string SignalRConnection { get; set; }
  }
}