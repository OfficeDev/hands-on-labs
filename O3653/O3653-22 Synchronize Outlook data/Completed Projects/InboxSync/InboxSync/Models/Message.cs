using System;
using MongoDB.Bson;
using System.ComponentModel.DataAnnotations;

namespace InboxSync.Models
{
  public class EmailAddress
  {
    [DisplayFormat(ConvertEmptyStringToNull = true, NullDisplayText = "*no name*")]
    public string Name { get; set; }
    [DisplayFormat(ConvertEmptyStringToNull = true, NullDisplayText = "*no email*")]
    public string Address { get; set; }
  }
  public class FromField
  {
    public EmailAddress EmailAddress { get; set; }
  }

  public class Message
  {
    public ObjectId Id { get; set; }
    public ObjectId Owner { get; set; }
    public string OutlookId { get; set; }
    [DisplayFormat(ConvertEmptyStringToNull = true, NullDisplayText = "*no body*")]
    public string BodyPreview { get; set; }
    public FromField From { get; set; }
    public bool IsRead { get; set; }
    public DateTime ReceivedDateTime { get; set; }
    [DisplayFormat(ConvertEmptyStringToNull = true, NullDisplayText = "*no subject*")]
    public string Subject { get; set; }
  }
}