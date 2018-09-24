using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Teams;
using System.Configuration;
using TeamsTalentMgmtApp.Utils;
using Newtonsoft.Json.Linq;

namespace TeamsTalentMgmtApp
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {

        /// <summary>
        /// POST: api/messages
        /// Receive a message from a user and reply to it
        /// </summary>
        public async Task<HttpResponseMessage> Post([FromBody]Activity activity)
        {
            if (activity.Type == ActivityTypes.Message) 
            {
                //Handle basic message types, e.g. user initiated
                await Conversation.SendAsync(activity, () => new Dialogs.RootDialog());
            }
            else if (activity.Type == ActivityTypes.Invoke) 
            {
                //Compose extensions come in as Invokes. Leverage the Teams SDK helper functions
                if (activity.IsComposeExtensionQuery())
                {
                    // Determine the response object to reply with
                    MessagingExtension msgExt = new MessagingExtension(activity);
                    var invokeResponse = msgExt.CreateResponse();

                    // Return the response
                    return Request.CreateResponse(HttpStatusCode.OK, invokeResponse);
                }
                else if (activity.Name == "fileConsent/invoke")
                {
                    // Try to replace with File uploaded card.
                    return Request.CreateResponse(HttpStatusCode.OK);
                }
                else if (activity.IsTeamsVerificationInvoke())
                {
                    await Conversation.SendAsync(activity, () => new Dialogs.RootDialog());
                }
                else if (activity.Name == "task/fetch")
                {
                    JObject parameters = activity.Value as JObject;

                    if (parameters != null)
                    {
                        string command = parameters["command"].ToString();

                        // Fetch dynamic adaptive card for task module.
                        if (command == "createPostingExtended")
                        {
                            JObject resp = new TaskModuleHelper().CreateJobPostingTaskModuleResponse();
                            return Request.CreateResponse(HttpStatusCode.OK, resp);
                        }
                    }
                }
            }
            else
            {
                await HandleSystemMessage(activity);
            }
            var response = Request.CreateResponse(HttpStatusCode.OK);
            return response;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        private async Task HandleSystemMessage(Activity message)
        {

            if (message.Type == ActivityTypes.ConversationUpdate)
            {
                // Handle conversation state changes, like members being added and removed
                // Use Activity.MembersAdded and Activity.MembersRemoved and Activity.Action for info

                // This event is called when the Bot is added too, so we can trigger a welcome message if the member added is the bot:
                TeamEventBase eventData = message.GetConversationUpdateData();
                if (eventData.EventType == TeamEventType.MembersAdded)
                {
                    for (int i = 0; i < message.MembersAdded.Count; i++)
                    {
                        //Check to see if the member added was the bot itself.  We're leveraging the fact that the inbound payload's Recipient is the bot.
                        if (message.MembersAdded[i].Id == message.Recipient.Id)
                        {
                            // We'll use normal message parsing to display the welcome message.
                            message.Text = "welcome";
                            await Conversation.SendAsync(message, () => new Dialogs.RootDialog());

                            break;
                        }
                    }
                }
            }
            else if (message.Type == ActivityTypes.DeleteUserData)
            {
                // Implement user deletion here
                // If we handle user deletion, return a real message
            }
            else if (message.Type == ActivityTypes.ContactRelationUpdate)
            {
                // Handle add/remove from contact lists
            }
            else if (message.Type == ActivityTypes.Typing)
            {
                // Handle knowing that the user is typing
            }
            else if (message.Type == ActivityTypes.Ping)
            {

            }
        }
    }
}