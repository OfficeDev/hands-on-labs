



using Microsoft.AspNet.WebHooks;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using GithubTest;
using System.Net.Http;
using System;
using System.Text;

namespace WebApplication1.WebHookHandlers
{
    public class GitHubWebHookHandler : WebHookHandler
    {
        // $TODO: Copy and paste the group webhook URL here
        public const string groupWebHookURL = @"paste the Office 365 Group webhook URL here";

        public override Task ExecuteAsync(string receiver, WebHookHandlerContext context)
        {
			// make sure we're only processing the intended type of hook
			if("GitHub".Equals(receiver, System.StringComparison.CurrentCultureIgnoreCase))
			{
				// todo: replace this placeholder functionality with your own code
				string action = context.Actions.First();
				JObject incoming = context.GetDataOrDefault<JObject>();
                string connectorCardPayload = ConnectorCard.ConvertGithubJsonToConnectorCard(incoming.ToString());
                var body = PostRequest(connectorCardPayload).Result;
            }

            return Task.FromResult(true);
        }

        private static async Task<HttpResponseMessage> PostRequest(string payload)
        {
            var targetUri = new Uri(groupWebHookURL);
            var httpClient = new HttpClient();

            return await httpClient.PostAsync(targetUri,
                 new StringContent(payload, Encoding.UTF8, "application/json"));
        }

    }
}

