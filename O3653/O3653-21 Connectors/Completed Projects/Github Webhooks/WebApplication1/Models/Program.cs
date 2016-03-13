using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Griffin.Connectors.Common.Formatting.NewSwift.ObjectModel;
using Newtonsoft.Json;
using System.Net.Http;

namespace GithubTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string jsonFile = "F:\\Temp\\Swift\\Github.json";
            string jsonContent = File.ReadAllText(jsonFile);

            GithubIssueEvent issueEvent = JsonConvert.DeserializeObject<GithubIssueEvent>(jsonContent, new JsonSerializerSettings { DefaultValueHandling = DefaultValueHandling.Populate });

            ModelBuilder builder = new ModelBuilder(issueEvent);

            SwiftModel model = new SwiftModel();

            model.Sender = "Github";
            model.SenderImage = "https://assets-cdn.github.com/images/modules/logos_page/GitHub-Mark.png";
            model.Summary = builder.BuildSubject();
            model.ThemeColor = "FFFFFF";
            model.Sections = builder.BuildSections();
            model.PotentialActions = builder.BuildActions();

            string payload = JsonConvert.SerializeObject(model);
            Console.WriteLine("Posting...");
            var body = PostRequest(payload).Result;
           
        }

        private static async Task<HttpResponseMessage> PostRequest(string payload)
        {
            var targetUri = new Uri("https://outlook.office365.com/webhook/1bf70727-5c6c-4f66-9d05-8100878908c7@72f988bf-86f1-41af-91ab-2d7cd011db47/IncomingWebhook/fee9d3c5f0124921af82031849c6d4b5/50746df0-ecd2-4840-8b1a-8a4c27a73595");
            var httpClient = new HttpClient();

            return await httpClient.PostAsync(targetUri,
                 new StringContent(payload, Encoding.UTF8, "application/json"));
        }
    }
}
