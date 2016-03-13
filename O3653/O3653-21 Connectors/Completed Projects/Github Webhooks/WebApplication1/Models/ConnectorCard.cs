using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using GithubTest;
using Newtonsoft.Json;
using Microsoft.Connectors.Common.Card.Formatting.ObjectModel;

namespace GithubTest
{

    public static  class ConnectorCard
    {

        public static string ConvertGithubJsonToConnectorCard(string gitJsonContent)
        {
            GithubIssueEvent issueEvent = JsonConvert.DeserializeObject<GithubIssueEvent>(gitJsonContent, new JsonSerializerSettings { DefaultValueHandling = DefaultValueHandling.Populate });

            ModelBuilder builder = new ModelBuilder(issueEvent);

            SwiftModel model = new SwiftModel();

            model.Sender = "Github";
            model.SenderImage = "https://assets-cdn.github.com/images/modules/logos_page/GitHub-Mark.png";
            model.Summary = builder.BuildSubject();
            model.ThemeColor = "FFFFFF";
            model.Sections = builder.BuildSections();
            model.PotentialActions = builder.BuildActions();

            return JsonConvert.SerializeObject(model);
        }        

    }
}