using AdaptiveCards;
using Microsoft.Bot.Connector;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using TeamsTalentMgmtApp.DataModel;

namespace TeamsTalentMgmtApp.Utils
{
    public class TaskModuleHelper
    {
        public JObject CreateJobPostingTaskModuleResponse()
        {
            var card = CardHelper.CreateExtendedCardForNewJobPosting();
            Attachment attachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card
            };
            return CreateTaskResponseFromCard(attachment, "Create new job posting");
        }

        private JObject CreateTaskResponseFromCard(Attachment card, string title)
        {
            // TODO: Convert this to helpers once available.
            JObject taskEnvelope = new JObject();

            JObject taskObj = new JObject();
            JObject taskInfo = new JObject();

            taskObj["type"] = "continue";
            taskObj["value"] = taskInfo;

            taskInfo["card"] = JObject.FromObject(card);
            taskInfo["title"] = title;

            taskEnvelope["task"] = taskObj;
            return taskEnvelope;
        }
    }
}