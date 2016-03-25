using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using FindMeetingTimesLab.Models;
using Newtonsoft.Json;
using System.Text;

namespace FindMeetingTimesLab
{
    public class GraphHelper
    {
        // Used to set the base API endpoint, e.g. "https://outlook.office.com/api/beta"
        public string apiEndpoint { get; set; }
        public string anchorMailbox { get; set; }

        public GraphHelper()
        {
            apiEndpoint = "https://outlook.office.com/api/beta";
            anchorMailbox = string.Empty;
        }

        public async Task<HttpResponseMessage> MakeGraphApiCall(string method, string token, string apiUrl, string userEmail, string payload, Dictionary<string, string> preferHeaders)
        {
            using (var httpClient = new HttpClient())
            {
                var request = new HttpRequestMessage(new HttpMethod(method), apiUrl);

                // Headers
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                request.Headers.UserAgent.Add(new System.Net.Http.Headers.ProductInfoHeaderValue("FindMeetingTimesLab", "beta"));
                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Add("client-request-id", Guid.NewGuid().ToString());
                request.Headers.Add("return-client-request-id", "true");
                request.Headers.Add("X-AnchorMailbox", userEmail);

                if (preferHeaders != null)
                {
                    foreach (KeyValuePair<string, string> header in preferHeaders)
                    {
                        request.Headers.Add("Prefer", string.Format("{0}=\"{1}\"", header.Key, header.Value));
                    }
                }

                // Content
                if ((method.ToUpper() == "POST" || method.ToUpper() == "PATCH") &&
                    !string.IsNullOrEmpty(payload))
                {
                    request.Content = new StringContent(payload);
                    request.Content.Headers.ContentType.MediaType = "application/json";
                }

                var apiResult = await httpClient.SendAsync(request);
                return apiResult;
            }
        }

        public async Task<object> GetMeetingTimes(string token, string userEmail, string payload)
        {
            string findMeetingTimesEndpoint = this.apiEndpoint + "/Me/FindMeetingTimes";

            //var jsonPayload = await Task.Run(() => JsonConvert.SerializeObject(payload));

            var result = await MakeGraphApiCall("POST", token, findMeetingTimesEndpoint, userEmail, payload, null);

            var response = await result.Content.ReadAsStringAsync();

            JObject responseJson = JObject.Parse(response);
            JArray eventJson = (JArray)responseJson["value"];

            List<MeetingTimeCandidate> meetingTimes = new List<MeetingTimeCandidate>();
            foreach (var e in eventJson)
            {
                MeetingTimeCandidate nextItem = new MeetingTimeCandidate();

                //add all the values
                nextItem.StartDate = DateTime.Parse((string)e["MeetingTimeSlot"]["Start"]["Date"]);
                nextItem.StartTime = DateTime.Parse((string)e["MeetingTimeSlot"]["Start"]["Time"]);
                nextItem.EndDate = DateTime.Parse((string)e["MeetingTimeSlot"]["End"]["Date"]);
                nextItem.EndTime = DateTime.Parse((string)e["MeetingTimeSlot"]["End"]["Time"]);
                nextItem.Confidence = int.Parse((string)e["Confidence"]);
                nextItem.Score = int.Parse((string)e["Score"]);
                if (e["MeetingTimeSlot"]["Location"] != null)
                {
                    nextItem.LocationDisplayName = (string)e["MeetingTimeSlot"]["Location"]["Time"];
                    nextItem.LocationAddress = BuildAddressString(e["MeetingTimeSlot"]["Location"]["Address"]);
                    nextItem.LocationCoordinates = BuildCoordinatesString(e["MeetingTimeSlot"]["Location"]["Coordinates"]);
                }

                meetingTimes.Add(nextItem);
            }

            return meetingTimes;
        }

        public String BuildAddressString(JToken address)
        {
            if (address == null)
            {
                return "null";
            }
            else {
                return String.Format("{0}, {1}, {2}, {3}, {4}",
                    (string)address["Street"] == "" ? "<No Street>" : (string)address["Street"],
                    (string)address["City"] == "" ? "<No City>" : (string)address["City"],
                    (string)address["State"] == "" ? "<No State>" : (string)address["State"],
                    (string)address["CountryOrRegion"] == "" ? "<No Country Or Region>" : (string)address["CountryOrRegion"],
                    (string)address["PostalCode"] == "" ? "<No Postal Code>" : (string)address["PostalCode"]
                    );
            }
        }

        public String BuildCoordinatesString(JToken coordinates)
        {
            if (coordinates == null)
            {
                return "null";
            }
            else
            {
                return String.Format("{0}, {1}", (string)coordinates["Latitude"], (string)coordinates["Longitude"]);
            }
        }

        public string GeneratePayload(string attendees)
        {
            const string payloadStart = "{";
            const string AttendeesListStart = "\"Attendees\": [";
            const string AttendeeStart = "{\"Type\":\"Required\",\"EmailAddress\":{\"Address\":\"";
            const string AttendeeEnd = "\"}}";
            const string AttendeesListEnd = "]";
            const string payloadEnd = "}";

            StringBuilder payloadBuilder = new StringBuilder(payloadStart);

            //Add all attendees
            string[] attendeeEmails = attendees.Split(',');

            payloadBuilder.Append(AttendeesListStart);
            foreach (var e in attendeeEmails)
            {
                payloadBuilder.Append(AttendeeStart);
                payloadBuilder.Append(e);
                payloadBuilder.Append(AttendeeEnd);
                payloadBuilder.Append(',');
            }
            payloadBuilder.Remove(payloadBuilder.Length - 1, 1);
            payloadBuilder.Append(AttendeesListEnd);

            payloadBuilder.Append(payloadEnd);

            return payloadBuilder.ToString();
        }

    }
}