using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FindMeetingTimesLab.Models
{
    public class MeetingTimeCandidate
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public int Confidence { get; set; }
        public int Score { get; set; }
        public string LocationDisplayName { get; set; }
        public string LocationAddress { get; set; }
        public string LocationCoordinates { get; set; }
    }
}