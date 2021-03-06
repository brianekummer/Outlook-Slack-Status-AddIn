﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OutlookSlackStatusAddIn
{
    class SlackStatusAddInConfig
    {
        public SlackStatusAddInConfig()
        {
            MySlackTokens = Environment.GetEnvironmentVariable("SLACK_TOKEN")?.Split('|').ToList() ?? new List<string>();
            MyLastName = Environment.GetEnvironmentVariable("SLACK_LAST_NAME");
            OfficeNetworkNames = Environment.GetEnvironmentVariable("SLACK_OFFICE_NETWORKS");
            InMeeting = new SlackStatus(
                Environment.GetEnvironmentVariable("SLACK_STATUS_MEETING") 
                ?? "In a meeting|:spiral_calendar_pad:");
            WorkingInOffice = new SlackStatus(
                Environment.GetEnvironmentVariable("SLACK_STATUS_WORKING_OFFICE")
                ?? "|");
            WorkingRemotely = new SlackStatus(
                Environment.GetEnvironmentVariable("SLACK_STATUS_WORKING_REMOTELY") 
                ?? "Working remotely|:house_with_garden:");
            OnVacation = new SlackStatus(
                Environment.GetEnvironmentVariable("SLACK_STATUS_VACATION")
                ?? "Vacationing|:palm_tree:");
        }

        public List<string> MySlackTokens;
        public string MyLastName;
        public string OfficeNetworkNames;
        public SlackStatus InMeeting;
        public SlackStatus WorkingInOffice;
        public SlackStatus WorkingRemotely;
        public SlackStatus OnVacation;
    }
}
