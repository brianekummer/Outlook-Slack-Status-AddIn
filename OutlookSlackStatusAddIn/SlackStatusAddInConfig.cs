using System.Xml;

namespace OutlookSlackStatusAddIn
{
    class SlackStatusAddInConfig
    {
        public SlackStatusAddInConfig(XmlDocument settings)
        {
            MySlackToken = settings.SelectSingleNode("/configuration/mySlackToken")?.InnerText;
            MyLastName = settings.SelectSingleNode("/configuration/myLastName")?.InnerText;
            NetworkOffice = settings.SelectSingleNode("/configuration/networkOffice")?.InnerText;
            InMeeting = new SlackStatus(settings.SelectSingleNode("/configuration/meeting"));
            WorkingInOffice = new SlackStatus(settings.SelectSingleNode("/configuration/workingInOffice"));
            WorkingRemotely = new SlackStatus(settings.SelectSingleNode("/configuration/workingRemotely"));
            OnVacation = new SlackStatus(settings.SelectSingleNode("/configuration/vacation"));
        }

        public string MySlackToken;
        public string MyLastName;
        public string NetworkOffice;
        public SlackStatus InMeeting;
        public SlackStatus WorkingInOffice;
        public SlackStatus WorkingRemotely;
        public SlackStatus OnVacation;
    }
}
