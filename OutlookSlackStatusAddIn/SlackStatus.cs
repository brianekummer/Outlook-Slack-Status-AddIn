using System.Xml;

namespace OutlookSlackStatusAddIn
{
    class SlackStatus
    {
        public SlackStatus()
        {
        }

        public SlackStatus(XmlNode slackSettingXml)
        {
            Text = slackSettingXml.SelectSingleNode("text")?.InnerText;
            Emoji = slackSettingXml.SelectSingleNode("emoji")?.InnerText;
        }

        public string Text;
        public string Emoji;
    }
}
