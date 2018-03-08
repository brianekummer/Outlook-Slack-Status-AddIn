using System.Xml;

namespace OutlookSlackStatusAddIn
{
    class SlackStatus
    {
        public SlackStatus()
        {
        }

        public SlackStatus(string slackStatusAsDelimitedText)
        {
            var parts = slackStatusAsDelimitedText.Split('|');
            Text = parts[0];
            Emoji = parts[1];
        }

        public string Text;
        public string Emoji;
    }
}
