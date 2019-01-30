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
            if (parts.Length > 2)
                Expiration = int.Parse(parts[2]);
        }

        public string Text;
        public string Emoji;
        public long Expiration;
    }
}
