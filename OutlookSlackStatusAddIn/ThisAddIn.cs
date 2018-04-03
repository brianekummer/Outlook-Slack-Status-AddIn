using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Net;
using System.Net.NetworkInformation;

namespace OutlookSlackStatusAddIn
{
    public partial class ThisAddIn
    {
        // To deploy rebuild a setup.exe for this addin, right-click the PROJECT "OutlookSlackStatusAddin"
        // and select "Publish"
        

        private const string TASK_PREFIX = @"SLACK-STATUS-UPDATE";
        private const string CRLF = @"(\n|\r|\r\n)";

        private SlackStatusAddInConfig _config;
        private WebRequest _webRequest;


        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.Reminder += ThisAddIn_Reminder;
            Application.Reminders.BeforeReminderShow += ThisAddin_BeforeReminderShow;

            WriteToLog("Starting");

            _config = new SlackStatusAddInConfig();

            if (_config.MySlackToken == null || _config.MyLastName == null)
            {
                WriteToLog("  CONFIGURATION ERROR- No environment variables");

                string errMsg = "A number of Windows environment variables need set to make this\n" +
                                "add-in work.\n" +
                                "\n" +
                                "INSTRUCTIONS:\n" +
                                "- Update the contents of \"Slack Status Update Config.bat\"\n" +
                                "- Run \"Slack Status Update Config.bat\"\n" +
                                "- Logout of Windows or reboot\n" +
                                "- From a command prompt, run \"SET\" to verify your changes took\n" +
                                "- Delete \"Slack Status Update Config.bat\" - you don't need it anymore!";
                MessageBox.Show(errMsg, @"Slack Status Update Add-In for Outlook", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void ThisAddIn_Reminder(Object item)
        {
            if (item is Outlook.AppointmentItem myAppointmentItem)
            {
                // This is the reminder for the start of an appointment
                WriteToLog("ThisAddIn_Reminder for APPOINTMENT: " + myAppointmentItem.Subject);
                WriteToLog("  Body: " + TruncateAndCleanUpText(myAppointmentItem.Body));
                WriteToLog("  Start: " + myAppointmentItem.Start + ", End: " + myAppointmentItem.End + ", BusyStatus: " + myAppointmentItem.BusyStatus);

                if (DateTime.Now >= myAppointmentItem.Start && DateTime.Now < myAppointmentItem.End && myAppointmentItem.BusyStatus != Outlook.OlBusyStatus.olFree)
                {
                    // This reminder has fired sometime between the start and end of the appointment, 
                    // and the appointment has me bus, or out of the office, etc.

                    if (myAppointmentItem.Subject.Contains("PTO") && myAppointmentItem.Organizer.Contains(_config.MyLastName))
                    {
                        WriteToLog("    Is PTO");
                        // This appointment is for my PTO

                        var slackStatusText = GetSlackStatus().Item1;

                        if (slackStatusText.Contains("On PTO"))
                        {
                            // My Slack status says I'm on already PTO
                            if (DateTime.Now >= myAppointmentItem.End)
                            {
                                // My PTO is over
                                WriteToLog("    PTO is OVER");
                                SetSlackStatusBasedOnNetwork();
                            }
                            else
                            {
                                // My PTO is not yet over, so create a task with a reminder
                                // for the time my PTO ends. When the reminder fires, this 
                                // will fire for that task, and we will set our Slack status 
                                // appropriately.
                                WriteToLog("    PTO is -NOT- over");
                                CreateTaskWithReminder(myAppointmentItem);
                            }
                        }
                        else
                        {
                            WriteToLog("    Is -NOT- PTO");

                            // I'm not on PTO
                            if (DateTime.Now >= myAppointmentItem.Start)
                            {
                                // It's time to start my PTO!
                                slackStatusText = "On PTO ";
                                if (myAppointmentItem.End.Date == DateTime.Today ||
                                    myAppointmentItem.End == DateTime.Today.AddDays(1))
                                {
                                    // PTO ends sometime today or at midnight tomorrow
                                    slackStatusText += "today";
                                }
                                else
                                {
                                    // If PTO does not end at midnight, then myAppointmentItem.End 
                                    // is the day we're returning to work. If PTO ends at midnight, 
                                    // then myAppointmentItem.End is the day AFTER our PTO ends, 
                                    // and we should calculate the next working day.
                                    var nextWorkingDay = (myAppointmentItem.End.ToString("HHmmss") != "000000")
                                        ? myAppointmentItem.End
                                        : AddBusinessDays(myAppointmentItem.End.AddMinutes(-1), -1);
                                    var dateFormat = (nextWorkingDay.Date - DateTime.Today).TotalDays  < 7
                                        ? "dddd"
                                        : "dddd, MMM d";
                                    slackStatusText += "until " + nextWorkingDay.ToString(dateFormat);
                                }

                                SetSlackStatus(new SlackStatus { Text = slackStatusText, Emoji = _config.OnVacation.Emoji });
                            }
                        }
                    }
                    else
                    {
                        WriteToLog("    Is not my PTO");

                        // For this appointment/meeting, we want to change Slack status if
                        //   - Meeting is starting now or has already started
                        //   - I am not free (ASSUMES that if I add a meeting to my calendar and status is Free, 
                        //     then I want to be available by Slack)
                        SetSlackStatus(_config.InMeeting);

                        // Create a task with a reminder for the time the meeting ends.
                        // When the reminder fires, this will fire for that task, and
                        // we will set our Slack status appropriately.
                        CreateTaskWithReminder(myAppointmentItem);
                    }
                }
            }

            else if (item is Outlook.TaskItem myTaskItem)
            {
                WriteToLog("ThisAddIn_Reminder for TASK: " + TruncateAndCleanUpText(myTaskItem.Subject));
                WriteToLog("  ReminderTime: " + myTaskItem.ReminderTime);

                // This is the reminder for the task that marks the end of the approintment
                if (myTaskItem.Subject.Contains(TASK_PREFIX))
                {
                    WriteToLog("    Deleting task");
                    myTaskItem.Delete();

                    WriteToLog("    Setting slack status");
                    SetSlackStatusBasedOnNetwork();
                }
            }
        }


        private string TruncateAndCleanUpText(string subject)
        {

            var regex = new Regex("[\t\r\n]");

            var cleanedUpText = regex.Replace(subject, " ");

            return (cleanedUpText.Length > 50) ? $"{cleanedUpText.Substring(0, 50)}..." : cleanedUpText;
        }


        private void CreateTaskWithReminder(Outlook.AppointmentItem myAppointmentItem)
        {
            // Create a task with a reminder for the time the meeting ends.
            // When the reminder fires, this will fire for that task, and
            // we will set our Slack status appropriately.
            var olTask = Application.CreateItem(Outlook.OlItemType.olTaskItem);
            olTask.Subject = $"{TASK_PREFIX}-{myAppointmentItem.Subject}:{myAppointmentItem.Start:yyyyMMddHHmmss}-{myAppointmentItem.End:yyyyMMddHHmmss}";
            olTask.Status = Outlook.OlTaskStatus.olTaskInProgress;
            olTask.Importance = Outlook.OlImportance.olImportanceLow;
            olTask.ReminderSet = true;
            olTask.ReminderTime = myAppointmentItem.End;
            olTask.Save();
        }


        private void ThisAddin_BeforeReminderShow(ref bool cancel)
        {
            // Automatically close the reminder for the task that we 
            // created that fires at the end of the appointment
            //
            // NOTE that this sometimes can take several seconds to
            // close the reminder

            foreach (Outlook.Reminder objRem in Application.Reminders)
            {
                if (objRem.Caption.Contains(TASK_PREFIX))
                {
                    if (objRem.IsVisible)
                    {
                        objRem.Dismiss();
                        cancel = true;
                    }
                    break;
                }
            }
        }


        private bool connectedToInternet()
        {
            var connected = false;

            try
            {
                connected = new Ping().Send("www.google.com.mx").Status == IPStatus.Success;
            }
            catch 
            {
                connected = false;
            }

            return connected;
        }

        private void SetSlackStatusBasedOnNetwork()
        {
            if (AmNearOfficeWifiNetwork())
            {
                SetSlackStatus(_config.WorkingInOffice);
            }
            else if (connectedToInternet())
            {
                SetSlackStatus(_config.WorkingRemotely);
            }
            else
            {
                // Not connected any network. Ideally, we'd put this on a timer and 
                // try again shortly.
            }
        }


        private bool AmNearOfficeWifiNetwork()
        {
            // Look at each wifi network that is available to the user. If any of them
            // match the regular expression _config.OfficeNetworkNames then we are at
            // work.
            var atWork = false;

            var allNetworks = RunShell("cmd.exe", "/c netsh wlan show networks");
            MatchCollection matches = Regex.Matches(allNetworks, @"\r\nssid.+?:\s(.*)\r\n", RegexOptions.IgnoreCase|RegexOptions.Multiline);
            foreach (Match match in matches)
            {
                if (match.Groups.Count > 1 && Regex.IsMatch(match.Groups[1].Value, _config.OfficeNetworkNames, RegexOptions.IgnoreCase))
                {
                    atWork = true;
                    break;
                }
            }

            return atWork;
        }


        private string RunShell(string cmd, string cmdParams)
        {
            Process proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = cmd,
                    Arguments = cmdParams,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
            };

            proc.Start();
            var cmdResults = proc.StandardOutput.ReadToEnd();
            proc.WaitForExit();

            return cmdResults;
        }        


        private void SetSlackStatus(SlackStatus slackStatus)
        {
            WriteToLog("      >> Setting Slack status to " + slackStatus.Emoji + " " + slackStatus.Text);

            byte[] byteArray = Encoding.UTF8.GetBytes(
                $"profile={{'status_text': '{slackStatus.Text}', 'status_emoji': '{slackStatus.Emoji}'}}");

            _webRequest = WebRequest.Create("https://slack.com/api/users.profile.set");
            _webRequest.ContentType = "application/x-www-form-urlencoded";
            _webRequest.ContentLength = byteArray.Length;
            _webRequest.Method = "POST";
            _webRequest.Headers.Add("Authorization", $"Bearer {_config.MySlackToken}");

            using (Stream s = _webRequest.GetRequestStream())
            {
                s.Write(byteArray, 0, byteArray.Length);
            }
        }


        private Tuple<string, string> GetSlackStatus()
        {
            string responseFromServer;

            _webRequest = WebRequest.Create("https://slack.com/api/users.profile.get");
            _webRequest.ContentType = "application/x-www-form-urlencoded";
            _webRequest.Method = "GET";
            _webRequest.Headers.Add("Authorization", $"Bearer {_config.MySlackToken}");

            using (WebResponse response = _webRequest.GetResponse())
            {
                using (Stream stream = response.GetResponseStream())
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        responseFromServer = reader.ReadToEnd();
                    }
                }
            }

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            dynamic slackProfile = serializer.DeserializeObject(responseFromServer);

            return Tuple.Create(slackProfile["profile"]["status_text"], 
                slackProfile["profile"]["status_emoji"]);
        }


        private DateTime AddBusinessDays(DateTime current, int days)
        {
            var sign = Math.Sign(days);
            var unsignedDays = Math.Abs(days);
            for (var i = 0; i < unsignedDays; i++)
            {
                do
                {
                    current = current.AddDays(sign);
                }
                while (current.DayOfWeek == DayOfWeek.Saturday ||
                       current.DayOfWeek == DayOfWeek.Sunday);
            }
            return current;
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }


        private void WriteToLog(string whatToWrite)
        {
            using (StreamWriter outputFile = new StreamWriter(@"C:\Temp\SlackStatusUpdateAddIn.log", true))
            {
                outputFile.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss}: {whatToWrite}");
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
 }