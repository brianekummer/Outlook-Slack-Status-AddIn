@ECHO OFF
REM --------------------------------------------------------------------------
REM - Slack Status Update is configured using Windows environment variables. 
REM - This batch file will set the necessary environment variables for the 
REM - current Windows USER.
REM -
REM - Instructions
REM -   1. Modify this batch file as necessary 
REM -   2. Run this batch file
REM -   3. Logout and log back into Windows, or reboot
REM -   4. Make sure your changes are there (e.g. run the command "SET")
REM -   5. Undo your changes to this file or or delete it- you don't need
REM -      them anymore
REM --------------------------------------------------------------------------

REM You can get a Slack legacy token here: https://api.slack.com/custom-integrations/legacy-tokens
SETX SLACK_TOKEN "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
SETX SLACK_LAST_NAME "xxxxx"

REM A regular expression for the title of the Slack application window, which
REM is often "Slack - mycompanyname"
SETX SLACK_WINDOW_TITLE "Slack - xxxxxxxxxx"

REM A regular expression that lists all the possible names of your company's 
REM wifi networks
SETX SLACK_OFFICE_NETWORKS "(company_network1|company_network2)"

REM These are only necessary to override the defaults used in "Slack Status Update.ahk"
REM SETX SLACK_STATUS_MEETING "In a meeting|:spiral_calendar_pad:"
REM SETX SLACK_STATUS_WORKING_OFFICE "|"
REM SETX SLACK_STATUS_WORKING_REMOTELY "Working remotely|:house_with_garden:"
REM SETX SLACK_STATUS_VACATION = "Vacationing|:palm_tree:"
REM SETX SLACK_STATUS_LUNCH = "At lunch|:hamburger:"
