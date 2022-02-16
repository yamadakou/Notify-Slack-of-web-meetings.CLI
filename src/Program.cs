using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using CommandLine;
using Newtonsoft.Json;
using Notify_Slack_of_web_meetings.CLI.Settings;
using Notify_Slack_of_web_meetings.CLI.WebMeetings;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Notify_Slack_of_web_meetings.CLI
{
    class Program
    {
	    [Verb("setting", HelpText = "Register Slack channel information and create a configuration file.")]
        public class SettingOptions
        {
            [Option('u', "url", HelpText = "The web service endpoint url.", Required = true)]
            public string EndpointUrl { get; set; }

            [Option('n', "name", HelpText = "The Slack channel name.", Required = true)]
            public string Name { get; set; }

            [Option('w', "webhookurl", HelpText = "The web hook url. (Slack incoming webhook)", Required = true)]
            public string WebhookUrl { get; set; }

            [Option('r', "register", HelpText = "The registered name.", Required = true)]
            public string RegisteredBy { get; set; }

            [Option('f', "filepath", HelpText = "Output setting file path.", Default = "./setting.json")]
            public string Filepath { get; set; }

        }
        [Verb("register", HelpText = "Register the web conference information to be notified.")]

        public class RegisterOptions
        {
            [Option('f', "filepath", HelpText = "Input setting file path.", Default = "./setting.json")]
            public string Filepath { get; set; }
        }

        private static HttpClient s_HttpClient = new HttpClient();

        static int Main(string[] args)
        {
            Func<SettingOptions, int> RunSettingAndReturnExitCode = opts =>
            {
                Console.WriteLine("Run Setting");

                return 1;
            };
            Func<RegisterOptions, int> RunRegisterAndReturnExitCode = opts =>
            {
                Console.WriteLine("Run Register");
                Console.WriteLine($"filepath:{opts.Filepath}");

                var application = new Outlook.Application();

                #region ログインユーザーのOutlookから、翌稼働日の予定を取得

                // ログインユーザーのOutlookの予定表フォルダを取得
                Outlook.Folder calFolder =
                    application.Session.GetDefaultFolder(
                            Outlook.OlDefaultFolders.olFolderCalendar)
                        as Outlook.Folder;

                DateTime start = DateTime.Today.AddDays(1);
                DateTime end = start.AddDays(1);
                Outlook.Items nextOperatingDayAppointments = GetAppointmentsInRange(calFolder, start, end);

                #endregion

                #region 取得した予定一覧の中からWeb会議情報を含む予定を抽出

                var webMeetingAppointments = new List<Outlook.AppointmentItem>();

                // ZoomURLを特定するための正規表現
                var zoomUrlRegexp = @"https?://[^(?!.*(/|.|\n).*$)]*\.?zoom\.us/[A-Za-z0-9/?=]+";

                foreach (Outlook.AppointmentItem nextOperatingDayAppointment in nextOperatingDayAppointments)
                {
                    // ZoomURLが本文に含まれる予定を正規表現で検索し、リストに詰める
                    if (Regex.IsMatch(nextOperatingDayAppointment.Body, zoomUrlRegexp))
                    {
                        webMeetingAppointments.Add(nextOperatingDayAppointment);
                    }
                }

                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているSlackチャンネル情報のIDと抽出した予定を使い、Web会議情報を作成

                // jsonファイルから設定を取り出す
                var fileContent = string.Empty;
                using (var sr = new StreamReader(opts.Filepath, Encoding.GetEncoding("utf-8")))
                {
                    fileContent = sr.ReadToEnd();
                }

                var setting = JsonConvert.DeserializeObject<Setting>(fileContent);

                // 追加する会議情報の一覧を作成
                var addWebMeetings = new List<WebMeeting>();
                foreach (var webMeetingAppointment in webMeetingAppointments)
                {
                    var url = Regex.Match(webMeetingAppointment.Body, zoomUrlRegexp).Value;
                    var name = webMeetingAppointment.Subject;
                    var startDateTime = webMeetingAppointment.Start;
                    var addWebMeeting = new WebMeeting()
                    {
                        Name = name,
                        StartDateTime = startDateTime,
                        Url = url,
                        RegisteredBy = setting.RegisteredBy,
                        SlackChannelId = setting.SlackChannelId
                    };
                    addWebMeetings.Add(addWebMeeting);
                }

                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているエンドポイントURLを使い、Web会議情報を削除



                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているエンドポイントURLを使い、Web会議情報を登録

                var postUrl = $"{setting.EndpointUrl}WebMeetings";

                // Web会議情報を登録
                foreach (var webMeeting in addWebMeetings)
                {
                    var postData = JsonConvert.SerializeObject(webMeeting);
                    var content = new StringContent(postData, Encoding.UTF8, "application/json");
                }

                #endregion

                return 1;
            };

            return CommandLine.Parser.Default.ParseArguments<SettingOptions, RegisterOptions>(args)
                .MapResult(
                    (SettingOptions opts) => RunSettingAndReturnExitCode(opts),
                    (RegisterOptions opts) => RunRegisterAndReturnExitCode(opts),
                    errs => 1);
        }

        /// <summary>
        /// Get recurring appointments in date range.
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <returns>Outlook.Items</returns>
        private static Outlook.Items GetAppointmentsInRange(
	        Outlook.Folder folder, DateTime startTime, DateTime endTime)
        {
	        string filter = "[Start] >= '"
	                        + startTime.ToString("g")
	                        + "' AND [End] <= '"
	                        + endTime.ToString("g") + "'";
	        Debug.WriteLine(filter);
	        try
	        {
		        Outlook.Items calItems = folder.Items;
		        calItems.IncludeRecurrences = true;
		        calItems.Sort("[Start]", Type.Missing);
		        Outlook.Items restrictItems = calItems.Restrict(filter);
		        if (restrictItems.Count > 0)
		        {
			        return restrictItems;
		        }
		        else
		        {
			        return null;
		        }
	        }
	        catch { return null; }
        }
    }
}
