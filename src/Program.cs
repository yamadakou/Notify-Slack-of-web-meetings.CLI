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
using Notify_Slack_of_web_meetings.CLI.SlackChannels;
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

                #region 引数の値でSlackチャンネル情報を登録

                var addSlackChannel = new SlackChannel()
                {
	                Name = opts.Name,
	                WebhookUrl = opts.WebhookUrl,
	                RegisteredBy = opts.RegisteredBy
                };
                var endPointUrl = $"{opts.EndpointUrl}SlackChannels";
                var postData = JsonConvert.SerializeObject(addSlackChannel);
                var postContent = new StringContent(postData, Encoding.UTF8, "application/json");
                var response = s_HttpClient.PostAsync(endPointUrl, postContent).Result;
                var addSlackChannelString = response.Content.ReadAsStringAsync().Result;

                // Getしたコンテンツはメッセージ+Jsonコンテンツなので、Jsonコンテンツだけ無理やり取り出す
                var addSlackChannels = JsonConvert.DeserializeObject<SlackChannel>(addSlackChannelString);

                #endregion

                #region 登録したSlackチャンネル情報のIDと引数のWeb会議情報通知サービスのエンドポイントURLをsetting.jsonに保存

                var setting = new Setting()
                {
	                SlackChannelId = addSlackChannels.Id,
	                Name = addSlackChannel.Name,
	                RegisteredBy = addSlackChannel.RegisteredBy,
	                EndpointUrl = opts.EndpointUrl
                };

                // jsonに設定を出力
                var settingJsonString = JsonConvert.SerializeObject(setting);
                if (File.Exists(opts.Filepath))
                {
	                File.Delete(opts.Filepath);
                }

                using (var fs = File.CreateText(opts.Filepath))
                {
	                fs.WriteLine(settingJsonString);
                }

                #endregion

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

                DateTime startDate = DateTime.Today.AddDays(1);
                DateTime endDate = startDate.AddDays(1);
                Outlook.Items nextOperatingDayAppointments = GetAppointmentsInRange(calFolder, startDate, endDate);

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
                string fileContent;
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

                var endPointUrl = $"{setting.EndpointUrl}WebMeetings";
                var getEndPointUrl = $"{endPointUrl}?fromDate={startDate}&toDate={endDate}";
                var getWebMeetingsResult = s_HttpClient.GetAsync(getEndPointUrl).Result;
                var getWebMeetingsString = getWebMeetingsResult.Content.ReadAsStringAsync().Result;
                
                // Getしたコンテンツはメッセージ+Jsonコンテンツなので、Jsonコンテンツだけ無理やり取り出す
                var getWebMeetings = JsonConvert.DeserializeObject<List<WebMeeting>>(getWebMeetingsString);

                foreach (var getWebMeeting in getWebMeetings)
                {
	                var deleteEndPointUrl = $"{endPointUrl}/{getWebMeeting.Id}";
	                s_HttpClient.DeleteAsync(deleteEndPointUrl).Wait();
                }

                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているエンドポイントURLを使い、Web会議情報を登録

                // Web会議情報を登録
                foreach (var addWebMeeting in addWebMeetings)
                {
	                var postData = JsonConvert.SerializeObject(addWebMeeting);
	                var postContent = new StringContent(postData, Encoding.UTF8, "application/json");
	                var response = s_HttpClient.PostAsync(endPointUrl, postContent).Result;
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
