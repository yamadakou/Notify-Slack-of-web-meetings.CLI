using Newtonsoft.Json;

namespace Notify_Slack_of_web_meetings.CLI.Settings
{
	/// <summary>
	/// 設定
	/// </summary>
	public class Setting
	{
		/// <summary>
		/// SlackチェンネルのID
		/// </summary>
		[JsonProperty("slackChannelId")]
		public string SlackChannelId { get; set; }

		/// <summary>
		/// チャンネル名
		/// </summary>
		[JsonProperty("name")]
		public string Name { get; set; }

		/// <summary>
		/// 登録者
		/// </summary>
		[JsonProperty("registeredBy")]
		public string RegisteredBy { get; set; }

		/// <summary>
		/// エンドポイントURL
		/// </summary>
		[JsonProperty("endpointUrl")]
		public string EndpointUrl { get; set; }
	}
}