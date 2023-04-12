# Notify-Slack-of-web-meetings.CLI

当日の Web 会議の情報を Slack に通知するWeb サービス「[Notify Slack of web meeting](https://github.com/yamadakou/Notify-Slack-of-web-meetings)」を利用するクライアントアプリです。

## 概要

### Notify Slack of web meeting CLI の特徴

* 当日の Web 会議の情報を Slack に通知するWeb サービス「[Notify Slack of web meeting](https://github.com/yamadakou/Notify-Slack-of-web-meetings)」を利用するためのクライアントです。
* Outlook クライアントからログインユーザーの翌日から指定日数の Web 会議情報を登録するコンソールアプリで、以下の機能を提供します。
  * 初期設定を行う Setting コマンド
  * Web会議情報を登録する Register コマンド

### 機能説明
#### Settingコマンド
* 引数の情報から通知先となる Slack チャンネル情報を Web 会議情報通知 Web サービスに登録する。
* 引数のWeb会議情報通知 Web サービスのエンドポイントと通知先となる Slack チャンネル情報の ID を保持する設定ファイル（ JSON 形式）を作成する。
* コマンド オプション
  ```shell
  -u, --url           Required. The web service endpoint url.
  -n, --name          Required. The Slack channel name.
  -w, --webhookurl    Required. The web hook url. (Slack incoming webhook)
  -r, --register      Required. The registered name.
  -f, --filepath      (Default: ./setting.json) Ourput setting file path.
  --help              Display this help screen.
  --version           Display version information.
  ```
* コマンド例
  ```shell
  notify-slack-of-web-meeting.cli.exe setting -n "SlackChannelName" -u https://・・・/api/ -w https://hooks.slack.com/services/・・・ -r "YourName"

  ```

#### Registerコマンド
* Outlook クライアントから翌稼働日の会議情報を取得する。
* 翌稼働日以降の会議を削除後に翌稼働日の会議情報を追加する。
* WebAPI 呼び出し時に必要な情報は、設定ファイル（ JSON形式 ）の情報を使用する。
* コマンド オプション
  ```shell
  -f, --filepath    (Default: ./setting.json) Input setting file path.
  -d, --days        (Default: 1) Number of days to get an appointment.
  --help            Display this help screen.
  --version         Display version information.
  ```
* コマンド例
  ```shell
  notify-slack-of-web-meeting.cli.exe register
  ```

## 利用方法
### 前提条件
* Outlook クライアントがインストールされていること。
  * 動作確認済み Outlook クライアント
    * Outlook 2016 (Windows 版)
    * Outlook for Microsoft 365 (Windows 版)
*  Web サービス「[Notify Slack of web meeting](https://github.com/yamadakou/Notify-Slack-of-web-meetings)」にアクセスするための設定ファイル「appsettings.json」を本実行ファイルと同じパスに配置すること。
    *  「appsettings.json」の書式は以下
        ```
        {
          "Instance": "https://login.microsoftonline.com/{0}",
          "Tenant": "{デーモンアプリのテナントID}",
          "ClientId": "{デーモンアプリのクライアントID}",
          "ClientSecret": "{デーモンアプリのシークレット}",
          "ApiUrl": "{ファンクションアプリのアプリケーション ID の URI}/"
        }
        ```
        * 各値の取得方法などは、「[Notify Slack of web meeting](https://github.com/yamadakou/Notify-Slack-of-web-meetings)」のページを参照
* Slackの「Incoming Webhooks」を設定し、通知先のURLを取得する。
  * 詳細は、Slackドキュメント「 [Sending messages using Incoming Webhooks](https://api.slack.com/messaging/webhooks)」参照

### 初期設定
* [Setting コマンド](Settingコマンド) を実行する。
  * 詳細は [Setting コマンド](Settingコマンド) を参照

### Web会議情報を登録（手動実行）
* [Register コマンド](Registerコマンド) を実行する。
  * 詳細は [Register コマンド](Registerコマンド) を参照

### Windows タスクスケジューラに登録（自動実行）
* Windows タスクスケジューラで Register コマンドを毎日実行するタスクを設定する。
  1. 基本タスクを作成する
     * 名前：任意のタスク名　（例：Notify Slack of web meeting CLI）
     * トリガー：毎日
       * 開始：Register コマンドを実行する時刻　（例：17:55:00）
       * 間隔：1
     * 操作：プログラムの開始
       * プログラム/スクリプト： `notify-slack-of-web-meeting.cli.exe` を選択
       * 引数の追加（オプション）：register
         * 取得日数を指定するなど、コマンドのデフォルトと異なる指定をする場合、オプション パラメータも指定する。
           * 例：register -d 3 -f test-setting.json
       * 開始： `notify-slack-of-web-meeting.cli.exe` が存在するフォルダのパス
  2. タスクの構成を最新にする設定（オプション）
      * 登録したタスクのプロパティを開く。iew=vs-2022
      * 全般タブの「構成」を「Windows 10」に変更する。
  3. ログオフ時も実行する設定（オプション）
      * 登録したタスクのプロパティを開く。
      * 全般タブの「セキュリティ オプション」グループ内の「ユーザーがログオンしているかどうかにかかわらず実行する」を選択する。

  * 参考
    * Schedule a Task
      * https://learn.microsoft.com/ja-jp/previous-versions/windows/it-pro/windows-server-2008-r2-and-2008/cc748993(v=ws.11)
    * 【Windows 10対応】タスクスケジューラで定期的な作業を自動化する
      * https://atmarkit.itmedia.co.jp/ait/articles/1305/31/news049.html

## ビルド環境
Visual Studio 2022 バージョン 17.1 以降
* .NET デスクトップ開発ワークロードがインストールされている事
### 参考
  * チュートリアル: Visual Studio を使用して .NET コンソール アプリケーションを作成する
    * https://learn.microsoft.com/ja-jp/dotnet/core/tutorials/with-visual-studio?pivots=dotnet-6-0

### ビルドと発行
1. `gir clone ・・・` などでローカルに取得し、 Visual Studio でソリューションを開く。
2. ソリューションのリビルドを行う
   * 参考
     * Visual Studio でのプロジェクトとソリューションのビルドおよびクリーン
       * https://learn.microsoft.com/ja-jp/visualstudio/ide/building-and-cleaning-projects-and-solutions-in-visual-studio?view=vs-2022
3. コンソール アプリケーションを発行する。
   * 発行先の `publish` フォルダを実行用フォルダ（任意のフォルダ）にコピーする。
   * 参考
     * チュートリアル: Visual Studio を使用して .NET コンソール アプリケーションを発行する
       * https://learn.microsoft.com/ja-jp/dotnet/core/tutorials/publishing-with-visual-studio?pivots=dotnet-6-0
  
#### COM参照
* Microsoft Outlook 16.0 Object Library
  * https://learn.microsoft.com/ja-jp/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2022#primary-interop-assemblies-for-microsoft-office-applications
  * 参考
    * 方法: 参照マネージャーを使用して参照を追加または削除する
      * https://learn.microsoft.com/ja-jp/visualstudio/ide/how-to-add-or-remove-references-by-using-the-reference-manager?view=vs-2022
    * .NET 5プロジェクトでのCOM参照の追加
      * https://opcdiary.net/net-5%e3%83%97%e3%83%ad%e3%82%b8%e3%82%a7%e3%82%af%e3%83%88%e3%81%a7%e3%81%aecom%e5%8f%82%e7%85%a7%e3%81%ae%e8%bf%bd%e5%8a%a0/

#### 依存パッケージ
※ `dotnet list package` の結果から作成
  |最上位レベル パッケージ|バージョン|Nuget|
  |:--|:--|:--|
  |CommandLineParser                             |2.8.0   |https://www.nuget.org/packages/CommandLineParser/2.8.0|
  |Microsoft.Extensions.Configuration.Binder     |7.0.0   |https://www.nuget.org/packages/Microsoft.Extensions.Configuration.Binder/7.0.0|
  |Microsoft.Extensions.Configuration.Json       |7.0.0   |https://www.nuget.org/packages/Microsoft.Extensions.Configuration.Json/7.0.0|
  |Microsoft.Identity.Web.Certificate            |1.25.10 |https://www.nuget.org/packages/Microsoft.Identity.Web.Certificate/1.25.10|
  |Microsoft.Identity.Web.TokenCache             |1.25.10 |https://www.nuget.org/packages/Microsoft.Identity.Web.TokenCache/1.25.10|
  |Newtonsoft.Json                               |13.0.1  |https://www.nuget.org/packages/Newtonsoft.Json/13.0.1|


## （関連リポジトリ）
* Notify-Slack-of-web-meetings
  * https://github.com/yamadakou/Notify-Slack-of-web-meetings
* Notify-Slack-of-web-meetings-Admin
  * https://github.com/yamadakou/Notify-Slack-of-web-meetings-Admin
