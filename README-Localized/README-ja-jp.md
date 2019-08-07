---
page_type: sample
products:
- office-powerpoint
- office-excel
- office-365
- office-onedrive
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  - Microsoft Graph
  services:
  - Excel
  - Office 365
  - OneDrive
  createdDate: 3/17/2016 9:42:20 AM
---
 PowerPoint アドインで Microsoft Graph を使用して Excel グラフを挿入する 

Microsoft Graph に接続し、OneDrive for Business に保存されたすべてのブックを検索し、Excel REST API を使用してブック内のすべてのグラフをフェッチし、Office.js を使用してグラフの画像を PowerPoint スライドに挿入する Microsoft Office アドインの作成方法について説明します。

![PowerPoint アドインで Microsoft Graph を使用して Excel グラフを挿入するサンプル](../images/InsertChart.png)

## はじめに

オンライン サービス プロバイダーからのデータを統合すると、アドインの価値が向上し、採用できる機会が増えます。このコード サンプルでは、Microsoft Graph にアドインを接続する方法を示します。このコード サンプルを使用して、以下を実行します。

* Office アドインから Microsoft Graph に接続します。
* MSAL .NET ライブラリを使用して、アドインに OAuth 2.0 承認フレームワークを実装します。
* Microsoft Graph から Excel および OneDrive の REST API を使用します。
* Office UI 名前空間を使用してダイアログを表示します。
* ASP.NET MVC、MSAL、Office.js を使用してアドインをビルドします。 
* PowerPoint でアドイン コマンドを使用します。


## 前提条件

このコード サンプルを実行するには、以下が必要です。

* Visual Studio 2019 以降。

* SQL Server Express (最新バージョンの Visual Studio では自動的にインストールされなくなりました。)

* [Office 365 開発者プログラム](https://aka.ms/devprogramsignup)に参加すると取得できる Office 365 アカウント。Office 365 の 1 年間の無料サブスクリプションが含まれています。

* Office 365 サブスクリプションの OneDrive for Business に保存された Excel ブック (グラフ付き)。

* Windows デスクトップ用の PowerPoint (バージョン 16.0.6769.2001 以上)。
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Microsoft Azure テナント。このアドインには、Azure Active Directory (AD) が必要です。Azure AD は、アプリケーションでの認証と承認に使う ID サービスを提供します。ここでは、試用版サブスクリプションを取得できます。[Microsoft Azure](https://account.windowsazure.com/SignUp)。

## プロジェクトを構成する

1. **Visual Studio** で、**PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb** プロジェクトを選択します。**\[プロパティ]** で、**\[SSL が有効]** が **True** であることを確認します。**\[SSL URL]** で、以下の手順 3 でリストされているのと同じドメイン名とポート番号が使用されていることを確認します。
 
2. Azure サブスクリプションが Office 365 テナントにバインドされていることを確認します。詳細については、Active Directory チームのブログ投稿「[複数の Windows Azure Active Directory を作成して管理する](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx)」をご覧ください。「**新しいディレクトリを追加する**」セクションで、この方法を説明しています。また、詳しくは、「[Office 365 開発環境のセットアップ](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription)」と「**Office 365 アカウントを Azure AD と関連付けて、アプリを作成して管理する**」のセクションもご覧ください。

3. [Azure の管理ポータル](https://manage.windowsazure.com)を使用してアプリケーションを登録します。管理者または Office 365 サブスクリプションのアカウントでサインインします。アプリケーションの登録の方法については、「[Microsoft ID プラットフォームにアプリケーションを登録する](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually)」を参照してください。次に示す設定を使用します。

 - REDIRCT URI: https://localhost:44301/AzureADAuth/Authorize	
 - サポートされているアカウントの種類:"この組織のディレクトリ内のアカウントのみ"
 - 暗黙的な付与:暗黙的な付与オプションを有効にしない
 - API アクセス許可:**Files.Read.All** と **User.Read**

	> 注:注: アプリケーションを登録したら、Azure の管理ポータルにある \[アプリの登録] の **\[概要]** ブレードの**アプリケーション (クライアント) ID** と**ディレクトリ (テナント) ID** をコピーします。**\[証明書とシークレット]** ブレードでクライアント シークレットを作成したら、それもコピーします。 
	 
4.  web.config で、前の手順でコピーした値を使用します。**\[AAD:ClientID]** にクライアント ID、**\[AAD:ClientSecret]** にクライアント シークレット、**\[AAD:O365TenantID]** にテナント ID を設定します。 

## プロジェクトを実行する
1. Visual Studio ソリューション ファイルを開きます。 
2. **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** を右クリックし、**\[スタートアップ プロジェクトに設定]** を選択します。
2. F5 キーを押します。 
3. PowerPoint で、**\[挿入]** タブを開き、**\[グラフの選択]** を選択し、作業ウィンドウ アドインを開きます。

## 既知の問題

* シナリオ:サンプル コードを実行しようとしても、アドインが読み込まれません。
	* 解決方法: 
		1. Visual Studio で **SQL Server オブジェクト エクスプローラー**を開きます。
		2. **(localdb)\\MSSQLLocalDB** > **\[データベース]** の順に展開します。
		3. **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** を右クリックし、**\[削除]** を選択します。 
* シナリオ:コード サンプルを実行すると、*Office.context.ui.messageParent* の行でエラーが発生します。	
	* 解決方法:サンプル コードの実行を停止して再起動します。 
* zip ファイルをダウンロードし、そのファイルを解凍するときに、ファイル パスが長すぎることを示すエラーが発生します。
	* 解決方法:ルート直下のフォルダー (例: C:\\sample) にファイルを解凍します。

## 質問とコメント
*PowerPoint アドインで Microsoft Graph を使用して Excel グラフを挿入する*サンプルに関するフィードバックをお寄せください。このリポジトリの「*問題*」セクションでフィードバックを送信できます。Office 365 開発全般の質問につきましては、「[Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API)」に投稿してください。質問には、\[office-js]、\[MicrosoftGraph]、\[API] のタグを付けてください。

## その他の技術情報

* [Microsoft Graph (Excel) ToDo コード サンプル](https://github.com/microsoftgraph/aspnet-todo-rest-sample)
* [Microsoft Graph ドキュメント](https://docs.microsoft.com/en-us/graph/)
* [Office アドイン ドキュメント](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)

## 著作権
Copyright (c) 2016 - 2019 Microsoft Corporation.All rights reserved.



このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/)が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
