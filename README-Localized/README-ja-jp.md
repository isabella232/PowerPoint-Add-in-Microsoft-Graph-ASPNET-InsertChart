# PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入 

Microsoft Graph に接続し、OneDrive for Business に保存されたすべてのブックを検索し、Excel REST API を使用したブック内のすべてのグラフをフェッチし、Office.js を使用してグラフのイメージの PowerPoint スライドに挿入する Microsoft Office アドインの作成方法について説明します。

![PowerPoint アドインで Microsoft Graph を使用して Excel グラフを挿入するサンプル](../images/InsertChart.png)

## 概要

オンライン サービス プロバイダーからのデータを統合すると、アドインの価値が向上し、採用できる機会が増えます。このコード サンプルでは、Microsoft Graph にアドインを接続する方法を示します。このコード サンプルを使用して、以下を実行します。

* Office アドインから Microsoft Graph に接続します。
* アドインで OAuth 2.0 認証フレームワークを使用します。
* Microsoft Graph から Excel および OneDrive の REST API を使用します。
* Office UI 名前空間を使用してダイアログを表示します。
* ASP.NET MVC と Office.js を使用してアドインをビルドします。 
* PowerPoint でアドイン コマンドを使用します。


## 前提条件
このコード サンプルを実行するには、以下が必要です。

* Visual Studio 2015。

* <a herf="https://profile.microsoft.com/RegSysProfileCenter/wizardnp.aspx?wizid=14b845d0-938c-45af-b061-f798fbb4d170&amp;lcid=1033">Office 365 開発者プログラム</a>に参加することによって取得できる Office 365 アカウント。このプログラムには、Office 365 の 1 年間の無料サブスクリプションが含まれます。

* Office 365 サブスクリプションの OneDrive for Busines に保存された Excel ブック (グラフ付き)。

* Windows デスクトップ用の PowerPoint (バージョン 16.0.6769.2001 以上)。
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Microsoft Azure テナント。このアドインには、Azure Active Directory (AD) が必要です。Azure AD は、アプリケーションが認証と承認に使用する ID サービスを提供します。試用版サブスクリプションは、[Microsoft Azure](https://account.windowsazure.com/SignUp) で取得できます。

## プロジェクトを構成する

1. **Visual Studio** で、**PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb** プロジェクトを選択します。**[プロパティ]** で、**[SSL が有効]** が **True** であることを確認します。**SSL URL** で、以下の手順 3 でリストされているのと同じドメイン名とポート番号が使用されていることを確認します。
 
2. Azure サブスクリプションが Office 365 テナントにバインドされていることを確認します。詳細については、Active Directory チームのブログ投稿「[複数の Windows Azure Active Directory を作成および管理する](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx)」を参照してください。「**新しいディレクトリを追加する**」セクションで、この方法について説明しています。また、「[Office 365 開発環境をセットアップする](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription)」や「**Office 365 アカウントを Azure AD と関連付けてアプリを作成および管理する**」セクションも参照してください。

3. [Azure 管理ポータル](https://manage.windowsazure.com)を使用してアプリケーションを登録します。アプリケーションを登録する方法については、「[Azure 管理ポータルにブラウザー ベースの Web アプリケーションを登録する](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp)」を参照してください。以下の設定を使用します。

 - SIGN-ON URL: https://localhost:44301/AzureADAuth/Authorize 
 - APP ID URI: https://localhost:44301
 - REPLY URL: https://localhost:44301/AzureADAuth/Authorize	

	> 注: アプリケーションを登録した後は、Azure 管理ポータルに表示されているクライアント ID とクライアント シークレットをコピーします。
	 
4. アプリケーションにアクセス許可を付与します。
	*  Azure 管理ポータルで、**[Active Directory]** タブと Office 365 テナントを選択します。
	*  **[アプリケーション]** タブを選択し、構成するアプリケーションをクリックします。**[構成]** を選択します。
	*  **[他のアプリケーションに対するアクセス許可]** で、**[Microsoft Graph]** を追加します。
	*  **[デリゲートされたアクセス許可]** で、**[ユーザーのファイルと共有ファイルの読み取り]** を選択します。

5.  web.config で、**[AAD:ClientID]** にクライアント ID、**[AAD:ClientSecret]** にクライアント シークレットを設定します。 

## プロジェクトを実行する
1. Visual Studio ソリューション ファイルを開きます。 
2. **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** を右クリックし、**[スタートアップ プロジェクトに設定]** を選択します。
2. F5 キーを押します。 
3. PowerPoint で、**[挿入]** > **[グラフの選択]** の順に選択し、作業ウィンドウ アドインを開きます。

## 既知の問題

* シナリオ:サンプル コードを実行しようとしても、アドインが読み込まれません。
	* 解決方法: 
		1. Visual Studio で **SQL Server オブジェクト エクスプローラー**を開きます。
		2. **[(localdb)\MSSQLLocalDB]** > **[データベース]** の順に展開します。
		3. **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** を右クリックし、**[削除]** を選択します。 
* シナリオ:コード サンプルを実行すると、「Office.context.ui.messageParent」の行でエラーが発生します。	
	* 解決方法:サンプル コードの実行を停止して再起動します。 
* zip ファイルをダウンロードし、そのファイルを解凍するときに、ファイル パスが長すぎることを示すエラーが発生します。
	* 解決方法:ルート直下のフォルダー (例: C:\sample) にファイルを解凍します。

## 質問とコメント
*PowerPoint アドインで Microsoft Graph を使用して Excel グラフを挿入する*サンプルに関するフィードバックをお寄せください。このリポジトリの「*問題*」セクションでフィードバックを送信できます。Office 365 開発全般の質問につきましては、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/Office365+API)」に投稿してください。質問には、[office-js]、[MicrosoftGraph]、[API] のタグを付けてください。

## その他の技術情報

* [Microsoft Graph (Excel) ToDo コード サンプル](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-ExcelREST-ToDo)
* [Microsoft Graph ドキュメント](https://graph.microsoft.io/en-us/docs)
* [Office アドイン ドキュメント](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* //Build からビデオをチェックアウトする - [Office プラットフォームの概要](https://channel9.msdn.com/Events/Build/2016/B872 "Office プラットフォームの概要")。

## 著作権
Copyright (c) 2016 Microsoft Corporation. All rights reserved.


