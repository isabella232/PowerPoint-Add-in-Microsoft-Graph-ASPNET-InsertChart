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
 使用 Microsoft Graph 在 PowerPoint 增益集中插入 Excel 圖表 

了解如何建置 Microsoft Office 增益集，此增益集可連接到 Microsoft Graph、尋找儲存在商務用 OneDrive 中的所有活頁簿、使用 Excel REST API 擷取活頁簿中的所有圖表，以及使用 Office.js 將圖表的影像插入 PowerPoint 投影片。

![使用 Microsoft Graph 在 PowerPoint 增益集中插入 Excel 圖表範例](../images/InsertChart.png)

## 簡介

整合線上服務提供者的資料可以增加價值，以及您的增益集的採用。這個程式碼範例會示範如何將增益集連接至 Microsoft Graph。使用這個程式碼範例以︰

* 從 Office 增益集連接至 Microsoft Graph。
* 使用 MSAL.NET 程式庫實作增益集中的 OAuth 2.0 授權架構。
* 從 Microsoft Graph 使用 Excel 和 OneDrive REST API。
* 使用 Office UI 命名空間顯示對話方塊。
* 使用 ASP.NET MVC、MSAL 和 Office.js 建置增益集。 
* 在 PowerPoint 中使用增益集命令。


## 必要條件

若要使用此程式碼範例，需要有下列各項。

* Visual Studio 2019 或更新版本。

* SQL Server Express (不再自動安裝最新版本的 Visual Studio)。

* 您可以透過加入 [Office 365 開發人員計畫](https://aka.ms/devprogramsignup)取得 Office 365 帳戶，該帳戶包含 Office 365 的免費 1 年訂用帳戶。

* 您 Office 365 訂用帳戶中的商務用 OneDrive 上儲存的 Excel 活頁簿 (含圖表)。

* PowerPoint for Windows Desktop，16.0.6769.2001 版或更新版本。
* [Office 開發人員工具](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Microsoft Azure 租用戶。此增益集需要使用 Azure Active Directiory (AD)。Azure AD 會提供識別服務，以便應用程式用於驗證和授權。在這裡可以取得試用版的訂用帳戶：[Microsoft Azure](https://account.windowsazure.com/SignUp)。

## 設定專案

1. 在 **Visual Studio** 中，選擇 **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb** 專案。在 \[內容] 中，確保 \[SSL 已啟用] 為 **True**。確認 **SSL URL** 使用與以下步驟 3 中所列出項目相同的網域名稱和連接埠號碼。
 
2. 請確定您的 Azure 訂用帳戶已繫結至您的 Office 365 租用戶。如需詳細資訊，請參閱 Active Directory 小組的部落格文章：[建立和管理多個 Windows Azure Active Directory](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx)。**新增目錄**一節將說明如何執行這項操作。如需詳細資訊，也可以參閱[設定 Office 365 開發環境](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription)和**建立 Office 365 帳戶與 Azure AD 的關聯以便建立和管理應用程式**一節。

3. 使用 [Azure 管理入口網站](https://manage.windowsazure.com)註冊您的應用程式。使用系統管理員帳戶或您的 Office 365 訂用帳戶登入。若要了解如何註冊應用程式，請參閱[使用 Microsoft 身分識別平台來註冊應用程式](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually)。使用下列設定：

 - 重新導向 URI：https://localhost:44301/AzureADAuth/Authorize	
 - 支援的帳戶類型：「僅此組織目錄中的帳戶」
 - 隱含授與：請不要啟用任何 \[隱含授與] 選項
 - API 權限：**Files.Read.All** 與 **User.Read**

	> 注意：註冊您的應用程式之後，請在 Azure 管理入口網站中 \[應用程式註冊] 的 \[概觀] 刀鋒視窗上，複製 \[應用程式 (用戶端) 識別碼] 和 \[目錄 (租用戶) 識別碼]。當您在 \[憑證和祕密] 刀鋒視窗上建立用戶端密碼時，請也複製這項資訊。 
	 
4.  在 web.config 中，使用您在上一個步驟中複製的值。將 **AAD:ClientID** 設定為您的用戶端識別碼、將 **AAD:ClientSecret** 設定為您的用戶端密碼，並將 **"AAD:O365TenantID"** 設定為您的租用戶識別碼。 

## 執行專案
1. 開啟 Visual Studio 方案檔。 
2. 以滑鼠右鍵按一下 **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart**，然後選擇 \[設定為啟始專案]。
2. 按下 F5。 
3. 在 PowerPoint 中，開啟 \[插入] 索引標籤，然後選取 \[挑選圖表] 以開啟工作窗格增益集。

## 已知問題

* 案例：嘗試執行程式碼範例時，增益集將不會載入。
	* 解決方案： 
		1. 在 Visual Studio 中，開啟 \[SQL Server 物件總管]。
		2. 展開 **(localdb)\\MSSQLLocalDB** > \[資料庫]。
		3. 以滑鼠右鍵按一下 **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart**，然後選擇 \[刪除]。 
* 案例：當您執行程式碼範例時，您會在 *Office.context.ui.messageParent* 行上遇到錯誤。	
	* 解決方案：停止執行程式碼範例，然後重新啟動它。 
* 如果下載 zip 檔，當您解壓縮檔案，您會遇到錯誤，指出檔案路徑太長。
	* 解決方案：將檔案直接解壓縮到根目錄底下的資料夾 (例如. c:\\sample)。

## 問題與意見
我們樂於獲得您關於*使用 Microsoft Graph 在 PowerPoint 增益集中插入 Excel 圖表*範例的意見反應。您可以在此儲存機制的 \[問題]** 區段中，將您的意見反應傳送給我們。請在 [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API) 提出有關 Office 365 開發的一般問題。務必以 \[office-js]、\[MicrosoftGraph] 和 \[API] 標記您的問題。

## 其他資源

* [Microsoft Graph (Excel) ToDo 程式碼範例](https://github.com/microsoftgraph/aspnet-todo-rest-sample)
* [Microsoft Graph 文件](https://docs.microsoft.com/en-us/graph/)
* [Office 增益集文件](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)

## 著作權
Copyright (c) 2016 - 2019 Microsoft Corporation.著作權所有，並保留一切權利。



此專案已採用 [Microsoft 開放原始碼管理辦法](https://opensource.microsoft.com/codeofconduct/)。如需詳細資訊，請參閱[管理辦法常見問題集](https://opensource.microsoft.com/codeofconduct/faq/)，如果有其他問題或意見，請連絡 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
