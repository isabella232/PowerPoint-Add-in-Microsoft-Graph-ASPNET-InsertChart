# <a name="insert-excel-charts-using-microsoft-graph-in-a-powerpoint-add-in"></a>使用 Microsoft Graph 在 PowerPoint 增益集中插入 Excel 圖表 

了解如何建置 Microsoft Office 增益集，連接到 Microsoft Graph、尋找儲存在商務用 OneDrive 的所有活頁簿、使用 Excel REST API 擷取活頁簿中的所有圖表，以及使用 Office.js 將圖表的影像插入 PowerPoint 投影片。

![使用 Microsoft Graph 在 PowerPoint 增益集中插入 Excel 圖表範例](../images/InsertChart.png)

## <a name="introduction"></a>簡介

整合線上服務提供者的資料可以增加價值，以及您的增益集的採用。這個程式碼範例會示範如何將增益集連接至 Microsoft Graph。使用這個程式碼範例以︰

* 從 Office 增益集連接至 Microsoft Graph。
* 在增益集中使用 OAuth 2.0 的授權架構。
* 從 Microsoft Graph 使用 Excel 和 OneDrive REST API。
* 使用 Office UI 命名空間顯示對話方塊。
* 使用 ASP.NET MVC 和 Office.js 建置增益集。 
* 在 PowerPoint 中使用增益集命令。


## <a name="prerequisites"></a>必要條件
若要使用此程式碼範例，需要有下列各項。

* Visual Studio 2015。

* 您可以透過加入 [Office 365 開發人員計畫](https://aka.ms/devprogramsignup)取得 Office 365 帳戶，該帳戶包含 Office 365 的免費 1 年訂用帳戶。

* 您的 Office 365 訂用帳戶中的商務用 OneDrive 上儲存的 Excel 活頁簿 (具有圖表)。

* PowerPoint for Windows Desktop，版本 16.0.6769.2001 或更高版本。
* [Office 開發人員工具](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Microsoft Azure 租用戶。此增益集需要 Azure Active Directory (AD)。Azure AD 會提供識別服務，以便應用程式用於驗證和授權。在這裡可以取得試用版的訂用帳戶：[Microsoft Azure](https://account.windowsazure.com/SignUp)。

## <a name="configure-the-project"></a>設定專案

1. 在 **Visual Studio** 中，選擇 **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb** 專案。在 [屬性]**** 中，確保 [SSL 已啟用]**** 為 **True**。確認 **SSL URL** 使用與以下步驟 3 中所列出的那些項目相同的網域名稱和通訊埠號碼。
 
2. 確定您的 Azure 訂用帳戶已繫結至您的 Office 365 租用戶。如需詳細資訊，請參閱 Active Directory 小組的部落格文章：[建立和管理多個 Windows Azure Active Directory](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx)。**新增目錄**一節將說明如何執行這項操作。如需詳細資訊，也可以參閱[設定 Office 365 開發環境](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription)和**建立 Office 365 帳戶與 Azure AD 的關聯以便建立和管理應用程式**一節。

3. 使用 [Azure 管理入口網站](https://manage.windowsazure.com)註冊您的應用程式。若要了解如何註冊您的應用程式，請參閱[使用 Azure 管理入口網站註冊以瀏覽器為基礎的 Web 應用程式](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp)。使用下列設定：

 - 登入 URL：https://localhost:44301/AzureADAuth/Authorize 
 - APP ID URI：https://localhost:44301
 - 回覆 URL：https://localhost:44301/AzureADAuth/Authorize 

    > 附註：註冊您的應用程式之後，複製 Azure 管理入口網站中顯示的用戶端 ID 和用戶端密碼。
     
4. 授與權限給您的應用程式。
    *  在 Azure 管理入口網站中，選取 [Active Directory]**** 索引標籤和 Office 365 租用戶。
    *  選取 [應用程式]**** 索引標籤，然後按一下您要設定的應用程式。選擇 [設定]****。
    *  在 [其他應用程式的權限]**** 中，新增 **Microsoft Graph**。
    *  在 [委派權限]**** 中，選擇 [讀取使用者檔案及與使用者共用的檔案]****。

5.  在 web.config 中，設定 **AAD:ClientID** 為您的用戶端 ID，並且設定 **AAD:ClientSecret** 為您的用戶端密碼。 

## <a name="run-the-project"></a>執行專案
1. 開啟 Visual Studio 解決方案檔案。 
2. 以滑鼠右鍵按一下 **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart**，然後選擇 [設定為啟始專案]****。
2. 按下 F5。 
3. 在 PowerPoint 中，選擇 [插入]**** > [挑選圖表]**** 以開啟工作窗格增益集。

## <a name="known-issues"></a>已知問題

* 案例：嘗試執行程式碼範例時，增益集將不會載入。
    * 解決方案： 
        1. 在 Visual Studio 中，開啟 [SQL Server 物件總管]****。
        2. 展開 **(localdb) \MSSQLLocalDB** > **資料庫**。
        3. 以滑鼠右鍵按一下 **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart**，然後選擇 [刪除]****。 
* 案例：當您執行程式碼範例時，您會在 *Office.context.ui.messageParent*行上遇到錯誤。   
    * 解決方案：停止執行程式碼範例，然後重新啟動它。 
* 如果下載 zip 檔，當您解壓縮檔案，您會遇到錯誤，指出檔案路徑太長。
    * 解決方案：將檔案直接解壓縮到根目錄底下的資料夾 (例如. c:\sample)。

## <a name="questions-and-comments"></a>問題和建議
我們樂於獲得您關於*使用 Microsoft Graph 在 PowerPoint 增益集中插入 Excel 圖表*範例的意見反應。您可以在此儲存機制的 [問題]** 區段中，將您的意見反應傳送給我們。請在 [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API) 提出有關 Office 365 開發的一般問題。務必以 [office-js]、[MicrosoftGraph] 和 [API] 標記您的問題。

## <a name="additional-resources"></a>其他資源

* [Microsoft Graph (Excel) ToDo 程式碼範例](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-ExcelREST-ToDo)
* [Microsoft Graph 文件](https://graph.microsoft.io/en-us/docs)
* [Office 增益集文件](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* 參閱 //Build - [Office Platform Overview](https://channel9.msdn.com/Events/Build/2016/B872 "Office Platform Overview") 的影片。

## <a name="copyright"></a>著作權
Copyright (c) 2016 Microsoft Corporation.著作權所有，並保留一切權利。



此專案已採用 [Microsoft 開放原始碼管理辦法](https://opensource.microsoft.com/codeofconduct/)。如需詳細資訊，請參閱[管理辦法常見問題集](https://opensource.microsoft.com/codeofconduct/faq/)，如果有其他問題或意見，請連絡 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
