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
 在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表 

了解如何构建连接到 Microsoft Graph 的 Microsoft Office 外接程序，查找存储在 OneDrive for Business 中的所有工作簿，使用 Excel REST API 提取工作簿中的所有图表，以及使用 Office.js 将图表的图像插入到 PowerPoint 幻灯片中。

![在 PowerPoint 外接程序示例中使用 Microsoft Graph 插入 Excel 图表](images/InsertChart.png)

## 简介

集成来自联机服务提供程序的数据可提高外接程序的价值和采用率。此代码示例演示了如何将外接程序连接到 Microsoft Graph。使用此代码示例可执行以下操作：

* 从 Office 外接程序连接到 Microsoft Graph。
* 使用 MSAL .NET 库在外接程序中实现 OAuth 2.0 授权框架。
* 从 Microsoft Graph 中使用 Excel 和 OneDrive REST API。
* 使用 Office UI 命名空间显示对话框。
* 使用 ASP.NET MVC、MSAL 和 Office.js 构建外接程序。 
* 在 PowerPoint 中使用外接程序命令。


## 先决条件

必须符合以下条件才能运行此代码示例。

* Visual Studio 2019 或更高版本。

* SQL Server Express（不再随最新版本的 Visual Studio 一起自动安装。）

* Office 365 帐户，获取方法为加入 [Office 365 开发人员计划](https://aka.ms/devprogramsignup)，其中包含为期 1 年的免费 Office 365 订阅。

* 在 Office 365 订阅的 OneDrive for Business 中存储的 Excel 工作簿（含图表）。

* PowerPoint for Windows Desktop 版本 16.0.6769.2001 或更高版本。
* [Office 开发人员工具](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* 一个 Microsoft Azure 租户。此外接程序需要 Azure Active Directiory (AD)。Azure AD 为应用程序提供了用于进行身份验证和授权的标识服务。你还可在此处获得试用订阅：[Microsoft Azure](https://account.windowsazure.com/SignUp)。

## 配置项目

1. 在 **Visual Studio** 中，选择**“PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb”**项目。在**“属性”**中，确保**“已启用 SSL”**为**“True”**。验证 **SSL URL** 使用的域名和端口号与下面步骤 3 中列出的相同。
 
2. 确保你的 Azure 订阅已绑定到 Office 365 租户。有关详细信息，请参阅 Active Directory 团队的博客文章：[创建和管理多个 Microsoft Azure Active Directory](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx)。**“添加新目录”**部分将介绍如何执行此操作。你还可以参阅[“设置 Office 365 开发环境”](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription)和**“关联你的 Office 365 帐户和 Azure AD 以创建并管理应用”**部分获取详细信息。

3. 使用 [Azure 管理门户](https://manage.windowsazure.com)注册你的应用程序。使用管理员或 Office 365 订阅的帐户登录。若要了解如何注册应用程序，请参阅[向 Microsoft 标识平台注册应用程序](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually)。使用以下设置：

 - 重定向 URI：https://localhost:44301/AzureADAuth/Authorize	
 - 支持的帐户类型：“仅限此组织目录中的帐户”
 - 隐式授权：不启用任何隐式授权选项
 - API 权限：**Files.Read.All** 和 **User.Read**

	> 注意：注册应用程序之后，复制 Azure 管理门户的**“概览”**部分上的**“应用程序(客户端) ID”**和**“目录(租户) ID”**。在**“证书和密码”**部分创建客户端密码时，同样复制该密码。 
	 
4.  在 web.config 中，使用你在上一步中复制的值。将**“AAD:ClientID”**设置为客户端 ID，将**“AAD:ClientSecret”**设置为客户端密码，并将**“AAD:O365TenantID”**设置为租户 ID。 

## 运行项目
1. 打开 Visual Studio 解决方案文件。 
2. 右键单击**“PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart”**，再选择**“设为启动项目”**。
2. 按 F5。 
3. 在 PowerPoint 中，打开**“插入”** 选项卡，并选择**“选取图表”**以打开任务窗格外接程序。

## 已知问题

* 应用场景：当尝试运行该代码示例时，外接程序不会加载。
	* 解决方案： 
		1. 在 Visual Studio 中，打开**“SQL Server 对象资源管理器”**。
		2. 展开**“(localdb) \\MSSQLLocalDB”**>**“数据库”**。
		3. 右键单击**“PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart”**，再选择**“删除”**。 
* 应用场景：运行代码示例时，*Office.context.ui.messageParent* 行出错。	
	* 解决方案：停止运行该代码示例并重启它。 
* 如果下载 zip 文件，当提取文件时出错，指示该文件路径太长。
	* 解决方案：将文件直接解压缩到根目录下的文件夹中（例如 c:\\sample）。

## 问题和意见
我们希望得到你对*“在 PowerPoint 外接程序中使用 Microsoft Graph 插入 Excel 图表”* 示例的相关反馈。可以在此存储库中的*“问题”*部分向我们发送反馈。与 Office 365 开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API)。确保你的问题使用了 \[office-js]、\[MicrosoftGraph] 和 \[API] 标记。

## 其他资源

* [Microsoft Graph (Excel) ToDo 代码示例](https://github.com/microsoftgraph/aspnet-todo-rest-sample)
* [Microsoft Graph 文档](https://docs.microsoft.com/en-us/graph/)
* [Office 外接程序文档](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)

## 版权信息
版权所有 (c) 2016 - 2019 Microsoft Corporation。保留所有权利。



此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
