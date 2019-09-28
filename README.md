---
page_type: sample
products:
- office-powerpoint
- office-excel
- office-365
- office-onedrive
- ms-graph
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
 # Insert Excel charts using Microsoft Graph in a PowerPoint Add-in 

Learn how to build a Microsoft Office Add-in that connects to Microsoft Graph, finds all workbooks stored in OneDrive for Business, fetches all charts in the workbooks using the Excel REST APIs, and inserts an image of a chart into a PowerPoint slide using Office.js.

![Insert Excel charts using Microsoft Graph in a PowerPoint Add-in sample](images/InsertChart.png)

## Introduction

Integrating data from online service providers increases the value and adoption of your add-ins. This code sample shows you how to connect your add-in to Microsoft Graph. Use this code sample to:

* Connect to Microsoft Graph from an Office Add-in.
* Use the MSAL .NET Library to implement the OAuth 2.0 authorization framework in an add-in.
* Use the Excel and OneDrive REST APIs from Microsoft Graph.
* Show a dialog using the Office UI namespace.
* Build an Add-in using ASP.NET MVC,MSAL, and Office.js. 
* Use add-in commands in PowerPoint.


## Prerequisites

To run this code sample, the following are required.

* Visual Studio 2019 or later.

* SQL Server Express (No longer automatically installed with recent versions of Visual Studio.)

* An Office 365 account which you can get by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365.

* Excel workbooks (with charts) stored on OneDrive for Business in your Office 365 subscription.

* PowerPoint for Windows Desktop, version 16.0.6769.2001 or higher.
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* A Microsoft Azure Tenant. This add-in requires Azure Active Directiory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Configure the project

1. Ensure your Azure subscription is bound to your Office 365 tenant. For more information, see the Active Directory team's blog post, [Creating and Managing Multiple Windows Azure Active Directories](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). The section **Adding a new directory** will explain how to do this. You can also see [Set up your Office 365 development environment](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) and the section **Associate your Office 365 account with Azure AD to create and manage apps** for more information.

2. Register your application using the [Azure Management Portal](https://manage.windowsazure.com). Sign-in with the account of an administrator or your Office 365 subscription. To learn how to register your application, see [Register an application with the Microsoft identity platforml](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually). Use the following settings:

 - REDIRCT URI: https://localhost:44301/AzureADAuth/Authorize	
 - SUPPORTED ACCOUNT TYPES: "Accounts in this organizational directory only"
 - IMPLICIT GRANT: Do not enable any Implicit Grant options
 - API PERMISSIONS: **Files.Read.All** and **User.Read**

	> Note: After you register your application, copy the **Application (client) ID** and the **Directory (tenant) ID** on the **Overview** blade of the App Registration in the Azure Management Portal. When you create the client secret on the **Certificates & secrets** blade, copy it too. 
	 
3.  In web.config, use the values that you copied in the previous step. Set **AAD:ClientID** to your client id, set **AAD:ClientSecret** to your client secret, and set **"AAD:O365TenantID"** to your tenant ID. 

4. Open the solution in **Visual Studio**, and open **Solution Explorer** and then choose the **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb** project. In **Properties**, ensure **SSL Enabled** is **True**. Verify that the **SSL URL**property uses the same domain name and port number as those listed in step 2 above.

5. In **Solution Explorer**, right-click the topmost node -- the **Solution ...** node. Select **Set Startup Projects**. In the dialog that opens, expand **Common Properites** and select **Startup Project**. Enable **Multiple startup projects**. Ensure that the project whose name ends with "Web" is listed first and that both projects are set to **Start** in the **Action** column. 

## Run the project

1. Build the solution.
2. Press F5. 
3. In PowerPoint, open the **Insert** tab, and select **Pick a chart** to open the task pane add-in. The home page provides instructions.

## Known issues

* Scenario: When trying to run the code sample, the add-in will not load.
	* Resolution: 
		1. In Visual Studio, open **SQL Server Object Explorer**.
		2. Expand **(localdb)\MSSQLLocalDB** > **Databases**.
		3. Right click **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart**, then choose **Delete**. 
* Scenario: When you run the code sample, you get an error on the line *Office.context.ui.messageParent*.	
	* Resolution: Stop running the code sample and restart it. 
* If download the zip file, when you extract the files you get an error indicating that the file path is too long.
	* Resolution: Unzip your files to a folder directly under the root (e.g. c:\sample).

## Questions and comments

We'd love to get your feedback about the *Insert Excel charts using Microsoft Graph in a PowerPoint Add-in* sample. You can send your feedback to us in the *Issues* section of this repository.
Questions about Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Ensure your questions are tagged with [office-js], [MicrosoftGraph] and [API].

## Additional resources

* [Microsoft Graph (Excel) ToDo code sample](https://github.com/microsoftgraph/aspnet-todo-rest-sample)
* [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph/)
* [Office Add-ins documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)

## Copyright

Copyright (c) 2016 - 2019 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
