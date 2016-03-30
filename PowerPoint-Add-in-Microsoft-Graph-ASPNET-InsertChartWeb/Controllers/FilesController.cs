// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using System.Threading.Tasks;
using System.Web.Mvc;
using PowerPointAddinMicrosoftGraphASPNETInsertChart.Helpers;
using PowerPointAddinMicrosoftGraphASPNETInsertChart.Models;

namespace PowerPointAddinMicrosoftGraphASPNETInsertChart.Controllers
{
    public class FilesController : Controller
    {      
        /// <summary>
        /// Recursively searches OneDrive for Business.
        /// </summary>
        /// <returns>A view listing the workbooks in OneDrive for Business.</returns>
        public async Task<ActionResult> OneDriveFiles()
        {
            // Get access token
            var token = Data.GetUserSessionToken(Settings.GetUserAuthStateId(ControllerContext.HttpContext), Settings.AzureADAuthority);

            // Get all the Excel files in OneDrive for Business by using the Microsoft Graph API. Select only properties needed.
            var fullWorkbooksSearchUrl = GraphApiHelper.GetWorkbookSearchUrl("?$select=name,id");
            var getFilesResult = await ODataHelper.GetItems<ExcelWorkbook>(fullWorkbooksSearchUrl, token.AccessToken);
            return View(getFilesResult);
        }      
    }
}
/*

PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart, https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart
 
Copyright (c) Microsoft Corporation
All rights reserved. 
 
MIT License:
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:
 
The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.    
  
*/