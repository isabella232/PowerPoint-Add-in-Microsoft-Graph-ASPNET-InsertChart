// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using PowerPointAddinMicrosoftGraphASPNETInsertChart.Helpers;
using PowerPointAddinMicrosoftGraphASPNETInsertChart.Models;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace PowerPointAddinMicrosoftGraphASPNETInsertChart.Controllers
{
    public class ChartsController : Controller
    {
        /// <summary>
        /// Gets all the charts in the specified Excel workbook and presents them in a view.
        /// </summary>
        /// <param name="id">The internal ID of the workbook.</param>
        /// <returns>The view with the list of charts.</returns>
        public async Task<ActionResult> Index(string id)
        {
            // Get access token from the local database
            var token = Data.GetUserSessionToken(Settings.GetUserAuthStateId(ControllerContext.HttpContext), Settings.AzureADAuthority);

            var sheetsUrl = GraphApiHelper.GetSheetsWithChartsUrl(id, "&$select=name,id");
            var sheets = await ODataHelper.GetItems<ExcelSheet>(sheetsUrl, token.AccessToken);

            // Merge the charts from each worksheet into a single list
            List<Chart> allChartsInWorkbook = new List<Chart>();

            foreach (var sheet in sheets)
            {
                var chartsFromSheet = sheet.Charts;

                // The GetChartImage method requires a clean charts URL, that is, no $select option.
                string cleanFullChartsUrl = GraphApiHelper.GetChartsUrl(id, sheet.Id, null);

                foreach (var chart in chartsFromSheet)
                {
                   // string singleChartImageUrl = GraphApiHelper.GetSingleChartImageUrl(cleanFullChartsUrl, chart.Id);
                    chart.ImageAsBase64String = await GraphApiHelper.GetChartImage(cleanFullChartsUrl, chart.Id, token.AccessToken);
                }

                allChartsInWorkbook = allChartsInWorkbook.Concat(chartsFromSheet).ToList();
            }
            return View(allChartsInWorkbook);
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