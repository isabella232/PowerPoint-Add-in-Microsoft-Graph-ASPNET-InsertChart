// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using System.Threading.Tasks;

namespace PowerPointAddinMicrosoftGraphASPNETInsertChart.Helpers
{
    /// <summary>
    /// Provides methods for Microsoft Graph-specific endpoints.
    /// </summary>
    internal static class GraphApiHelper
    {
        // Microsoft Graph-related base URLs
        internal static string GetFilesUrl = @"https://graph.microsoft.com/v1.0/me/drive/root/children";
        internal static string BaseMSGraphSearchUrl = @"https://graph.microsoft.com/v1.0/me/drive/root/microsoft.graph.search";
        // **** REPLACE beta IN THE NEXT LINE WHEN NEXT GRAPH VERSION RELEASES ****
        internal static string BaseItemsUrl = @"https://graph.microsoft.com/beta/me/drive/items/";

        internal static string GetWorkbookSearchUrl(string selectedProperties)
        {
            // Construct URL to search OneDrive for Business for Excel workbooks                
            var workbooksSearchRelativeUrl = "(q = '.xlsx')";
            return BaseMSGraphSearchUrl + workbooksSearchRelativeUrl + selectedProperties;
        }

        internal static string GetSheetsWithChartsUrl(string workbookId, string selectedProperties)
        {
            // Construct URL for sheets with charts property expanded
            var chartsExpansionOption = "?$expand=charts";
            return BaseItemsUrl + workbookId + "/workbook/worksheets" + chartsExpansionOption + selectedProperties;
        }

        internal static string GetChartsUrl(string workbookId, string sheetId, string selectedProperties)
        {
            // Construct URL for charts
            var sheetsRelativeChartUrl = "('" + sheetId + "')/Charts";
            return BaseItemsUrl + workbookId + "/workbook/worksheets" + sheetsRelativeChartUrl + selectedProperties;
        }

        internal static async Task<string> GetChartImage(string chartUrl, string chartId, string accessToken)
        {
            // Create image URL.
            string strChartUrl = chartUrl + "('" + chartId + "')/Image";

            dynamic jsonData = await ODataHelper.SendRequestWithAccessToken(strChartUrl, accessToken);

            // Convert to string and use to set the Chart.ImageAsBase64String property.
            return jsonData.value.ToString();
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