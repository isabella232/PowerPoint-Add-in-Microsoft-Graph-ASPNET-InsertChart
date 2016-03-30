// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace PowerPointAddinMicrosoftGraphASPNETInsertChart.Helpers
{
    /// <summary>
    /// Provides methods for accessing OData endpoints.
    /// </summary>
    internal static class ODataHelper
    {
        /// <summary>
        /// Gets any JSON array from any OData endpoint that requires an access token.
        /// </summary>
        /// <typeparam name="T">The .NET type to which the members of the array will be converted.</typeparam>
        /// <param name="itemsUrl">The URL of the OData endpoint.</param>
        /// <param name="accessToken">An OAuth access token.</param>
        /// <returns>Collection of T items that the caller can cast to any IEnumerable type.</returns>
        internal static async Task<IEnumerable<T>> GetItems<T>(string itemsUrl, string accessToken)
        {
            dynamic jsonData = await SendRequestWithAccessToken(itemsUrl, accessToken);

            // Convert to .NET class and populate the properties of the model objects,
            // and then populate the IEnumerable object and return it.
            JArray jsonArray = jsonData.value;
            return JsonConvert.DeserializeObject<IEnumerable<T>>(jsonArray.ToString());
        }

        /// <summary>
        /// Sends a request to the specified OData URL with the specified access token.
        /// </summary>
        /// <param name="itemsUrl">The OData endpoint URL.</param>
        /// <param name="accessToken">The access token for the endpoint resource.</param>
        /// <returns></returns>
        internal static async Task<dynamic> SendRequestWithAccessToken(string itemsUrl, string accessToken)
        {
            dynamic jsonData = null;

            using (var client = new HttpClient())
            {
                // Create and send the HTTP Request
                using (var request = new HttpRequestMessage(HttpMethod.Get, itemsUrl))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            HttpContent content = response.Content;
                            string responseContent = await content.ReadAsStringAsync();

                            jsonData = JsonConvert.DeserializeObject(responseContent);
                        }
                    }
                }
            }
            return jsonData;
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