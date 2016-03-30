// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using System;
using System.Configuration;
using System.Web;

namespace PowerPointAddinMicrosoftGraphASPNETInsertChart.Helpers
{
    /// <summary>
    /// Provides management of basic user and web application authentication and authorization information. 
    /// </summary>
    public static class Settings
    {
        public static string AzureADClientId = ConfigurationManager.AppSettings["AAD:ClientID"];
        public static string AzureADClientSecret = ConfigurationManager.AppSettings["AAD:ClientSecret"];

        public static string AzureADAuthority = @"https://login.microsoftonline.com/common";
        public static string AzureADLogoutAuthority = @"https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=";
        public static string GraphApiResource = @"https://graph.microsoft.com/";

       
        /// <summary>
        /// Ensures that the current key to the SessionToken table is in the cookie.
        /// </summary>
        /// <param name="ctx">The HTTP request context.</param>
        /// <returns></returns>
        public static string GetUserAuthStateId(HttpContextBase ctx)
        {
            string id;
            if (ctx.Request.Cookies[SessionKeys.Login.UserAuthStateId] == null)
            {
                // Convert GUID to a string and format as numeral to remove hyphens.
                id = Guid.NewGuid().ToString("N");
                ctx.Response.Cookies.Add(new HttpCookie(SessionKeys.Login.UserAuthStateId)
                {
                    Expires = DateTime.Now.AddMinutes(20),
                    Value = id
                });
            }
            else
            {
                id = ctx.Request.Cookies[SessionKeys.Login.UserAuthStateId].Value;
            }

            return id;
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