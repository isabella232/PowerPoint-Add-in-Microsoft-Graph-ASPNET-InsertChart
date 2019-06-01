// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
//using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using PowerPointAddinMicrosoftGraphASPNETInsertChart.Helpers;
using PowerPointAddinMicrosoftGraphASPNETInsertChart.Models;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace PowerPointAddinMicrosoftGraphASPNETInsertChart.Controllers
{


    public class AzureADAuthController : Controller
    {
        // The URL that auth should redirect to after a successful login.
        Uri loginRedirectUri => new Uri(Url.Action(nameof(Authorize), "AzureADAuth", null, Request.Url.Scheme));

        // The URL to redirect to after a logout.
        Uri logoutRedirectUri => new Uri(Url.Action(nameof(HomeController.Index), "Home", null, Request.Url.Scheme));

        /// <summary>
        /// Logs the user out.
        /// </summary>
        /// <returns>Redirect to Azure logout.</returns>
        public ActionResult Logout()
        {
            var userAuthStateId = Settings.GetUserAuthStateId(ControllerContext.HttpContext);
            Data.DeleteUserSessionToken(userAuthStateId, Settings.AzureADAuthority);
            Response.Cookies.Clear();

            return Redirect(Settings.AzureADLogoutAuthority + logoutRedirectUri.ToString());
        }

        /// <summary>
        /// Logs the user into Office 365.
        /// </summary>
        /// <param name="authState">The login or logout status of the user.</param>
        /// <returns>A redirect to the Office 365 login page.</returns>
        public async Task<ActionResult> Login(string authState)
        {
            if (string.IsNullOrEmpty(Settings.AzureADClientId) || string.IsNullOrEmpty(Settings.AzureADClientSecret))
            {
                ViewBag.Message = "Please set your client ID and client secret in the Web.config file";
                return View();
            }

            ConfidentialClientApplicationBuilder clientBuilder = ConfidentialClientApplicationBuilder.Create(Settings.AzureADClientId);
            ConfidentialClientApplication clientApp = (ConfidentialClientApplication)clientBuilder.Build();

            // Generate the parameterized URL for Azure login.
            string[] graphScopes = { "Files.Read.All", "User.Read" };
            var urlBuilder = clientApp.GetAuthorizationRequestUrl(graphScopes);
            urlBuilder.WithRedirectUri(loginRedirectUri.ToString());
            urlBuilder.WithAuthority(Settings.AzureADAuthority);
            urlBuilder.WithExtraQueryParameters("state=" + authState);
            var authUrl = await urlBuilder.ExecuteAsync(System.Threading.CancellationToken.None);

            // Redirect the browser to the login page, then come back to the Authorize method below.
            return Redirect(authUrl.ToString());

        }

        /// <summary>
        /// Authorizes the web application (not the user) to access Microsoft Graph resources by using
        /// the Authorization Code flow of OAuth.
        /// </summary>
        /// <returns>The default view.</returns>
        public async Task<ActionResult> Authorize()
        {

            ConfidentialClientApplicationBuilder clientBuilder = ConfidentialClientApplicationBuilder.Create(Settings.AzureADClientId);
            clientBuilder.WithClientSecret(Settings.AzureADClientSecret);
            clientBuilder.WithRedirectUri(loginRedirectUri.ToString());
            clientBuilder.WithAuthority(Settings.AzureADAuthority);
            ConfidentialClientApplication clientApp = (ConfidentialClientApplication)clientBuilder.Build();

            string[] graphScopes = { "Files.Read.All", "User.Read" };


            var authStateString = Request.QueryString["state"];
            var authState = JsonConvert.DeserializeObject<AuthState>(authStateString);
            try
            {
                // Get and save the token.
                var authResultBuilder = clientApp.AcquireTokenByAuthorizationCode(
                    graphScopes,
                    Request.Params["code"]   // The auth 'code' parameter from the Azure redirect.
                );

                var authResult = await authResultBuilder.ExecuteAsync();

                await SaveAuthToken(authState, authResult);

                authState.authStatus = "success";

            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                authState.authStatus = "failure";
            }

            // Instead of doing a server-side redirect, we have to do a client-side redirect to get around  
            // some issues with the display dialog API not getting properly wired up after a server-side redirect  
            var redirectUrl = Url.Action(nameof(AuthorizeComplete), new { authState = JsonConvert.SerializeObject(authState) });
            ViewBag.redirectUrl = redirectUrl;
            return View();

        }

        /// <summary>
        /// Stores the access token in a local database. 
        /// </summary>
        /// <param name="authState">Contains user's session ID.</param>
        /// <param name="authResult">The results of the attempt to get the access token.</param>
        /// <returns></returns>
        private static async Task SaveAuthToken(AuthState authState, AuthenticationResult authResult)
        {
            var idToken = SessionToken.ParseJwtToken(authResult.IdToken);
            string username = null;
            var userNameClaim = idToken.Claims.FirstOrDefault(x => x.Type == "preferred_username");
            if (userNameClaim != null)
                username = userNameClaim.Value;

            using (var db = new AddInContext())
            {
                var token = new SessionToken()
                {
                    Id = authState.stateKey,
                    CreatedOn = DateTime.Now,
                    AccessToken = authResult.AccessToken,
                    Provider = Settings.AzureADAuthority,
                    Username = username
                };
                db.SessionTokens.Add(token);
                await db.SaveChangesAsync();
            }
        }

        /// <summary>
        /// Changes the view in the pop-up to tell the user that authentication of the user
        /// and authorization of the web application are finished. 
        /// </summary>
        /// <param name="authState">The login or out status of the user.</param>
        /// <returns>The default view.</returns>
        public ActionResult AuthorizeComplete(string authState)
        {
            ViewBag.AuthState = authState;
            return View();
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
