// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using PowerPointAddinMicrosoftGraphASPNETInsertChart.Models;
using System.Linq;

namespace PowerPointAddinMicrosoftGraphASPNETInsertChart.Helpers
{
    /// <summary>
    /// Handles session ID storage.
    /// </summary>
    public static class Data
    {
        /// <summary>
        /// Gets the user session token from the database.
        /// </summary>
        /// <param name="userAuthSessionId"></param>
        /// <param name="provider"></param>
        /// <returns></returns>
        public static SessionToken GetUserSessionToken(string userAuthSessionId, string provider)
        {
            SessionToken st = null;
            using (var db = new AddInContext())
            {
                st = db.SessionTokens.FirstOrDefault(t => t.Id == userAuthSessionId && t.Provider == provider);
            }
            return st;
        }

        public static void DeleteUserSessionToken(string userAuthSessionId, string provider)
        {
            using (var db = new AddInContext())
            {
                var st = db.SessionTokens.Where(t => t.Id == userAuthSessionId && t.Provider == provider);
                if (st.Any())
                {
                    db.SessionTokens.RemoveRange(st);
                    db.SaveChanges();
                }
            }
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