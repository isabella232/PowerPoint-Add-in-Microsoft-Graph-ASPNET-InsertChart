// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.IdentityModel.Tokens;

namespace PowerPointAddinMicrosoftGraphASPNETInsertChart.Models
{
    /// <summary>
    /// Models a row of the SessionToken table in the database.
    /// </summary>
    public class SessionToken
    {
        /// <summary>
        /// This is the user SessionID
        /// </summary>
        [Key, Column(Order = 1)]
        [MaxLength(36)]
        public string Id { get; set; }

        // The user identity provider
        [Key, Column(Order = 2)]
        [MaxLength(150)]
        public string Provider { get; set; }

        // The access token for the OData endpoint.
        public string AccessToken { get; set; }

        public DateTime CreatedOn { get; set; }

        [MaxLength(100)]
        public string Username { get; set; }

        //TODO: Validate the token so we can extract the user name and user id properties from the id_token
        public static JwtSecurityToken ParseJwtToken(string jwtToken)
        {
            JwtSecurityToken jst = new JwtSecurityToken(jwtToken);
            return jst;
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