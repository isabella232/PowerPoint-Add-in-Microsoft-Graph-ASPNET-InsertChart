/// <reference path="App.js" />
/// <reference path="~/Scripts/_officeintellisense.js" />
(function () {
    "use strict";

    var _dlg;
    var redirectTo = "/files/onedrivefiles";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function () {

        $(document).ready(function () {
            app.initialize();

            // Enable the sign in button
            $(".popupButton").prop("disabled", false);

            var authState = {stateKey: stateKey};

            // The stateKey variable must be set on the parent page
            $("#loginO365PopupButton").click(function () {
                var url = "/azureadauth/login?authState=" + encodeURIComponent(JSON.stringify(authState));
                showLoginPopup(url);
            });
        });
    };

    // This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
    // and access token provider.
    function processMessage(arg) {
        console.log("Message received in processMessage: " + JSON.stringify(arg));
        if (arg.message === "success") {
            // We now have a valid access token in the database.
            _dlg.close();
            window.location.href = redirectTo;
        } else {
            // Something went wrong with authentication or the authorization of the web application.
            _dlg.close();
            app.showNotification("User authentication and application authorization", "Unable to successfully authenticate user or authorize application. Status is " + arg.message);
        }
    }

    // Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
    function showLoginPopup(url) {
        $("#connectContainer").hide();
        $("#footerButton").hide();
        $("#waitContainer").show();
        var fullUrl = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + url;

        // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
        Office.context.ui.displayDialogAsync(fullUrl,
                {height: 40, width: 40, requireHTTPS: true}, function (result) {
            console.log("Dialog has initialized. Wiring up events");
            _dlg = result.value;
            _dlg.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
        });
    }
}());

// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

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