// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

"use strict";

Office.initialize = function () {
    $(document).ready(function () {
        app.initialize();

        app.previousSelection = '';

        // Toggles the selected state of the list view used to display chart details.
        app.toggleSelection = function (event) {

            // We need a roundabout way to get a reference to the <div> that is selected 
            // because its id is variable, its formatting classes can change, and any of 
            // its many descendant elements can be the one that the user actually clicks. 
            
            // First, get a reference to the parent table cell (<td>) of the selected chart.
            var cell = $(event.target).closest('td');

            // Then get the first child of the cell, since this is the <div> that contains the
            // chart image, name, and check box. 
            // The children(":first") method returns an array with just one member.
            var items = $(cell).children(":first");
            var containingDiv = items[0];

            var currentSelection;

            if (containingDiv !== null && containingDiv !== undefined) {
                currentSelection = containingDiv.id;
            }

            // If the user selects one chart and then changes to another reverse, the selection
            // status of both.
            if (app.previousSelection !== '' && currentSelection !== app.previousSelection) {
                $('#' + currentSelection).attr('class', 'ms-ListItem is-selected is-selectable');
                $('#' + app.previousSelection).attr('class', 'ms-ListItem is-selectable');

            } else { // Either this is the first time any chart is clicked or the same chart
                     // has been clicked twice in a row.

                // If the same chart is clicked twice in a row, toggle its selection status.
                if ($('#' + currentSelection).hasClass("is-selected")) {
                    $('#' + currentSelection).attr('class', 'ms-ListItem is-selectable');
                } else {
                    $('#' + currentSelection).attr('class', 'ms-ListItem is-selected is-selectable');
                }
            }
            app.previousSelection = currentSelection;
        };

        // Assign the handler to the <td> (table cell) that contains the image AND 
        // all its children, including the chart name, the chart image, the chart check box,
        // and the selectable <div>.
        $('.selectedCell').click(app.toggleSelection);

        // Inserts the chart into the slide. Displays success or failure message in the footer.
        app.insertChart = function () {

            // Test to ensure that the last click did not simply clear a previously selected
            // chart. If it did, then nothing is selected and the button should do nothing.
            if ($('#' + app.previousSelection).hasClass("is-selected")) {

                var base64EncodedImageStr = $('#txt' + app.previousSelection).val();

                Office.context.document.setSelectedDataAsync(base64EncodedImageStr, {
                    coercionType: Office.CoercionType.Image,
                    imageLeft: 50,
                    imageTop: 50,
                    imageWidth: 600
                }, function (asyncResult) {               
                });
           }
        };

        // Assign the button click handler.
        $('#btnInsertChart').click(app.insertChart);
    });
};

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



