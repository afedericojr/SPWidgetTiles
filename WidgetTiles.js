/**
 * @description Queries a SharePoint list and displays the query result in a widget tile.
 * @author AFEDERICO
 * @todo Test against lookup fields other than a person lookup
 * @requires DIV areas be defined with data attributes. Multiple divs with unique ID's may be used.
 * <div id="widget1" 
 * class="widget style4" 
 * data-heading="My Heading" 
 * data-labels='{"Title":"From"}' 
 * data-query="../_api/web/lists/getbytitle('List Name')/Items?$filter=Status eq 'Some Status'&$select=id,Title,To/Title,Field_x0020_Name, Qty, Status&$OrderBy=Id desc&$top=1&$expand=To">
 * </div>
 *
 * Note: SharePoint's lookup fields return asynchronously, which affects field load order. In order to keep the ID at the top, adjustments were made to element rendering.
 */

var W = {}; // Main widget object
document.addEventListener('DOMContentLoaded', function () {

    // Capture all instances of widgets by class
    function getWidgets() {
        var widgetsNode = document.querySelectorAll('.widget');
        var widgets = Array.prototype.slice.call(widgetsNode);
        console.log(widgets);
        for (var i = 0; i < widgets.length; i++) {
            var query = widgets[i].dataset.query;
            var heading = widgets[i].dataset.heading;
            var labelSubs = widgets[i].dataset.labels;
            var stopRefresh; 
            W.queryList(query, heading, widgets[i].id, labelSubs, stopRefresh)
        }
    }
    getWidgets(); // Run first

    // Refresh widgets at interval until manual function with stopRefresh:true is invoked
    W.widgetRefresh = setInterval(function(){
        getWidgets();
    }, 5000);

});

/**
 * Main function for processor
 * 
 * Example for manual call: W.queryList("<query string>","Widget Name","WidgetID",'{JSON DATA}',@boolean);
 * 
 * @param {*} query SP API Query String such as ../_api/web/lists/getbytitle('List Name')/Items?$filter=Status eq 'Some Status'&$select=id,Title,To/Title,Field_x0020_Name, Qty, Status&$OrderBy=Id desc&$top=1&$expand=To
 * @param {*} heading Heading for the widget
 * @param {*} widget Widget ID
 * @param {*} labelSubs JSON object of renamed columns. {"Old Name","New Name"}
 * @param {*} stopRefresh stopRefresh if true
 */
W.queryList = function queryList(query, heading, widget, labelSubs, stopRefresh) {
    labelSubs = labelSubs || ''; // set empty if null
    stopRefresh = stopRefresh || false; // set false if null
    if (stopRefresh){
        clearInterval(W.widgetRefresh); // clear widget refresh when function passes stopRefresh:true
    }
    // new request
    var xhr = new XMLHttpRequest();
    
    // Confirm minum parameters given
    if (query && heading && widget) {
        function renderData() {
            /**
            * SCAFFOLD WIDGET
            */
            W.widget = document.getElementById(widget);
            W.widget.innerHTML = "";
            var htmlSpanTitle = document.createElement('SPAN');
            var htmlSpanItemID = document.createElement('SPAN');
            var htmlUL = document.createElement('UL');
            htmlSpanTitle.setAttribute('class', 'title');
            htmlSpanItemID.setAttribute('class', 'itemID');
            htmlSpanTitle.innerHTML = heading;
            W.widget.appendChild(htmlSpanTitle);
            W.widget.appendChild(htmlSpanItemID);
            W.widget.appendChild(htmlUL);

            // Parse data from response
            W.rData = JSON.parse(xhr.response)['d'];

            // Check if there are column name substitutions
            if (labelSubs) {
                labelSubs = JSON.parse(labelSubs);
            }

            // Check if it is an ItemCount query
            if (W.rData.hasOwnProperty('ItemCount')) {
                console.log('has item count');
                W.rData = W.rData.ItemCount;
                htmlSpanItemID.innerHTML = W.rData;
            } else {
                // Render all other queries
                W.rData = W.rData.results[0];
                console.log(W.rData);
                var keys = Object.keys(W.rData);
                for (var i = 1; i < keys.length - 1; i++) {
                    var label = keys[i];
                    var column = W.rData[label];
                    if (typeof (column) == 'object') { // Check for nested lookup values such as people picker
                        column = W.rData[label]['results'][0]['Title'];
                    }
                    // Check for and set column name overrides
                    if (labelSubs.hasOwnProperty(label)) {
                        label = labelSubs[label];
                    }
                    // Always place ID in top span
                    if (label == 'Id') {
                        htmlSpanItemID.innerHTML = 'ID: ' + column;
                    } else {
                        // Render widget contents
                        var htmlLI = document.createElement('LI');
                        var htmlLiItemLabel = document.createElement('SPAN');
                        var htmlLiItemID = document.createElement('SPAN');
                        htmlLiItemLabel.setAttribute('class', 'label');
                        htmlLiItemID.setAttribute('class', 'column');
                        htmlLI.appendChild(htmlLiItemLabel);
                        htmlLI.appendChild(htmlLiItemID);
                        htmlUL.appendChild(htmlLI);
                        htmlLiItemLabel.innerHTML = label.replace('_x0020_',' ') + ': ';
                        htmlLiItemID.innerHTML = column;
                    }
                }
                console.log(widget + ' loaded');
            }
        }

        // listener
        xhr.onload = function () {
            // return data
            if (xhr.status >= 200 && xhr.status < 300) {
                // on success
                console.log('success', xhr);
                try {
                    renderData();
                } catch (err) {
                    console.log(widget + " experienced an error.")
                    W.widget.outerHTML = "";
                    console.log(widget + " removed.")
                }
            } else {
                // on fail
                console.log('The request failed', xhr.status);
            }
            // always runs
            console.log('completed onload for xhr');
        };

        // create and send GET request
        // first arg is post type (GET, POST, PUT, DELETE, etc.)
        // second arg is the endpoint
        xhr.open("GET", query);
        xhr.setRequestHeader("Accept", "application/json; odata=verbose");
        xhr.send();
    }
}