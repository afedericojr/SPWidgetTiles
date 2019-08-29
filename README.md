# SPWidgetTiles - SharePoint 2016 REST API

Commonly, with a workflow intensive process in a SharePoint list, users are left just adding the list to their homepage with a custom view which is static and boring.
These tiles work nicely as a dashboard with a query for each tile to display status updates, etc, and keep it fresh. If the data isn't available, the widget collapses, and will be redrawn when it is available.
The interval timer is currently set to 5 seconds so that the tiles refresh with current data.

* @description Queries a SharePoint list and displays the query result in a widget tile.
* @description Currently accepts $select queries and /ItemCount
* @description Written in ES5 for IE11 compatability.
* @author AFEDERICO
* @todo Test against lookup fields other than a person lookup
* @requires DIV areas be defined with data attributes. Multiple divs with unique ID's may be used.
````
<div id="widget1" 
class="widget style4" 
data-heading="My Heading" 
data-labels='{"Title":"From"}' 
data-query="../_api/web/lists/getbytitle('List Name')/Items?$filter=Status eq 'Some Status'&$select=id,Title,To/Title,Field_x0020_Name, Qty, Status&$OrderBy=Id desc&$top=1&$expand=To">
</div>
````
Example use as a function for a button, etc.

```` W.queryList("<query string>","Widget Name","WidgetID",'{JSON}',@boolean); ````

 #### Note: SharePoint's lookup fields return asynchronously, which affects field load order. In order to keep the ID at the top, adjustments were made to element rendering.

## Screenshots
### (Side by side to show automatic refresh)
#### All Queries Successful
![All Queries](/screenshots/AllQueryResults.jpg)
#### Missing Query Result
![Missing Query Result](/screenshots/MissingQueryResult.jpg)
#### Recreate Tile From Query Result
![Recreate Tile From Query Result](/screenshots/RecreateTileFromQueryResult.jpg)
#### Script Snippet
![Script Snippet](/screenshots/ScriptSnippet.jpg)

