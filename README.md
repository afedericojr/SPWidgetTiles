# SPWidgetTiles - SharePoint 2016 REST API
Queries a SharePoint list and displays the query result in a widget tile.

Currently accepts $select queries and /ItemCount.

* @description Queries a SharePoint list and displays the query result in a widget tile.
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
 #### Note: SharePoint's lookup fields return asynchronously, which affects field load order. In order to keep the ID at the top, adjustments were made to element rendering.
