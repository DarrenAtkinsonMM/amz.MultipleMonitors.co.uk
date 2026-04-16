
<div id="pcUsefulLinks">
    <h3>Useful Links</h3>
    <ul>
        <li><a href="default.asp">Store Home</a></li>
        <li><a href="viewcategories.asp">Browse Catalog</a></li>
        <li><a href="viewbrands.asp">Browse Brands</a></li>
        <% 
          	'// START CONTENT PAGES
          	'// Select pages compatible with customer type
         	if session("customerCategory")<>0 then ' The customer belongs to a customer category
              ' Load pages accessible by ALL, plus those accessible by the customer pricing category that the customer belongs to
              if session("customerType")=0 then
                  ' Customer category does NOT have wholesale privileges, so exclude those pages
                    queryCustType = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType='CC_" & session("customerCategory") &"')"
                else
                    ' Customer category HAS wholesale privileges, so include wholesale-only pages
                    queryCustType = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType = 'W' OR pcCont_CustomerType='CC_" & session("customerCategory") &"')"
                end if
            else
                if session("customerType")=0 then
                    ' Retail customer or customer not logged in: load pages accessible by ALL
                    queryCustType = " AND pcCont_CustomerType = 'ALL'"
                else
                    ' Wholesale customer: load pages accessible by ALL and Wholesale customers only
                    queryCustType = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType = 'W')"
                end if
            end if
			
			'// Load pages from the database: active, not excluded from navigation, and compatible with customer type 
			call openDb()         
			sdquery="SELECT pcCont_IDPage, pcCont_PageName FROM pcContents WHERE pcCont_InActive=0 AND pcCont_MenuExclude<>1 " & queryCustType & " ORDER BY pcCont_Order ASC, pcCont_PageName ASC;"
			set rsSideCatObj=Server.CreateObject("ADODB.RecordSet")         
			set rsSideCatObj=connTemp.execute(sdquery)
			do while not rsSideCatObj.eof
			pcIntContentPageID=rsSideCatObj("pcCont_IDPage")
			pcvContentPageName=rsSideCatObj("pcCont_PageName")
			'// Call SEO Routine
			pcGenerateSeoLinks
			'//
		%>
		<li><a href="<%=pcStrCntPageLink%>"><%=pcvContentPageName%></a></li>
		<%
			rsSideCatObj.MoveNext
			loop
			set rsSideCatObj=nothing
			call closeDb()
			'// END CONTENT PAGES
        %>
        <li><a href="search.asp">Advanced Search</a></li>
        <li><a href="viewcart.asp">View Cart</a></li>
        <li><a href="contact.asp">Contact Us</a></li>
    </ul>
</div>