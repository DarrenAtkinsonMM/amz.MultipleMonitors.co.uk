<%
If pcv_showContentPages&"" = "" Then
	pcv_showContentPages = false
End If
%>

<div id="pcUsefulLinks">
  <h3><%= dictLanguage.Item(Session("language")&"_usefulLinks_1") %></h3>
  <ul>
		<li><a href="default.asp"><%= dictLanguage.Item(Session("language")&"_usefulLinks_2") %></a></li>
    <li><a href="viewcategories.asp"><%= dictLanguage.Item(Session("language")&"_usefulLinks_3") %></a></li>
    <li><a href="viewbrands.asp"><%= dictLanguage.Item(Session("language")&"_usefulLinks_4") %></a></li>
		<% 
			If pcv_showContentPages Then
				'// START CONTENT PAGES
				'// Select pages compatible with customer type
				If session("customerCategory")<>0 Then ' The customer belongs to a customer category
					' Load pages accessible by ALL, plus those accessible by the customer pricing category that the customer belongs to
					If session("customerType")=0 Then
						' Customer category does NOT have wholesale privileges, so exclude those pages
						queryCustType = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType='CC_" & session("customerCategory") &"')"
					Else
						' Customer category HAS wholesale privileges, so include wholesale-only pages
						queryCustType = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType = 'W' OR pcCont_CustomerType='CC_" & session("customerCategory") &"')"
					End If
				Else
					If session("customerType")=0 Then
						' Retail customer or customer not logged in: load pages accessible by ALL
						queryCustType = " AND pcCont_CustomerType = 'ALL'"
					Else
						' Wholesale customer: load pages accessible by ALL and Wholesale customers only
						queryCustType = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType = 'W')"
					End If
				End If
      
				'// Load pages from the database: active, not excluded from navigation, and compatible with customer type 
				Call openDb()         
				sdquery="SELECT pcCont_IDPage, pcCont_PageName FROM pcContents WHERE pcCont_InActive=0 AND pcCont_MenuExclude<>1 " & queryCustType & " ORDER BY pcCont_Order ASC, pcCont_PageName ASC;"
				Set rsSideCatObj=Server.CreateObject("ADODB.RecordSet")         
				Set rsSideCatObj=connTemp.execute(sdquery)
				Do While Not rsSideCatObj.eof
					pcIntContentPageID=rsSideCatObj("pcCont_IDPage")
					pcvContentPageName=rsSideCatObj("pcCont_PageName")
					'// Call SEO Routine
					pcGenerateSeoLinks
					'//
				%>
					<li><a href="<%=pcStrCntPageLink%>"><%=pcvContentPageName%></a></li>
				<%
					rsSideCatObj.MoveNext
				Loop
				Set rsSideCatObj=nothing
				Call closeDb()
				'// END CONTENT PAGES
			End If
		%>
    <li><a href="search.asp"><%= dictLanguage.Item(Session("language")&"_usefulLinks_5") %></a></li>
    <li><a href="viewcart.asp"><%= dictLanguage.Item(Session("language")&"_usefulLinks_6") %></a></li>
    <li><a href="contact.asp"><%= dictLanguage.Item(Session("language")&"_usefulLinks_7") %></a></li>
	</ul>
</div>