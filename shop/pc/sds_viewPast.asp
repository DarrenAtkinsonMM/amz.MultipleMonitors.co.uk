<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2014. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
Server.ScriptTimeout = 5400
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/common.asp"-->
<%
dim iPageCurrent

Const iPageSize=20
Dim SelectTop,SelectTop1
SelectTop=500

if request.querystring("iPageCurrent")="" or request.querystring("iPageCurrent")="0" then
	iPageCurrent=1
	session("sds_PageCount")=""
	SelectTop1=SelectTop
else
	iPageCurrent=Request.QueryString("iPageCurrent")
	SelectTop1=iPageCurrent*iPageSize
end if

tmpSQL=""
if SelectTop1>"0" then
	tmpSQL=" TOP " & SelectTop1
end if

If statusAPP="1" Then
	query="SELECT DISTINCT " & tmpSQL  & " orders.idOrder, orders.orderDate, orders.ord_OrderName, orders.OrderStatus,pcDropShippersOrders.pcDropShipO_OrderStatus FROM pcDropShippersSuppliers,Products,ProductsOrdered,orders LEFT OUTER JOIN pcDropShippersOrders ON (pcDropShipO_idOrder=orders.IdOrder AND pcDropShipO_DropShipper_ID=" & session("pc_idsds") & ") WHERE ProductsOrdered.idOrder = orders.idOrder AND ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & " AND products.idproduct=ProductsOrdered.idproduct  AND ((pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct) OR (pcDropShippersSuppliers.idproduct=products.pcprod_ParentPrd)) AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & " AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) ORDER BY orders.idOrder DESC;"
Else
	query="SELECT DISTINCT " & tmpSQL  & " orders.idOrder, orders.orderDate, orders.ord_OrderName, orders.OrderStatus,pcDropShippersOrders.pcDropShipO_OrderStatus FROM (pcDropShippersSuppliers INNER JOIN (orders INNER JOIN ProductsOrdered ON orders.idorder = ProductsOrdered.idorder) ON (pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper = " & session("pc_sdsIsDropShipper") & ")) LEFT OUTER JOIN pcDropShippersOrders ON (pcDropShipO_idOrder=orders.IdOrder AND pcDropShipO_DropShipper_ID=" & session("pc_idsds") & ") WHERE ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & " AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) ORDER BY orders.idOrder DESC;"
end if
set rstemp=Server.CreateObject("ADODB.Recordset")

rstemp.CursorLocation=adUseClient
rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, conntemp

if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rstemp.eof then
	set rstemp=nothing
	call closeDb()
 	response.redirect "msg.asp?message=34"     
else
	rstemp.MoveFirst
	' get the max number of pages
	Dim iPageCount
	if session("sds_PageCount")<>"" then
		iPageCount=session("sds_PageCount")
	else
		iPageCount=rstemp.PageCount
		session("sds_PageCount")=iPageCount
	end if
	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
	If iPageCurrent < 1 Then iPageCurrent=1
	' set the absolute page
	rstemp.AbsolutePage=iPageCurrent
end if          

%> 

<!--#include file="header_wrapper.asp"-->
<div id="pcMain">
	<div class="pcMainContent">   
		<h1><%= dictLanguage.Item(Session("language")&"_CustviewPast_4")%> (TOP <%=SelectTop%>)</h1>
       
		<div class="pcTable">
    	<div class="pcTableHeader">
				<div class="pcSdsViewPast_OrderNum"><%= dictLanguage.Item(Session("language")&"_CustviewPast_5")%></div>
				<%if scOrderNumber="1" then 'Show order name %>
					<div class="pcSdsViewPast_OrderName"><%= dictLanguage.Item(Session("language")&"_CustviewPast_9")%></div>
				<% end if %>
				<div class="pcSdsViewPast_OrderDate"><%= dictLanguage.Item(Session("language")&"_CustviewPast_6")%></div>
				<div class="pcSdsViewPast_OrderStatus"><%= dictLanguage.Item(Session("language")&"_sds_viewpast_1a")%></div>
        <div class="pcSdsViewPast_Actions">&nbsp;</div>
			</div>
			<div class="pcSpacer"></div>
			<%
        mcount=0
        pcArr=rstemp.getRows(iPageSize)
        intCount=ubound(pcArr,2)
        set rstemp=nothing
        For i=0 to intCount
          mcount=mcount+1
          pIdOrder = pcArr(0,i)
          pOrderDate = pcArr(1,i)
          pOrderName = pcArr(2,i)
          pOrderStatus= pcArr(3,i)
          pcDropShipO_OrderStatus=pcArr(4,i)
          if not IsNull(pcDropShipO_OrderStatus) then
            pOrderStatus=pcDropShipO_OrderStatus
          end if
          
          if IsNull(pOrderStatus) or pOrderStatus="" then
            pOrderStatus=0
          end if
      	%>
        <div class="pcTableRow">
          <div class="pcSdsViewPast_OrderNum">
            <a href="sds_viewPastD.asp?idOrder=<%= (scpre+int(pIdOrder))%>"><%= (scpre+int(pIdOrder))%></a>
          </div>
          <%if scOrderNumber="1" then 'Show order name %>
            <div class="pcSdsViewPast_OrderName">
              <%=pOrderName%>
            </div>
          <% end if %>
          <div class="pcSdsViewPast_OrderDate">
            <%=showdateFrmt(pOrderDate)%>
          </div>
          <div class="pcSdsViewPast_OrderStatus">
            <%Select Case pOrderStatus
            Case 2: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_2")
            Case 3: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_3")
            Case 4: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_4")
            Case 5: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_5")
            Case 6: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_6")
            Case 9: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_9")
            Case 10: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_4")
            Case 12: response.write dictLanguage.Item(Session("language")&"_sds_viewpast_4")						
            Case Else:
              queryQ="SELECT TOP 1 idProductOrdered FROM ProductsOrdered WHERE idOrder=" & pIdOrder & " AND pcDropShipper_ID=" & session("pc_idsds") & " AND pcPrdOrd_Shipped=0;"
              set rsQ=connTemp.execute(queryQ)
              if not rsQ.eof then
                queryQ="SELECT TOP 1 idProductOrdered FROM ProductsOrdered WHERE idOrder=" & pIdOrder & " AND pcDropShipper_ID=" & session("pc_idsds") & " AND pcPrdOrd_Shipped=1;"
                set rsQ=connTemp.execute(queryQ)
                if not rsQ.eof then
                  response.write dictLanguage.Item(Session("language")&"_sds_viewpast_7")
                else
                  response.write dictLanguage.Item(Session("language")&"_sds_viewpast_3")
                end if
                set rsQ=nothing
              else
                response.write dictLanguage.Item(Session("language")&"_sds_viewpast_4")
              end if
              set rsQ=nothing
            End Select%>
          </div>
          <div class="pcSdsViewPast_Actions">
            <div class="pcSmallText">
              <a href="sds_viewPastD.asp?idOrder=<%= (scpre+int(pIdOrder))%>"><%= dictLanguage.Item(Session("language")&"_CustviewPast_3")%></a><%if pOrderStatus="3" or pOrderStatus="7" or pOrderStatus="8" then%> - <a href="sds_ShipOrderWizard1.asp?idOrder=<%=pIdOrder%>"><%= dictLanguage.Item(Session("language")&"_sds_viewpast_1c")%></a><%end if%>
            </div>
          </div>
        </div>
				<%
				Next
				%>
				<div class="pcTableRowFull"><hr /></div>
			</div>
			<div class="pcClear"></div>
  			<% 
			set rstemp = nothing
			%>
        
		<% 
        if iPageCount>1 Then 
        %>

			<div class="pcPageNavigation"> 
				<%Response.Write("Page "& iPageCurrent & " of "& iPageCount & "<br />")%>
				<%'Display Next / Prev buttons
				if iPageCurrent > 1 then
					'We are not at the beginning, show the prev button %>
					<a href="sds_viewPast.asp?iPageCurrent=<%=iPageCurrent-1%>"><img src="<%=pcf_getImagePath("../pc/images","prev.gif")%>" width="10" height="10"></a> 
				<% end If
				If iPageCount <> 1 then
					For I=1 To iPageCount
						If int(I)=int(iPageCurrent) Then %>
							<%=I%> 
						<% Else %>
							<a href="sds_viewPast.asp?iPageCurrent=<%=I%>" style="text-decoration: underline;"><%=I%></a> 
						<% End If %>
					<% Next %>
				<% end if %>
				<% if CInt(iPageCurrent) <> CInt(iPageCount) then
					'We are not at the end, show a next link %>
					<a href="sds_viewPast.asp?iPageCurrent=<%=iPageCurrent+1%>"><img src="<%=pcf_getImagePath("../pc/images","next.gif")%>" width="10" height="10"></a> 
				<% end If %>
			</div>
    <% end if %>
        
    <div class="pcFormButtons">
      <a class="pcButton pcButtonBack" href="sds_MainMenu.asp">
        <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
        <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
      </a>
  	</div>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->