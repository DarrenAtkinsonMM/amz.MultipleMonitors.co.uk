<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2014. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<%
pcv_IdOrder=request("idorder")
if pcv_IdOrder="" then
	pcv_IdOrder=0
end if

if pcv_IdOrder=0 then
	response.redirect "default.asp"
end if

	query="SELECT ord_OrderName FROM Orders WHERE idorder=" & pcv_IdOrder & ";"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		pcv_OrderName=rs("ord_OrderName")
	end if
	set rs=nothing%>
<div id="pcMain">
	<div class="pcMainContent">
    <h1><%= dictLanguage.Item(Session("language")&"_sds_viewpast_1c")%> - <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_1")%> <%=(scpre+int(pcv_IdOrder))%></h1>
      
    <ul class="pcShipWizardHeader">
      <li class="pcShipWizardStep1 active">
        <img src="<%=pcf_getImagePath("images","step1a.gif")%>">
        <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_3")%>
      </li>
      <li class="pcShipWizardStep2">
        <img src="<%=pcf_getImagePath("images","step2.gif")%>">
        <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_4")%>
      </li>
      <li class="pcShipWizardStep3">
        <img src="<%=pcf_getImagePath("images","step3.gif")%>">
        <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_5")%>
      </li>
    </ul>
    
    <div class="pcClear"></div>
 
	<%
	If statusAPP="1" Then
		query = "SELECT idProduct FROM ProductsOrdered where idOrder = " & pcv_IdOrder & " AND ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & ";"
		set rs=connTemp.execute(query)
		do until rs.eof
			dsdProductID = rs("idProduct")
			query = "SELECT * FROM pcDropShippersSuppliers WHERE idProduct = "&dsdProductID&" AND pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & ";"
			set rsGObj = connTemp.execute(query)
			if rsGObj.eof then
				query = "INSERT INTO pcDropShippersSuppliers (idProduct,pcDS_IsDropShipper ) VALUES ("&dsdProductID&", " & session("pc_sdsIsDropShipper") & ");"
				set rsGOb2j = connTemp.execute(query)
				set rsGOb2j = nothing
			end if
			rs.moveNext
		loop
		set rsGObj = nothing
	End If
	query="SELECT Products.idproduct,Products.Description,Products.Stock,Products.sku,Products.pcProd_IsDropShipped,Products.pcDropShipper_ID,ProductsOrdered.quantity,ProductsOrdered.pcPrdOrd_BackOrder,ProductsOrdered.pcPrdOrd_Shipped FROM pcDropShippersSuppliers INNER JOIN (Products INNER JOIN ProductsOrdered ON Products.idproduct=ProductsOrdered.idproduct) ON (pcDropShippersSuppliers.idproduct=products.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & ")  WHERE ProductsOrdered.idorder=" & pcv_IdOrder & " AND ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & ";"
	set rs=connTemp.execute(query)
	
	IF rs.eof THEN
		set rs=nothing%>
		<div class="pcErrorMessage"><%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_6")%></div>
		<br>
    <a class="pcButton pcButtonBack" href="javascript:history.go(-1);" name="Back">
			<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
			<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
    </a>
<%
	ELSE
%>

		<div class="pcSpacer"></div>
  
    <Form name="form1" method="post" action="sds_ShipOrderWizard2.asp" class="pcForms">
    <div class="pcTable">
      <div class="pcTableHeader">
        <div class="pcShipWizard_Select">&nbsp;</div>
        <div class="pcShipWizard_ProductName"><%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_7")%></div>
        <div class="pcShipWizard_Quantity"><%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_8")%></div>
        <div class="pcShipWizard_Status"><%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_9")%></div>
      </div>
      <div class="pcSpacer"></div>
      <%
      pcv_count=0
      pcv_available=0
      Do while not rs.eof
        pcv_cancheck=0
        pcv_count=pcv_count+1
        pcv_IDProduct=rs("idproduct")
        pcv_Description=rs("description")
        pcv_Stock=rs("stock")
        pcv_Sku=rs("sku")			
        pcv_IsDropShipped=rs("pcProd_IsDropShipped")
        if IsNull(pcv_IsDropShipped) or pcv_IsDropShipped="" then
          pcv_IsDropShipped=0
        end if
        pcv_IDDropShipper=rs("pcDropShipper_ID")
        if IsNull(pcv_IDDropShipper) or pcv_IDDropShipper="" then
          pcv_IDDropShipper=0
        end if
        pcv_Qty=rs("quantity")
        if IsNull(pcv_Qty) or pcv_Qty="" then
          pcv_Qty=0
        end if
        pcv_BackOrder=rs("pcPrdOrd_BackOrder")
        if IsNull(pcv_BackOrder) or pcv_BackOrder="" then
          pcv_BackOrder=0
        end if
        pcv_Shipped=rs("pcPrdOrd_Shipped")
        if IsNull(pcv_Shipped) or pcv_Shipped="" then
          pcv_Shipped=0
        end if
        
        if (pcv_Shipped=0) or (pcv_BackOrder=1) then
          pcv_cancheck=1
          pcv_available=pcv_available+1
        end if
        %>
        <div class="pcTableRow">
          <div class="pcShipWizard_Select">
            <input type="checkbox" name="C<%=pcv_count%>" value="1" <%if pcv_cancheck=1 then%><%if (clng(pcv_Stock)>=clng(pcv_Qty)) and (pcv_BackOrder=0) then%>checked<%end if%><%else%>disabled<%end if%> class="clearBorder">
            <input type="hidden" name="IDPrd<%=pcv_count%>" value="<%=pcv_IDProduct%>">
          </div>
          <div class="pcShipWizard_ProductName">
            <%if pcv_cancheck=0 then%><%end if%><%=pcv_Description%> (<%=pcv_sku%>)<%if pcv_cancheck=0 then%><%end if%>
          </div>
          <div class="pcShipWizard_Quantity">
            <%if pcv_cancheck=0 then%><%end if%><%=pcv_Qty%><%if pcv_cancheck=0 then%><%end if%>
          </div>
          <div class="pcShipWizard_Status">
            <%IF (pcv_Shipped=1) THEN%>
            <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_10")%>
            <%ELSE%>
              <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_11")%>
              <%if (pcv_BackOrder=1) then%>
              <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_12")%>
              <%end if%>
            <%END IF%>
          </div>
        </div>
      <%	rs.MoveNext
      loop
      set rs=nothing%>
      <div class="pcSpacer"></div>
      <hr>
      
    </div>
    <div class="pcFormButtons">
      <%if pcv_available>0 then%>
        <button class="pcButton pcButtonProcessShip" name="submit1" id="submit">
          <img src="<%=pcf_getImagePath("",rslayout("pcLO_processShip"))%>" alt="<%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_13")%>" />
          <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_13")%></span>
        </button>
      <%end if%>
      <a class="pcButton pcButtonBack" href="javascript:history.go(-1);" name="Back">
        <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
        <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
      </a>
    
      <input type=hidden name="count" value="<%=pcv_count%>">
      <input type=hidden name="idorder" value="<%=pcv_IdOrder%>">
    </div>
  </form>
  <%END IF%>
  <div class="pcSpacer"></div>
  <div align="center"><a href="sds_MainMenu.asp"><%= dictLanguage.Item(Session("language")&"_CustPref_1") %></a> - <a href="sds_ViewPast.asp"><%= dictLanguage.Item(Session("language")&"_sdsMain_3") %></a></div></td>
  
  </div>
</div>
<!--#include file="footer_wrapper.asp"-->