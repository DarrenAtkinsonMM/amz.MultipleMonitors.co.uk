<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="pcStartSession.asp"-->
<%
err.number=0
dim pIdOrder
pIdOrder=getUserInput(request("idOrder"),10)
	if not validNum(pIdOrder) then
		response.redirect "msg.asp?message=35"
	end if

' extract real idorder (without prefix)
pIdOrder=(int(pIdOrder)-scpre)

dim pidCustomer, porderDate, pfirstname, plastname, pcustomerCompany, pphone, paddress, pzip, pstate, pcity, pcountryCode, pcomments, pshippingAddress, pshippingState, pshippingCity, pshippingCountryCode, pshippingZip, paddress2, pshippingFullName, pshippingCompany, pshippingAddress2, pshippingPhone, pOrderStatus

query="SELECT orders.idCustomer, orders.pcOrd_PaymentStatus,orders.orderstatus,orders.orderDate, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.customerType, orders.address, orders.zip, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.comments, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.pcOrd_shippingPhone, orders.shippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.idOrder, orders.rmaCredit, orders.ordPackageNum, orders.OrdShipType, orders.ord_DeliveryDate, orders.ord_OrderName, orders.ord_VAT,orders.pcOrd_CatDiscounts, orders.paymentDetails, orders.gwAuthCode, orders.gwTransId, orders.paymentCode FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE orders.idOrder="&pIdOrder&" AND ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rs.eof then
	query="SELECT orderstatus FROM orders WHERE idOrder="&pIdOrder&";"
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	if not rstemp.eof then
		pOrderStatus=rstemp("orderstatus")
		else
		pOrderStatus=""
	end if
	set rstemp=nothing
	set rs=nothing
	call closeDb()
	if pOrderStatus="2" then
		Session("message") = dictLanguage.Item(Session("language")&"_CustviewPastD_20")
		response.redirect "msgb.asp"
		else
 		response.redirect "msg.asp?message=35"
	end if
end if

pidCustomer=rs("idCustomer")

If statusAPP="1" Then
	query="SELECT Distinct ProductsOrdered.pcDropShipper_ID FROM pcDropShippersSuppliers,Products,ProductsOrdered WHERE ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & " AND products.idproduct=ProductsOrdered.idproduct AND ((pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct) OR (pcDropShippersSuppliers.idproduct=products.pcprod_ParentPrd)) AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & " AND ProductsOrdered.idorder=" & pIdOrder & ";"
Else
	query="SELECT ProductsOrdered.pcDropShipper_ID FROM pcDropShippersSuppliers INNER JOIN ProductsOrdered ON (pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & ") WHERE ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & " AND ProductsOrdered.idorder=" & pIdOrder & ";"
End If

set rsQ=connTemp.execute(query)
if rsQ.eof then
	set rsQ=nothing
	call closeDb()
	response.redirect "msg.asp?message=11"    
end if
set rsQ=nothing

if session("pc_sdsIsDropShipper")="1" then
	query="SELECT pcSupplier_NoticeType As A FROM pcSuppliers WHERE pcSupplier_ID=" & session("pc_idsds") & ";"
else
	query="SELECT pcDropShipper_NoticeType As A FROM pcDropShippers WHERE pcDropShipper_ID=" & session("pc_idsds") & ";"
end if
Set rsQ=connTemp.execute(query)
pcv_NoticeType=0
if not rsQ.eof then
	pcv_NoticeType=rsQ("A")
	if IsNull(pcv_NoticeType) or pcv_NoticeType="" then
		pcv_NoticeType=0
	end if
end if
set rsQ=nothing


pcv_PaymentStatus=rs("pcOrd_PaymentStatus")
if IsNull(pcv_PaymentStatus) or pcv_PaymentStatus="" then
	pcv_PaymentStatus=0
end if

pOrderStatus=rs("orderstatus")
if IsNull(pOrderStatus) or pOrderStatus="" then
	pOrderStatus=0
end if


porderDate=rs("orderDate")
porderDate=showdateFrmt(porderDate)
pfirstname=rs("name")
plastName=rs("lastName")
pcustomerCompany=rs("customerCompany")
pphone=rs("phone")
pcustomerType=rs("customerType")
paddress=rs("address")
pzip=rs("zip")
pstate=rs("stateCode")
if pstate="" then
	pstate=rs("state")
end if
pcity=rs("city")
pcountryCode=rs("countryCode")
pcomments=rs("comments")
pshippingAddress=rs("shippingAddress")
pshippingState=rs("shippingStateCode")
if pshippingState="" then
	pshippingState=rs("shippingState")
end if
pshippingCity=rs("shippingCity")
pshippingCountryCode=rs("shippingCountryCode")
pshippingZip=rs("shippingZip")
pshippingPhone=rs("pcOrd_shippingPhone")
pshippingFullName=rs("shippingFullName")
paddress2=rs("address2")
pshippingCompany=rs("shippingCompany")
pshippingAddress2=rs("shippingAddress2")
pidOrder=rs("idOrder")
pRmaCredit=rs("rmaCredit")
pOrdPackageNum=rs("ordPackageNum")
pOrdShipType=rs("OrdShipType")
pord_DeliveryDate=rs("ord_DeliveryDate")
pord_OrderName=rs("ord_OrderName")
pord_VAT=rs("ord_VAT")
pcv_CatDiscounts=rs("pcOrd_CatDiscounts")
if isNULL(pcv_CatDiscounts) OR pcv_CatDiscounts="" then
	pcv_CatDiscounts="0"
end if
pcpaymentDetails=trim(rs("paymentDetails"))
pcgwAuthCode=rs("gwAuthCode")
pcgwTransId=rs("gwTransId")
pcpaymentCode=rs("paymentCode")

If statusAPP="1" Then
	query="SELECT Distinct ProductsOrdered.idProductOrdered, ProductsOrdered.idProduct, ProductsOrdered.pcPrdOrd_Shipped, ProductsOrdered.quantity, ProductsOrdered.unitPrice, ProductsOrdered.QDiscounts, ProductsOrdered.ItemsDiscounts"

Else
	query="SELECT ProductsOrdered.idProductOrdered, ProductsOrdered.idProduct, ProductsOrdered.pcPrdOrd_Shipped, ProductsOrdered.quantity, ProductsOrdered.unitPrice, ProductsOrdered.QDiscounts, ProductsOrdered.ItemsDiscounts  "
End If

'CONFIGURATOR ADDON-S
If scBTO=1 then
	query=query&", ProductsOrdered.idconfigSession"
End If
'CONFIGURATOR ADDON-E
If statusAPP="1" Then
	query=query&", products.description, products.sku, orders.total, orders.paymentDetails, orders.taxamount, orders.shipmentDetails, orders.discountDetails,orders.orderstatus,orders.processDate, orders.shipdate, orders.shipvia, orders.trackingNum, orders.returnDate, orders.returnReason, orders.iRewardPoints, orders.iRewardValue,  orders.iRewardPointsCustAccrued, orders.dps FROM ProductsOrdered, products, orders,pcDropShippersSuppliers  WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct  AND ((pcDropShippersSuppliers.idproduct=ProductsOrdered.idproduct) OR (pcDropShippersSuppliers.idproduct=products.pcprod_ParentPrd)) AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & " AND ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & " AND orders.idOrder=" &pIdOrder
Else
	query=query&", products.description, products.sku, orders.total, orders.paymentDetails, orders.taxamount, orders.shipmentDetails, orders.discountDetails, orders.orderstatus,orders.processDate, orders.shipdate, orders.shipvia, orders.trackingNum, orders.returnDate, orders.returnReason, orders.iRewardPoints, orders.iRewardValue, orders.iRewardPointsCustAccrued, orders.dps FROM ProductsOrdered, products, orders,pcDropShippersSuppliers WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND pcDropShippersSuppliers.idproduct=products.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & session("pc_sdsIsDropShipper") & " AND ProductsOrdered.pcDropShipper_ID=" & session("pc_idsds") & " AND orders.idOrder=" &pIdOrder
End IF
set rsOrdObj=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsOrdObj=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rsOrdObj.eof then
	set rsOrdObj=nothing
	call closeDb()
 	response.redirect "msg.asp?message=35"
end if

query="SELECT pcPrdOrd_Shipped FROM ProductsOrdered WHERE idOrder=" & pIdOrder & " AND pcPrdOrd_Shipped=1;"
set rsQ=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
pcv_HaveShipped=0
if not rsQ.eof then
	pcv_HaveShipped=1
end if
set rsQ=nothing

%>
<!--#include file="header_wrapper.asp"-->
<div id="pcMain">
	<div class="pcMainContent">   
    <h1><%= dictLanguage.Item(Session("language")&"_CustviewPast_4")%></h1>
    <div class="pcSectionTitle">
      <%= dictLanguage.Item(Session("language")&"_CustviewOrd_9")&(int(pIdOrder)+scpre) & " - " & dictLanguage.Item(Session("language")&"_CustviewPastD_14") & porderDate%>
    </div>
		
    <div class="pcTable">
    
    	<div class="pcTableRow">
      	<div id="pcOrderLinks">
      		<a href="sds_OrdInvoice.asp?id=<%=pIdOrder%>" target="_blank">
          	<img src="<%=pcf_getImagePath("images","document.gif")%>" alt="<%= dictLanguage.Item(Session("language")&"_CustviewOrd_33") %>"></a> <a href="sds_OrdInvoice.asp?id=<%=pIdOrder%>" target="_blank"><%= dictLanguage.Item(Session("language")&"_CustviewOrd_33")%></a><%if pOrderStatus="3" or pOrderStatus="7" or pOrderStatus="8" then%> - <a href="sds_ShipOrderWizard1.asp?idOrder=<%=pIdOrder%>"><%= dictLanguage.Item(Session("language")&"_sds_viewpast_1c")%></a><%end if%>
  			</div>
        <div id="pcOrderBack">
      		<a class="pcButton pcButtonBack" href="sds_ViewPast.asp">
          	<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
            <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
          </a>
        </div>
    	</div>
		<% 
      ' START order delivery date, if any
      if (pord_DeliveryDate<>"") then
        if scDateFrmt="DD/MM/YY" then
          pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 4)
        else
          pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 3)
        end if
      pord_DeliveryDate = showdateFrmt(pord_DeliveryDate)
      
        if not scOrderName="1" Then 'Add <hr> only if the Order Name section is not shown %>
        <hr>
      <% end if %>
      
        <%=dictLanguage.Item(Session("language")&"_CustviewOrd_39")%><%=pord_DeliveryDate%> <% If pord_DeliveryTime <> "00:00" Then %><%=", " & pord_DeliveryTime%><% End If %>
  
        <hr>
      <%
      end if
      ' END order delivery date
      '
      ' START Billing and Shipping Addresses
      %>
	
  		<div class="pcSpacer"></div>
      
  		<div class="pcTableHeader">
        <div class="pcCustViewTableLabel">&nbsp;</div>
        <div class="pcCustViewTableField"><strong><%= dictLanguage.Item(Session("language")&"_orderverify_23")%></strong></div>
        <div class="pcCustViewTableField"><strong><%= dictLanguage.Item(Session("language")&"_orderverify_24")%></strong></div>
      </div>
      
      <% 'Billing/Shipping Name %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel"><%= replace(dictLanguage.Item(Session("language")&"_orderverify_7"),"''","'") %></div>
        <div class="pcCustViewTableField"><%= pFirstName&" "&plastname %></div>
        <div class="pcCustViewTableField"><%= pshippingFullName %></div>
      </div>
      
      <% 'Billing/Shipping Company %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel"><%= dictLanguage.Item(Session("language")&"_orderverify_8") %></div>
        <div class="pcCustViewTableField"><%= pcustomerCompany %></div>
        <div class="pcCustViewTableField">
					<% 
						if pshippingCompany<>"" then
							response.write pshippingCompany
						end if 
					%>
        </div>
      </div>
      
      <% 'Billing/Shipping Phone %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel"><%= dictLanguage.Item(Session("language")&"_orderverify_9") %></div>
        <div class="pcCustViewTableField"><%= pPhone %></div>
        <div class="pcCustViewTableField"><%= pshippingPhone %></div>
      </div>
      
      <% 'Billing/Shipping Address %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel"><%= dictLanguage.Item(Session("language")&"_orderverify_10") %></div>
        <div class="pcCustViewTableField"><%= paddress %></div>
        <div class="pcCustViewTableField">
        	<%					
						if pshippingAddress="" then
							response.write "Same as Billing Address"
						else
							response.write pshippingAddress
						end if
					%>
        </div>
      </div>
      
      <% 'Billing/Shipping Address 2 %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel">&nbsp;</div>
        <div class="pcCustViewTableField"><%= paddress2 %></div>
        <div class="pcCustViewTableField">
					<%
            if pshippingAddress2<>"" then
              response.write pshippingAddress2
            end if
          %>
        </div>
      </div>
      
      <% 'Billing/Shipping City, State, & Zip %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel">&nbsp;</div>
        <div class="pcCustViewTableField"><%= pCity&", "&pState&" "&pzip %></div>
        <div class="pcCustViewTableField">
					<%
						if pshippingAddress<>"" then
							response.write pShippingCity&", "&pshippingState
							If pshippingState="" then
								response.write pshippingStateCode
							End If
							response.write " "&pshippingZip
						end if
          %>
        </div>
      </div>
      
      <% 'Billing/Shipping Country %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel">&nbsp;</div>
        <div class="pcCustViewTableField"><%= pCountryCode %></div>
        <div class="pcCustViewTableField">
					<%
						if scAlwAltShipAddress="-1" then
							strFedExCountryCode=pshippingCountryCode
						else
							strFedExCountryCode=pCountryCode
						end if 
						response.write pshippingCountryCode
          %>
        </div>
      </div>
      
      <div class="pcSpacer"></div>
      
    	<div class="pcTableRowFull">
			<%if pcv_NoticeType="1" then%>
        <div class="pcErrorMessage"><%= dictLanguage.Item(Session("language")&"_sds_custviewpastD_7")%><br>
        <%=scCompanyName%><br>
        <%=scCompanyAddress%><br>
        <%=scCompanyCity & ", " & scCompanyState & " - " & scCompanyZip%><br>
        <%=scCompanyCountry%><br>
        <a href="mailto:<%=scFrmEmail%>"><%=scFrmEmail%></a></div>
      <%else%>
        <div class="pcErrorMessage"><%= dictLanguage.Item(Session("language")&"_sds_custviewpastD_8")%></div>
      <%end if%>
      </div>
      
		<% 
		' END Billing and Shipping Addresses
	
		' START Order Details
		%>
  
  	<div class="pcTableHeader">
    	<div class="pcSdsViewPastD_Qty"><%= dictLanguage.Item(Session("language")&"_orderverify_25") %></div>
    	<div class="pcSdsViewPastD_Sku"><%= dictLanguage.Item(Session("language")&"_orderverify_26") %></div>
    	<div class="pcSdsViewPastD_Desc"><%= dictLanguage.Item(Session("language")&"_orderverify_27") %></div>
     	<div class="pcSdsViewPastD_Status">
      	<% if pcv_HaveShipped=1 then %>
			  	<%= dictLanguage.Item(Session("language")&"_sds_custviewpastD_1") %>
				<% end if%>
      </div>
    </div>
    
	
		<% dim pidProduct, pquantity, punitPrice, pxfdetails, pidconfigSession, pdescription, pSku, pcDPs, ptotal, ppaymentDetails,ptaxamount,pshipmentDetails, pdiscountDetails
    dim pprocessDate, pshipdate, pshipvia, ptrackingNum, preturnDate, preturnReason, piRewardPoints, piRewardValue, piRewardPointsCustAccrued,ptaxdetails, pOpPrices, rsObjOptions, pRowPrice, count, rsConfigObj,stringProducts, stringValues, stringCategories, ArrProduct, ArrValue, ArrCategory,i, s,OptPrice,xfdetails, xfarray, q
    
    Dim pcv_strSelectedOptions, pcv_strOptionsPriceArray, pcv_strOptionsArray
    Dim pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice
    Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions
    
    do while not rsOrdObj.eof
    
    
      tint_idProductOrdered = rsOrdObj("idProductOrdered")
      pidProduct=rsOrdObj("idProduct")
      pcv_Shipped=rsOrdObj("pcPrdOrd_Shipped")
      if IsNull(pcv_Shipped) or pcv_Shipped="" then
        pcv_Shipped=1
      end if
      pquantity=rsOrdObj("quantity")
    
      punitPrice=rsOrdObj("unitPrice")
      QDiscounts=rsOrdObj("QDiscounts")
      ItemsDiscounts=rsOrdObj("ItemsDiscounts")
      'CONFIGURATOR ADDON-S
      if scBTO=1 then
        pidconfigSession=rsOrdObj("idconfigSession")
        if pidconfigSession="" then
          pidconfigSession="0"
        end if
      End If
      'CONFIGURATOR ADDON-E
      pdescription=rsOrdObj("description")
      pSku=rsOrdObj("sku")
      ptotal=rsOrdObj("total")
      ppaymentDetails=trim(rsOrdObj("paymentDetails"))
      ptaxamount=rsOrdObj("taxamount")
      pshipmentDetails=rsOrdObj("shipmentDetails")
      pdiscountDetails=rsOrdObj("discountDetails")
      porderstatus=rsOrdObj("orderstatus")
      pprocessDate=rsOrdObj("processDate")
      pprocessDate=ShowDateFrmt(pprocessDate)
      pshipdate=rsOrdObj("shipdate")
      pshipdate=ShowDateFrmt(pshipdate)
      pshipvia=rsOrdObj("shipvia")
      ptrackingNum=rsOrdObj("trackingNum")
      preturnDate=rsOrdObj("returnDate")
      preturnDate=ShowDateFrmt(preturnDate)
      preturnReason=rsOrdObj("returnReason")
      piRewardPoints=rsOrdObj("iRewardPoints")
      piRewardValue=rsOrdObj("iRewardValue")
      piRewardPointsCustAccrued=rsOrdObj("iRewardPointsCustAccrued")
      pcDPs=rsOrdObj("DPs")
      
      pIdConfigSession=trim(pidconfigSession)
      
      pOpPrices=0
      query = "SELECT ProductsOrdered.xfdetails, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray FROM ProductsOrdered WHERE idProductOrdered = " & tint_idProductOrdered & ";"
      set rsTObj1=server.CreateObject("ADODB.RecordSet")
      set rsTObj1=conntemp.execute(query)
      
      pxfdetails=rsTObj1("xfdetails")
      '// Product Options Arrays
      pcv_strSelectedOptions = rsTObj1("pcPrdOrd_SelectedOptions") ' Column 11
      pcv_strOptionsPriceArray = rsTObj1("pcPrdOrd_OptionsPriceArray") ' Column 25
      pcv_strOptionsArray = rsTObj1("pcPrdOrd_OptionsArray") ' Column 4
      
      set rsTObj1 = nothing
      
      query = "SELECT  orders.taxdetails FROM Orders WHERE idOrder = " & pIdOrder & ";"
      set rsTObj1=server.CreateObject("ADODB.RecordSet")
      set rsTObj1=conntemp.execute(query)
      
      ptaxdetails=rsTObj1("taxdetails")

      set rsTObj1 = nothing
      '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      ' START: Get the total Price of all options
      '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      pOpPrices=0
      dim pcv_tmpOptionLoopCounter, pcArray_TmpCounter
      
      If len(pcv_strOptionsPriceArray)>0 then
      
        pcArray_TmpCounter = split(pcv_strOptionsPriceArray,chr(124))
        For pcv_tmpOptionLoopCounter = 0 to ubound(pcArray_TmpCounter)
          pOpPrices = pOpPrices + pcArray_TmpCounter(pcv_tmpOptionLoopCounter)
        Next
        
      end if				

      if NOT isNumeric(pOpPrices) then
        pOpPrices=0
      end if	
      
      '// Apply Discounts to Options Total
      '   >>> call function "pcf_DiscountedOptions(OriginalOptionsTotal, ProductID, Quantity, CustomerType)" from stringfunctions.asp
      Dim pcv_intDiscountPerUnit
      pOpPrices = pcf_DiscountedOptions(pOpPrices, pidProduct, pquantity, pcustomerType)
      
      '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      ' END: Get the total Price of all options
      '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
      %>
					
			<% 'CONFIGURATOR ADDON-S
					err.number=0
					TotalUnit=0
					If scBTO=1 then
						pIdConfigSession=trim(pidconfigSession)
						if pIdConfigSession<>"0" then 
							query="SELECT stringProducts, stringValues, stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
							set rsConfigObj=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsConfigObj=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							stringProducts=rsConfigObj("stringProducts")
							stringValues=rsConfigObj("stringValues")
							stringCategories=rsConfigObj("stringCategories")
							stringQuantity=rsConfigObj("stringQuantity")
							stringPrice=rsConfigObj("stringPrice")
							ArrProduct=Split(stringProducts, ",")
							ArrValue=Split(stringValues, ",")
							ArrCategory=Split(stringCategories, ",")
							ArrQuantity=Split(stringQuantity, ",")
							ArrPrice=Split(stringPrice, ",")
							set rsConfigObj=nothing
							for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
							query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
							set rsConfigObj=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsConfigObj=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
				if NOT validNum(ArrQuantity(i)) then
								pIntQty=1
							else
								pIntQty=ArrQuantity(i)
							end if
							if (CDbl(ArrValue(i))<>0) or (((ArrQuantity(i)-1)*pQuantity>0) and (ArrPrice(i)>0)) then
								if (ArrQuantity(i)-1)>=0 then
									UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
								else
									UPrice=0
								end if
								TotalUnit=TotalUnit+((ArrValue(i)+UPrice)*pQuantity)
							end if
							set rsConfigObj=nothing
							next
						end if 
					End If 
		'CONFIGURATOR ADDON-E


		
		if TotalUnit>0 then
			punitPrice1=punitPrice
			if pIdConfigSession<>"0" then
				pRowPrice1=Cdbl(pquantity * ( punitPrice1 )) - TotalUnit
				punitPrice1=Round(pRowPrice1/pquantity,2)
			else
				pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
			end if
		else
			punitPrice1=punitPrice
			if pIdConfigSession<>"0" then
				pRowPrice1=Cdbl(pquantity * ( punitPrice1 ))
			else
				pRowPrice1=Cdbl(pquantity * ( punitPrice1 - pOpPrices) )
				punitPrice1=Round(pRowPrice1/pquantity,2)
			end if
		end if
		%>
		
    <div class="pcTableRow"> 
      <div class="pcSdsViewPastD_Qty"><%=pquantity%></div>
      <div class="pcSdsViewPastD_Sku"><%=pSku%></div>
      <div class="pcSdsViewPastD_Desc"><%=pdescription%></div>
      
      <div class="pcSdsViewPastD_Status">
        <%if pcv_HaveShipped=1 then
          if pcv_Shipped="1" then
            response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_3")
          else
            response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_2")
          end if
        end if%>
      </div>
    </div>
		
		<% 'CONFIGURATOR ADDON-S
    err.number=0
    If scBTO=1 then
      pIdConfigSession=trim(pidconfigSession)
      if pIdConfigSession<>"0" then 
        query="SELECT stringProducts, stringValues, stringCategories,stringQuantity,stringPrice FROM configSessions WHERE idconfigSession=" & pIdConfigSession
        set rsConfigObj=conntemp.execute(query)
        if err.number<>0 then
          call LogErrorToDatabase()
          set rsConfigObj=nothing
          call closedb()
          response.redirect "techErr.asp?err="&pcStrCustRefID
        end if
        stringProducts=rsConfigObj("stringProducts")
        stringValues=rsConfigObj("stringValues")
        stringCategories=rsConfigObj("stringCategories")
        stringQuantity=rsConfigObj("stringQuantity")
        stringPrice=rsConfigObj("stringPrice")
        ArrProduct=Split(stringProducts, ",")
        ArrValue=Split(stringValues, ",")
        ArrCategory=Split(stringCategories, ",")
        ArrQuantity=Split(stringQuantity, ",")
        ArrPrice=Split(stringPrice, ",")
        set rsConfigObj=nothing
        %>
  
        <tr> 
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td colspan="3"> 
            <table class="pcShowBTOconfiguration">
              <tr> 
                <td colspan="2">  
                  <p><%= bto_dictLanguage.Item(Session("language")&"_CustviewPastD_1")%></p>
                </td>
                <td>&nbsp;</td>
              </tr>
              <% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
                query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
                set rsConfigObj=conntemp.execute(query)
                if err.number<>0 then
                  call LogErrorToDatabase()
                  set rsConfigObj=nothing
                  call closedb()
                  response.redirect "techErr.asp?err="&pcStrCustRefID
                end if
                if NOT validNum(ArrQuantity(i)) then
                  pIntQty=1
                else
                  pIntQty=ArrQuantity(i)
                end if
                strCategoryDesc=rsConfigObj("categoryDesc")
                strDescription=rsConfigObj("description") %>
                <tr> 
                  <td width="85%" valign="top" colspan="2"> 
                    <p><%=strCategoryDesc%>:	<%=strDescription%><%if pIntQty>1 then%> - QTY: <%=ArrQuantity(i)%><%end if%></p>
                    </td>
                    <td width="15%" valign="top" nowrap align="right">
                    </td>
                  
                </tr>
            <% set rsConfigObj=nothing
            next %>
            </table>
          </td>
          <td>&nbsp;</td>
        </tr>
      <% end if 
    End If 
    'CONFIGURATOR ADDON-E
    %>
					
    <!-- start options -->
    <%
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' START: SHOW PRODUCT OPTIONS
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    if (len(pcv_strSelectedOptions)>0) AND (pcv_strSelectedOptions<>"NULL") then
		
			'#####################
			' START LOOP
			'#####################	
			
			'// Generate Our Local Arrays from our Stored Arrays  
			
			' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers	
			pcArray_strSelectedOptions = ""					
			pcArray_strSelectedOptions = Split(pcv_strSelectedOptions,chr(124))
			
			' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
			pcArray_strOptionsPrice = ""
			pcArray_strOptionsPrice = Split(pcv_strOptionsPriceArray,chr(124))
			
			' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
			pcArray_strOptions = ""
			pcArray_strOptions = Split(pcv_strOptionsArray,chr(124))
			
			' Get Our Loop Size
			pcv_intOptionLoopSize = 0
			pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)
			
			' Start in Position One
			pcv_intOptionLoopCounter = 0
			
			' Display Our Options
			For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize
    	%>
        <div class="pcTableRow">
          <div class="pcSdsViewPastD_Qty">&nbsp;</div>
          <div class="pcSdsViewPastD_Sku">&nbsp;</div>
          <div class="pcSdsViewPastD_Desc">
						<%=pcArray_strOptions(pcv_intOptionLoopCounter) %>
        									
						<% 
            tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
            
            if tempPrice="" or tempPrice=0 then
            else
              '// Adjust for Quantity Discounts
              tempPrice = tempPrice - ((pcv_intDiscountPerUnit/100) * tempPrice)
            end if 
            %>		
          </div>	
        <%
        Next
        '#####################
        ' END LOOP
        '#####################					
        %>
      </div>													
		<%					
    end if
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' END: SHOW PRODUCT OPTIONS
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    %>
    <!-- end options -->
					
					
					
			<% 'CONFIGURATOR ADDON-S
      err.number=0
      If scBTO=1 then
        pIdConfigSession=trim(pidconfigSession)
        if pIdConfigSession<>"0" then
          MyTest=0
          ItemsDiscounts=trim(ItemsDiscounts)
          if (ItemsDiscounts<>"") and (CDbl(ItemsDiscounts)<>"0") then
            MyTest=1%>
            <tr valign="top"> 
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <% pRowPrice=pRowPrice-Cdbl(ItemsDiscounts)
          end if%>
          <% 'BTO Additional Charges
          If scBTO=1 then
            pIdConfigSession=trim(pidconfigSession)
            if pIdConfigSession<>"0" then 
              query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & pIdConfigSession
              set rsConfigObj=conntemp.execute(query)
              if err.number<>0 then
                call LogErrorToDatabase()
                set rsConfigObj=nothing
                call closedb()
                response.redirect "techErr.asp?err="&pcStrCustRefID
              end if
              stringCProducts=rsConfigObj("stringCProducts")
              stringCValues=rsConfigObj("stringCValues")
              stringCCategories=rsConfigObj("stringCCategories")
              ArrCProduct=Split(stringCProducts, ",")
              ArrCValue=Split(stringCValues, ",")
              ArrCCategory=Split(stringCCategories, ",")
              set rsConfigObj=nothing
              if ArrCProduct(0)<>"na" then%>
                <tr> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td colspan="3"> 
                    <table class="pcShowBTOconfiguration">
                      <tr> 
                        <td colspan="2">
                        <p><%= bto_dictLanguage.Item(Session("language")&"_CustviewPastD_5")%></p> 
                        </td>
                      </tr>
                      <% Charges=0
                      for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
                        query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
                        set rsConfigObj=conntemp.execute(query)
                        if err.number<>0 then
                          call LogErrorToDatabase()
                          set rsConfigObj=nothing
                          call closedb()
                          response.redirect "techErr.asp?err="&pcStrCustRefID
                        end if
                        strCategoryDesc=rsConfigObj("categoryDesc")
                        strDescription=rsConfigObj("description")
                        if (CDbl(ArrCValue(i))>0)then
                          Charges=Charges+cdbl(ArrCValue(i))
                        end if %>
                        <tr> 
                          <td width="85%" valign="top"><p><%=strCategoryDesc%>:	<%=strDescription%></p></td>
                          <td width="15%" valign="top" nowrap align="right"></td>
                        </tr>
                        <% set rsConfigObj=nothing
                      next %>
                    </table>
                  </td>
                  <td>&nbsp;</td>
                </tr>
          
                <% pRowPrice=pRowPrice+Cdbl(Charges)
              end if 'Have Charges
            end if 
          End If 
          'BTO Additional Charges
          
          QDiscounts=trim(QDiscounts)
          if (QDiscounts<>"") and (CDbl(QDiscounts)<>"0") then
            MyTest=1%>
            <tr valign="top"> 
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td></td>
              <td>&nbsp;</td>
              <td></td>
              <td>&nbsp;</td>
            </tr>
            <% pRowPrice=pRowPrice-Cdbl(QDiscounts)
          end if%>
          
          <%if MyTest=1 then%>
            <tr valign="top"> 
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td></td>
              <td></td>
              <td>&nbsp;</td>
            </tr>
          <%end if
        end if
      end if 'BTO%>
				
			<% 'show xtra options
      '-----------------
      xfdetails=pxfdetails
      If len(xfdetails)>3 then
        xfarray=split(xfdetails,"|")
        for q=lbound(xfarray) to ubound(xfarray) %>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td><p><%=xfarray(q)%></p></td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
        <% next
      End If 
      '----------------- %>
      <% count=count+1
      If pshippingAddress="" then
        'grab shipping address from shipping...
        pshippingAddress=pAddress
        pshippingAddress2=pAddress2
        pshippingCity=pCity
        pshippingState=pState
        pshippingZip=pZip
        pshippingCountryCode=pCountryCode
      End if
      rsOrdObj.movenext  
    loop%>
	<%' END Order Details
	
	' START Other order information
	%>
  	
		<div class="pcSpacer"></div>
			
    <!-- if order was cancelled -->
    <% if pOrderStatus="5" then %>
      <div class="pcTableRowFull">
      	<%= dictLanguage.Item(Session("language")&"_CustviewOrd_18")%>
      </div>
    <% else %>
      
    <!-- if order was returned -->
    <% if pOrderStatus="6" then %>
      <div class="pcTableRowFull">
      	<%= dictLanguage.Item(Session("language")&"_CustviewOrd_26")%>
      </div>
      <div class="pcTableRowFull">
      	<hr>
      </div>
    <% end if %>
    <!-- end order returned -->
  
    <!-- order has been processed, show date -->
    <% if int(pOrderStatus)>2 then %>
    	<div class="pcTableRowFull">
				<%= dictLanguage.Item(Session("language")&"_CustviewPastD_22b")%>
      </div>
    	<div class="pcTableRowFull">
        <%= dictLanguage.Item(Session("language")&"_CustviewPastD_22") & pprocessDate %>
      </div>
    <% else %>
    <!-- else if order has not been processed, tell drop-shipper -->
    	<div class="pcTableRowFull">
      	<%= dictLanguage.Item(Session("language")&"_CustviewPastD_20")%>
      </div>
    <% end if %>
    <!-- end order processed check -->
    <%end if%>
	<%
	' START Shipment type
	%>
    <div class="pcTableRowFull">
  		<hr>
    </div>
    
    <div class="pcTableRowFull">
      <strong><%=dictLanguage.Item(Session("language")&"_sds_custviewpastD_15")%></strong>
    </div>
      
    <div class="pcTableRowFull">
      <%= dictLanguage.Item(Session("language")&"_sds_custviewpastD_16")%>
      <% 
      Shipper=""
      Service=""
      Postage=""
      serviceHandlingFee=""
      If pSRF="1" then
        response.write ship_dictLanguage.Item(Session("language")&"_noShip_b")
      else
        'get shipping details...
        shipping=split(pshipmentDetails,",")
        if ubound(shipping)>1 then
          if NOT isNumeric(trim(shipping(2))) then
            varShip="0"
            response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
          else
            Shipper=shipping(0)
            Service=shipping(1)
            Postage=trim(shipping(2))
          end if
          if len(Service)>0 then
            response.write Service
          End If
        else
          varShip="0"
          response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
        end if
      end if
      %>
    </div>
	<%
		if pOrdShipType=0 then
			pDisShipType=dictLanguage.Item(Session("language")&"_sds_custviewpastD_18")
		else
			pDisShipType=dictLanguage.Item(Session("language")&"_sds_custviewpastD_19")
		end if
		if varShip<>"0" then
	%>
    <div class="pcTableRowFull">
			<%=dictLanguage.Item(Session("language")&"_sds_custviewpastD_17")%><%=pDisShipType%>
    </div>
	<%
	end if
	' END Shipment Type
	%>
  
  </div>
  
  <div class="pcSpacer"></div>

    <div class="pcFormButtons">       
      <a class="pcButton pcButtonBack" href="sds_ViewPast.asp">
        <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
        <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
      </a>
  
    </div>
	</div>
</div>
</div>
<!--#include file="footer_wrapper.asp"-->
