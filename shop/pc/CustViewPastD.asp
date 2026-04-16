<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%'Allow Guest Account
AllowGuestAccess=1
%>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<%pcStrPageName="CustViewPastD.asp"%>
<%
err.number=0
dim pIdOrder
%>
<!--#include file="prv_getsettings.asp"-->
<%
pcv_RWActive=pcv_Active
pIdOrder=getUserInput(request("idOrder"),10)
if not validNum(pIdOrder) then response.Redirect "custPref.asp"

' extract real idorder (without prefix)
pIdOrder=(int(pIdOrder)-scpre)
session("idOrderConfirm") = pIdOrder

Dim pord_DeliveryDate, pord_OrderName

if request("action")="rename" then
  pord_OrderName=getUserInput(request("ord_OrderName"),0)
  if pord_OrderName = "" then
    pord_OrderName = "No Name"
  end if
  query="update orders set ord_OrderName=N'" & pord_OrderName & "' where idOrder=" & pidOrder
  set rs=server.CreateObject("ADODB.RecordSet")
  set rs=connTemp.execute(query)
  if err.number<>0 then
    call LogErrorToDatabase()
    set rs=nothing
    call closedb()
    response.redirect "techErr.asp?err="&pcStrCustRefID
  end if
end if

tmpReSent=0

if request("action")="resend" then
  GC_ReName=getUserInput(request("GC_RecName"),0)
  GC_ReEmail=getUserInput(request("GC_RecEmail"),0)
  GC_ReMsg=getUserInput(request("GC_RecMsg"),0)
  
  query="UPDATE orders SET pcOrd_GcReName=N'" & GC_ReName & "',pcOrd_GcReEmail='" & GC_ReEmail & "',pcOrd_GcReMsg=N'" & GC_ReMsg & "' WHERE idOrder="& pIdOrder
  Set rs=Server.CreateObject("ADODB.Recordset")
  Set rs=conntemp.execute(query)
  Set rs=nothing
  
  ReciEmail=""

  query="select idproduct from ProductsOrdered WHERE idOrder="& pIdOrder
  pidorder=pIdOrder
  set rs11=connTemp.execute(query)
  do while not rs11.eof
    query="select products.Description,pcGCOrdered.pcGO_GcCode,pcGc.pcGc_EOnly from Products,pcGc,pcGCOrdered where products.idproduct=" & rs11("idproduct") & " and pcGC.pcGc_IDProduct=products.idproduct and pcGCOrdered.pcGO_idproduct=Products.idproduct and products.pcprod_GC=1 and pcGCOrdered.pcGO_idOrder="& pIdOrder
    set rs=connTemp.execute(query)
  
    if not rs.eof then
      pIdproduct=rs11("idproduct")
      pName=rs("Description")
      pCode=rs("pcGO_GcCode")
      pEOnly=rs("pcGc_EOnly")
      
        query="select pcGO_Amount,pcGO_GcCode,pcGO_ExpDate from pcGCOrdered where pcGO_idproduct=" & rs11("idproduct") & " and pcGO_idorder=" & pidorder
        set rs19=connTemp.execute(query)
        
        do while not rs19.eof
        pAmount=rs19("pcGO_Amount")
        if pAmount<>"" then
        else
        pAmount="0"
        end if
        
        ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_68") & scCurSign & money(pAmount) & vbcrlf
        
        ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_69") & rs19("pcGO_GcCode") & vbcrlf
        pExpDate=rs19("pcGO_ExpDate")
         
        if year(pExpDate)="1900" then
        ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_45b") & vbcrlf
        else
        if scDateFrmt="DD/MM/YY" then
        pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
        else
        pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
        end if
        ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_70") & pExpDate & vbcrlf
        end if
        if pEOnly="1" then
        ReciEmail=ReciEmail & dictLanguage.Item(Session("language")&"_sendMail_71") & vbcrlf
        end if
        ReciEmail=ReciEmail & vbcrlf
        rs19.movenext
        loop

    end if
  rs11.MoveNext
  loop
  set rs11=nothing
  
  query="SELECT customers.name,customers.lastname,customers.email,orders.pcOrd_GcReName,orders.pcOrd_GcReEmail,orders.pcOrd_GcReMsg FROM customers INNER JOIN Orders ON customers.idcustomer=orders.idcustomer WHERE idOrder="& pIdOrder
  set rs11=connTemp.execute(query)

  if not rs11.eof then
    pCustomerFullName=rs11("name") & " " & rs11("lastname")
    pCustomerFullNamePlusEmail=pCustomerFullName & " (" & rs11("email") & ")"
    GcReName=rs11("pcOrd_GcReName")
    GcReEmail=rs11("pcOrd_GcReEmail")
    GcReMsg=rs11("pcOrd_GcReMsg")
  
    if GcReEmail<>"" then
      if GcReName<>"" then
      else
        GcReName=GcReEmail
      end if
      ReciEmail1=replace(dictLanguage.Item(Session("language")&"_sendMail_66"),"<recipient name>",GcReName)
      ReciEmail2=replace(dictLanguage.Item(Session("language")&"_sendMail_67"),"<customer name>",pCustomerFullNamePlusEmail)
      if GcReMsg<>"" then
        ReciEmail3=replace(dictLanguage.Item(Session("language")&"_sendMail_72"),"<customer name>",pCustomerFullNamePlusEmail) & vbcrlf & GcReMsg & vbcrlf
      else
        ReciEmail3=""
      end if
      ReciEmail=ReciEmail1 & vbcrlf & vbcrlf & ReciEmail2 & vbcrlf & vbcrlf & ReciEmail & ReciEmail3
      ReciEmail=ReciEmail & vbcrlf & scCompanyName & vbCrLf & scStoreURL & vbcrlf & vbCrLf
      call sendmail (scCompanyName, scEmail, GcReEmail,pCustomerFullName & dictLanguage.Item(Session("language")&"_sendMail_73"), replace(ReciEmail, "&quot;", chr(34)))
	  call pcs_hookGCOrderEmailSent(GcReEmail)
      tmpReSent=1
    end if
  end if
  set rs11=nothing
end if


query="SELECT orders.pcOrd_OrderKey,customers.email,customers.fax,orders.pcOrd_ShippingEmail,orders.pcOrd_ShippingFax,orders.pcOrd_ShowShipAddr,orders.idCustomer, orders.pcOrd_PaymentStatus,orders.orderDate, customers.name, customers.lastName, customers.customerCompany, customers.phone, customers.customerType, orders.address, orders.zip, orders.stateCode, orders.state, orders.city, orders.countryCode, orders.comments, orders.shippingAddress, orders.shippingStateCode, orders.shippingState, orders.shippingCity, orders.shippingCountryCode, orders.shippingZip, orders.pcOrd_shippingPhone, orders.shippingFullName, orders.address2, orders.shippingCompany, orders.shippingAddress2, orders.idOrder, orders.rmaCredit, orders.ordPackageNum, orders.ord_DeliveryDate, orders.ord_OrderName, orders.ord_VAT,orders.pcOrd_CatDiscounts, orders.paymentDetails, orders.gwAuthCode, orders.gwTransId, orders.paymentCode FROM customers INNER JOIN orders ON customers.idcustomer = orders.idCustomer WHERE (((orders.idOrder)="&pIdOrder&"));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
  call LogErrorToDatabase()
  set rs=nothing
  call closedb()
  response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rs.eof then
  set rs=nothing
  call closeDb()
  response.redirect "msg.asp?message=35"     
end if 

dim pidCustomer, porderDate, pfirstname, plastname,pcustomerCompany, pphone, paddress, pzip, pstate, pcity, pcountryCode, pcomments, pshippingAddress, pshippingState, pshippingCity, pshippingCountryCode, pshippingZip, paddress2, pshippingFullName, pshippingCompany, pshippingAddress2, pshippingPhone

pidCustomer=rs("idCustomer")
if int(Session("idcustomer"))<=0 then
  if session("REGidCustomer")>"0" then
    testidCustomer=int(session("REGidCustomer"))
  end if
else
  testidCustomer=int(Session("idcustomer"))
end if
if testidCustomer<>int(pidCustomer) then
  set rs=nothing
  call closeDb()
  session("REGidCustomer")=""
  response.redirect "msg.asp?message=11"    
end if

'Start SDBA
pcv_PaymentStatus=rs("pcOrd_PaymentStatus")
if IsNull(pcv_PaymentStatus) or pcv_PaymentStatus="" then
  pcv_PaymentStatus=0
end if
'End SDBA

pcOrderKey=rs("pcOrd_OrderKey")
pEmail=rs("email")
pFax=rs("fax")
pshippingEmail=rs("pcOrd_ShippingEmail")
pshippingFax=rs("pcOrd_ShippingFax")
pcShowShipAddr=rs("pcOrd_ShowShipAddr")
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

  '// START - Test for existence of separate shipping address
  if IsNull(pcShowShipAddr) OR (pcShowShipAddr="") OR (pcShowShipAddr="0") then
    'This might be a v3 store, check another field
    if trim(pshippingAddress)="" then
      pcShowShipAddr=0
      else
      pcShowShipAddr=1
    end if
  end if
  '// END

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

query="SELECT Orders.pcOrd_GWTotal,Orders.pcOrd_IDEvent,ProductsOrdered.pcPO_GWOpt,ProductsOrdered.pcPO_GWNote,ProductsOrdered.pcPO_GWPrice,orders.pcOrd_GCs,orders.pcOrd_GcCode,orders.pcOrd_GcUsed,ProductsOrdered.idProduct, ProductsOrdered.pcPrdOrd_Shipped, ProductsOrdered.quantity, ProductsOrdered.unitPrice,ProductsOrdered.QDiscounts,ProductsOrdered.ItemsDiscounts, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.xfdetails  "
'CONFIGURATOR ADDON-S
if scBTO=1 then
  query=query&", ProductsOrdered.idconfigSession"
end if
'CONFIGURATOR ADDON-E
query=query&", products.description, products.sku, orders.total, orders.paymentDetails, orders.taxamount, orders.shipmentDetails, orders.discountDetails, orders.pcOrd_GCDetails, orders.orderstatus,orders.processDate, orders.shipdate, orders.shipvia, orders.trackingNum, orders.returnDate, orders.returnReason, orders.iRewardPoints, orders.iRewardValue, orders.iRewardPointsCustAccrued, orders.taxdetails,orders.dps,pcPrdOrd_BundledDisc FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idCustomer=" &pidCustomer& " AND orders.idOrder=" &pIdOrder
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

pdescription=rsOrdObj("description")
pSku=rsOrdObj("sku")
ptotal=rsOrdObj("total")
ppaymentDetails=trim(rsOrdObj("paymentDetails"))
ptaxamount=rsOrdObj("taxamount")
pshipmentDetails=rsOrdObj("shipmentDetails")
pdiscountDetails=rsOrdObj("discountDetails")
GCDetails=rsOrdObj("pcOrd_GCDetails")
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
ptaxdetails=rsOrdObj("taxdetails")
pcDPs=rsOrdObj("DPs")
pcPrdOrd_BundledDisc=rsOrdObj("pcPrdOrd_BundledDisc")
pIdConfigSession=trim(pidconfigSession)

'GGG Add-on start
pGWTotal=rsOrdObj("pcOrd_GWTotal")
if pGWTotal<>"" then
else
pGWTotal="0"
end if
gIDEvent=rsOrdObj("pcOrd_IDEvent")
if gIDEvent<>"" then
else
gIDEvent="0"
end if
''GGG Add-on end

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
<script type=text/javascript>
  function openbrowser(url) {
      self.name = "productPageWin";
      popUpWin = window.open(url,'rating','toolbar=0,location=0,directories=0,status=0,top=0,scrollbars=yes,resizable=1,width=705,height=535');
      if (navigator.appName == 'Netscape') {
      popUpWin.focus();
    }
  }
</script>
<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="Contact Us">Customer Service Area</h3>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
	<!-- /Header: pagetitle -->

	<section id="intWarranties" class="intWarranties paddingtop-30 paddingbot-70">	
           <div class="container">
				<div class="row">
                	<div class="col-sm-12 warrantyHeading wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s">
<div id="pcMain">
  <div class="pcMainContent">
    
    <h1><%= dictLanguage.Item(Session("language")&"_CustviewPast_4")%></h1>
    
    <div class="pcSectionTitle">
      <%= dictLanguage.Item(Session("language")&"_CustviewOrd_9")&(int(pIdOrder)+scpre) & " - " & dictLanguage.Item(Session("language")&"_CustviewPastD_14") & porderDate%>
      <%if pcOrderKey<>"" then%> - <%=dictLanguage.Item(Session("language")&"_opc_common_1")%>&nbsp;<%=pcOrderKey%><%end if%>
    </div>
    
    <%if tmpReSent=1 then%>
      <div class="pcSuccessMessage">
        <%= dictLanguage.Item(Session("language")&"_GCRecipient_3")%>
      </div>
    <%end if%>
    <%if session("REGidCustomer")>"0" then %>
      <div class="pcInfoMessage">
        <%= dictLanguage.Item(Session("language")&"_opc_85")%>
      </div>
    <%end if%>
    
    <div class="pcTable hidden-xs">
      <div class="pcTableRow">
        <div id="pcOrderLinks">
          <a href="custOrdInvoice.asp?id=<%=pIdOrder%>" target="_blank"><img src="<%=pcf_getImagePath("images","document.gif")%>" align="middle" style="margin: -8px -3px 0 0;"></a> 
          <a href="custOrdInvoice.asp?id=<%=pIdOrder%>" target="_blank"><%= dictLanguage.Item(Session("language")&"_CustviewOrd_33")%></a>
          <a href="custOrdInvoicePDF.asp?id=<%=pIdOrder%>" target="_blank"><img src="<%=pcf_getImagePath("images","document.gif")%>" align="middle" style="margin: -8px -3px 0 0;"></a> 
          <a href="custOrdInvoicePDF.asp?id=<%=pIdOrder%>" target="_blank"><%= dictLanguage.Item(Session("language")&"_CustviewOrd_75")%></a>
          <%if (Session("CustomerGuest")="0") AND (Session("idCustomer")>"0") then%> | <a href="RepeatOrder.asp?idOrder=<%=pIdOrder%>"><%= dictLanguage.Item(Session("language")&"_CustviewPastD_32")%></a>
          <% ''Hide/show link to Help Desk
            if scShowHD <> 0 then %>
            |&nbsp;<a href="userviewallposts.asp?idOrder=<%=clng(scpre)+clng(pIdOrder)%>"><%= dictLanguage.Item(Session("language")&"_viewPostings_3")%></a>
          <% end if %>
          <%end if%>
        </div>   
        
        <div style="width:20%;text-align: right;float: right;">
          <%if (Session("CustomerGuest")="0") AND (Session("idCustomer")>"0") then%>
          <a class="pcButton pcButtonBack" href="custViewPast.asp">
						<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
          </a>
          <%end if%>
        </div>
      </div>
    

      
      <%''GGG Add-on start %>
      <% 
        if gIDEvent<>"0" then
          query="select pcEvents.pcEv_name,pcEvents.pcEv_Date, pcEv_HideAddress,customers.name,customers.lastname from pcEvents,Customers where Customers.idcustomer=pcEvents.pcEv_idcustomer and pcEvents.pcEv_IDEvent=" & gIDEvent
          set rs1=connTemp.execute(query)
  
          if err.number<>0 then
            call LogErrorToDatabase()
            set rs1=nothing
            call closedb()
            response.redirect "techErr.asp?err="&pcStrCustRefID
          end if
  
          geName=rs1("pcEv_name")
          geDate=rs1("pcEv_Date")
  
          if year(geDate)="1900" then
            geDate=""
          end if
          if gedate<>"" then
            if scDateFrmt="DD/MM/YY" then
              gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
            else
              gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
            end if
          end if
          geHideAddress=rs1("pcEv_HideAddress")
          if geHideAddress="" then
            geHideAddress=0
          end if
          gReg=rs1("name") & " " & rs1("lastname")
          
          set rs1=nothing
        %>
        <div class="pcTableRow">
          <div class="pcShowContent">
            <div class="pcFormItem">
              <div class="pcFormLabel"><b><%= dictLanguage.Item(Session("language")&"_CustviewPastD_39")%></b></div>
              <div class="pcFormField"><%=gename%></div>
            </div>
            <div class="pcFormItem">
              <div class="pcFormLabel"><b><%= dictLanguage.Item(Session("language")&"_CustviewPastD_40")%></b></div>
              <div class="pcFormField"><%=geDate%></div>
            </div>
            <div class="pcFormItem">
              <div class="pcFormLabel"><b><%= dictLanguage.Item(Session("language")&"_CustviewPastD_41")%></b></div>
              <div class="pcFormField"><%=gReg%></div>
            </div>
          </div>
        </div>
        <% 
        else
          geHideAddress=0
        end if
      %>
      <% ''GGG Add-on end %>
      
      <% ''Start allow customer to nickname this order %>
      <% if scOrderName="1" then %>
        <div class="pcTableRow" id="pcOrderName">
          <div class="pcTableRowFull"><hr></div>
          <% 
            if pord_OrderName="" then
              pord_OrderName="No Name"
            end if
          %>
          <form method="post" name="form1" id="form1" action="CustViewPastD.asp" class="pcForms pcClear">
            <div class="pcShowContent">
              <input type=hidden name="action" value="rename">
              <input type=hidden name="IDOrder" value="<%=int(pIdOrder)+scpre%>">
            
              <div class="pcFormItem">
                <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_CustviewOrd_40")%></div>
                <div class="pcFormField"><input type="text" size="30" maxsize="50" name="ord_OrderName" value="<%=pord_OrderName%>">&nbsp;<input type="submit" name="Submit" value="Update" class="submit2"></div>
              </div>
            </div>
          </form>
          <div class="pcTableRowFull"><hr></div>
        </div>
      <% end if %>
      <% ''End allow customer to nickname this order %>
    
      <% ''START order delivery date, if any %>
      <%
        if (pord_DeliveryDate<>"") then
          %>
          <div class="pcTableRow" id="pcOrderDeliveryDate">
            <% 
            if scDateFrmt="DD/MM/YY" then        
              pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 4)
            else
              pord_DeliveryTime = FormatDateTime(pord_DeliveryDate, 3)
            end if
            pord_DeliveryDate = showdateFrmt(pord_DeliveryDate)
          
            ''Add <hr> only if the Order Name section is not shown 
            if not scOrderName="1" then 
            %>
              <div class="pcTableRowFull"><hr></div>
            <% 
            end if 
            %>
            <div class="pcTableRowFull">
              <%=dictLanguage.Item(Session("language")&"_CustviewOrd_39")%><%=pord_DeliveryDate%> <% if pord_DeliveryTime <> "00:00" then %><%=", " & pord_DeliveryTime%><% end if %>
            </div>
            <div class="pcTableRowFull"><hr></div>
          </div>
        <%
        end if
      %>
      <% ''END order delivery date %>
    
      <%
        pcShowShipping = false
        if pcShowShipAddr="1" and geHideAddress=0 then
          pcShowShipping = true
        end if
      %>
      
      <% ' 'START Billing and Shipping Addresses %>
      <div class="pcTableHeader">
        <div class="pcCustViewTableLabel">&nbsp;</div>
        <div class="pcCustViewTableField"><strong><%= dictLanguage.Item(Session("language")&"_orderverify_23")%></strong></div>
        <div class="pcCustViewTableField"><strong><%= dictLanguage.Item(Session("language")&"_orderverify_24")%></strong></div>
      </div>
      
      <% ''Billing/Shipping Name %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel">
          <%= replace(dictLanguage.Item(Session("language")&"_orderverify_7"),"''","'")%>
        </div>
        <div class="pcCustViewTableField">
          <%= pFirstName&" "&plastname %>
        </div>
        <div class="pcCustViewTableField">
          <%if pcShowShipping then%>
            <%= pshippingFullName %>
          <% end if%>
        </div>
      </div>
      
      <% ''Billing/Shipping Company %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel">
          <%= dictLanguage.Item(Session("language")&"_orderverify_8")%>
        </div>
        <div class="pcCustViewTableField">
          <%= pcustomerCompany %>
        </div>
        <div class="pcCustViewTableField">
          <% 
            if pcShowShipping then
              if pshippingCompany<>"" then %>
                <%= pshippingCompany %>
              <% end if
            end if 
          %>
        </div>
      </div>
      
      <% ''Billing/Shipping Email %>
      <% if pEmail <> pshippingEmail AND pshippingEmail <> "" then %>
        <div class="pcTableRow">
          <div class="pcCustViewTableLabel">
            <%=dictLanguage.Item(Session("language")&"_opc_5")%>
          </div>
          <div class="pcCustViewTableField">
            <%= pEmail %>
          </div>
          <div class="pcCustViewTableField">
            <% if pcShowShipping then %>
              <%= pshippingEmail %>
            <% end if %>
          </div>
        </div>
      <% end if %>
      
      <% ''Billing/Shipping Phone %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel">
          <%= dictLanguage.Item(Session("language")&"_orderverify_9") %>
        </div>
        <div class="pcCustViewTableField">
          <%= pPhone %>
        </div>
        <div class="pcCustViewTableField">   
          <% if pcShowShipping then %>
            <%= pshippingPhone %>
          <% end if %>
        </div>
      </div>
      
      <% ''Billing/Shipping Fax %>
      <% if pFax <> "" or pshippingFax <> "" then %>
        <div class="pcTableRow">
          <div class="pcCustViewTableLabel">
            <%= dictLanguage.Item(Session("language")&"_opc_18") %>
          </div>
          <div class="pcCustViewTableField">
            <%= pFax %>
          </div>
          <div class="pcCustViewTableField">   
            <% if pcShowShipping then %>
              <%= pshippingFax %>
            <% end if %>
          </div>
        </div>
      <% end if %>
      
      <% ''Billing/Shipping Address %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel">
          <%= dictLanguage.Item(Session("language")&"_orderverify_10") %>
        </div>
        <div class="pcCustViewTableField">
          <%= paddress %>
        </div>
        <div class="pcCustViewTableField">              
          <%
            if pcShowShipping then
              if pshippingAddress="" then
                response.write dictLanguage.Item(Session("language")&"_CustviewOrd_48")
              else
                response.write pshippingAddreses
              end if
            else
              if pcShowShipAddr="0" AND geHideAddress=0 then
                response.write dictLanguage.Item(Session("language")&"_CustviewOrd_48")
              end if
            end if 
          %>
        </div>
      </div>
      
      <% ''Billing/Shipping Address 2 %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel">
          &nbsp;
        </div>
        <div class="pcCustViewTableField">
          <%= paddress2 %>
        </div>
        <div class="pcCustViewTableField">   
          <% if pcShowShipping and pshippingAddress2 <> "" then %>
            <%= pshippingAddress2 %>
          <% end if %>
        </div>
      </div>
      
      <% ''Billing/Shipping City, State, and Zip Code %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel">
          &nbsp;
        </div>
        <div class="pcCustViewTableField">
          <%= pCity & ", " & pState & " " & pzip %>
        </div>
        <div class="pcCustViewTableField">   
          <% 
            if pcShowShipping and pshippingAddress<>"" then
              response.write pShippingCity&", "&pshippingState
              if pshippingState="" then
                response.write pshippingStateCode
              end if
              response.write " "&pshippingZip
            end if 
          %>
        </div>
      </div>
      
      <% ''Billing/Shipping Country Code %>
      <div class="pcTableRow">
        <div class="pcCustViewTableLabel">
          &nbsp;
        </div>
        <div class="pcCustViewTableField">
          <%= pCountryCode %>
        </div>
        <div class="pcCustViewTableField">   
          <%
            if pcShowShipping then
              response.write pshippingCountryCode
              strFedExCountryCode=pshippingCountryCode
            else
              strFedExCountryCode=pCountryCode
            end if 
          %>
        </div>
      </div>
          
      <% ''END Billing and Shipping Addresses %>
  
      <% ''START of payment details %>
      <%
        payment = split(pcpaymentDetails,"||")
        PaymentType=trim(payment(0))
        
        ''Get payment nickname
        query="SELECT paymentDesc,paymentNickName FROM paytypes WHERE paymentDesc = '" & replace(PaymentType,"'","''") & "';"
        Set rsTemp=Server.CreateObject("ADODB.Recordset")
        Set rsTemp=connTemp.execute(query)
          if err.number<>0 then
            call LogErrorToDatabase()
            set rsTemp=nothing
            call closedb()    
            response.redirect "techErr.asp?err="&pcStrCustRefID
          end if
          if not rsTemp.EOF then
            PaymentName=trim(rsTemp("paymentNickName"))
            else
            PaymentName=""
          end if
        Set rsTemp = nothing
        ''End get payment nickname
      
        ''Get authorization and transaction IDs, if any
        varTransID=""
        varTransName= dictLanguage.Item(Session("language")&"_CustviewPastD_102")
        varAuthCode=""
        varAuthName= dictLanguage.Item(Session("language")&"_CustviewPastD_103")
      
        if not isNull(pcpaymentCode) AND pcpaymentCode<>"" then 
          varShowCCInfo=0
          select case pcpaymentCode
          case "LinkPoint"
            varAry=split(pcgwAuthCode,":")
            varTransName="Approval Number"
            varAuthName="Reference Number"
            varTransID=left(varAry(1),6)
            varAuthCode=right(varAry(1),10)
          case "PFLink", "PFPro", "PFPRO", "PFLINK"
            varTransID=pcgwTransId
            varAuthCode=pcgwAuthCode
            varShowCCInfo=1
            varGWInfo="P"
          case "Authorize"
            varTransID=pcgwTransId
            varAuthCode=pcgwAuthCode
            varShowCCInfo=1
            if instr(ucase(PaymentType),"CHECK") then
              varShowCCInfo=0
            end if
            varGWInfo="A"
          case "twoCheckout"
            varTransName="2Checkout Order No"
            varTransID=pcgwTransId
          case "BOFA"
            varTransName="Order No"
            varAuthName="Authorization Code"
            varTransID=pcgwTransId
            varAuthCode=pcgwAuthCode
          case "WorldPay"
            varTransID=""
            varAuthCode=""
          case "iTransact"
            varTransName="Transaction ID"
            varAuthName="Authorization Code"
            varTransID=pcgwTransId
            varAuthCode=pcgwAuthCode
          case "PSI", "PSIGate"
            varTransName="Transaction ID"
            varAuthName="Authorization Code"
            varTransID=pcgwTransId
            varAuthCode=pcgwAuthCode
          case "fasttransact", "FastTransact", "FAST","CyberSource"
            varTransName="Transaction ID"
            varAuthName="Authorization Code"
            varTransID=pcgwTransId
            varAuthCode=pcgwAuthCode
          case "USAePay","FastCharge"
            varTransName="Transaction reference code"
            varAuthName="Authorization code"
            varTransID=pcgwTransId
            varAuthCode=pcgwAuthCode
          case "PxPay"
            varTransName="DPS Transaction Reference Number"
            varAuthName="Authorization code"
            varTransID=pcgwTransId
            varAuthCode=pcgwAuthCode
          end select
        end if
        
        ''End get authorization and transaction IDs
      
        if payment(1)="" then
         if err.number<>0 then
          PayCharge=0
         end if
          PayCharge=0
        else
          PayCharge=payment(1)
        end if
        err.number=0
        if instr(PaymentType,"FREE") AND len(PaymentType)<6 then
        else %>
        <div class="pcTableRow"><div class="pcTableRowFull"><hr></div></div>
        <div class="pcTableRow">
          <%=dictLanguage.Item(Session("language")&"_CustviewPastD_101")%>
          <%
            if PaymentName <> "" and PaymentName <> PaymentType then
              Dim pcv_strPaymentType
              Select Case PaymentType
                Case "PayPal Website Payments Pro": pcv_strPaymentType=PaymentName
                Case else: pcv_strPaymentType=PaymentName & " (" & PaymentType & ")"
              End Select
              Response.Write pcv_strPaymentType
              else
              Response.Write PaymentType
            end if
          %>
          <% if PayCharge>0 then %>
            <br><%=dictLanguage.Item(Session("language")&"_CustviewOrd_14b")%><%= " " & scCurSign&money(PayCharge)%>
          <% end if %>
          <% if varTransID<>"" then %>
          <br><%=varTransName%>: <%=varTransID%>
          <% end if %>
          <% if varAuthCode<>"" then %>
          <br><%=varAuthName%>: <%=varAuthCode%>
          <% end if %>
        </div>
      <% 
      end if
      %>
    <% '' END Payment details %>
    
    <% '' START Order Comments %>
    <% if len(pcomments)>3 then %>
      <div class="pcTableRow">
        <div class="pcTableRowFull">
          <strong><%= dictLanguage.Item(Session("language")&"_orderverify_11")%></strong>
          <%=pcomments%>
        </div>
      </div>
    <% end if %>
    <% '' END Order Comments %>

    <div class="pcTableRow">
      <div class="pcSpacer">&nbsp;</div>
    </div>
  
  <% ''START Order Details
  %>
		</div>
        <!--#include file="inc_ordercomplete.asp"-->
		<div class="pcTable">
      <!--RMA CREDIT-->
      <% if not isNull(prmaCredit) AND prmaCredit<>"" AND prmaCredit>0 then %>
        <div class="pcTableRow">
          <div class="pcCustviewPastD_Qty">&nbsp;</div>
          <div class="pcCustviewPastD_Subtotal">
            <b><%=  dictLanguage.Item(Session("language")&"_CustviewPastD_31")%></b>
          </div>
          <div class="pcCustviewPastD_Price">
            <% response.write "-"&scCurSign& money(prmaCredit) %>
          </div>
          <div>&nbsp;</div>
        </div>
      <% end if %>
    <!-- END Order Details -->
   
    
      <!-- START Other order information -->
    <div class="pcTable">
      <div class="pcTableRow"><div class="pcTableRowFull"><hr></div></div>
      <form method="post" name="form2" id="form2" action="#" class="pcForms">
        <div class="pcShowContent">
          <%
            'Start SDBA
            'Show PaymentStatus
          %>
          <div class="pcFormItem">
            <p><%
            response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_4")
            Select Case pcv_PaymentStatus
                Case 1: response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_5")
                Case 2: response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_6")
                Case 3:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_9")
                Case 4:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_10")
                Case 5:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_11")
                Case 6:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_12")
                Case 7:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_13")
                Case 8:  response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_14")
                Case else: response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_2")
            End Select%></p>
          </div>
          <%''End SDBA%>
          <% if piRewardPointsCustAccrued>0 AND int(pOrderStatus)>2 AND int(pOrderStatus)<>5 AND int(pOrderStatus)<>6 then %>
          <div class="pcFormItem"> 
            <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_16")%><%=piRewardPointsCustAccrued%>&nbsp;<%=RewardsLabel%><%= dictLanguage.Item(Session("language")&"_CustviewOrd_17")%>
            </p>
          </div>
          <% end if %>
		  <% if (piRewardPointsCustAccrued>0) AND (Session("CustomerGuest")="1") then %>
          <div class="pcFormItem"> 
            <p>
			<b><%= dictLanguage.Item(Session("language")&"_CustviewOrd_49")%><%=RewardsLabel%></b><br><br>
            </p>
          </div>
          <% end if %>
    
          <!-- if order was cancelled -->
          <% if int(pOrderStatus)=5 then %>
          <div class="pcFormItem"> 
            <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_18")%></p>
          </div>
          <% else %>
                    
          <!-- if order was returned -->
          <% if int(pOrderStatus)=6 then %>
          <div class="pcFormItem"> 
            <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_26")%></p>
          </div>
          <div class="pcFormItem"> 
            <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_37")%></p>
          </div>
          <div class="pcFormItem"> 
            <hr>
          </div>
          <% end if %>
          <!-- end order returned -->

  
          <!-- order has been processed, show date -->
          <% if int(pOrderStatus)>2 then %>
          <div class="pcFormItem"> 
            <p><%= dictLanguage.Item(Session("language")&"_CustviewPastD_22b")%></p>
          </div>
          <div class="pcFormItem"> 
            <p><%= dictLanguage.Item(Session("language")&"_CustviewPastD_22") & pprocessDate %></p>
          </div>
          <% else %>
          <!-- else if order has not been processed, tell customer -->
          <div class="pcFormItem"> 
            <p><%= dictLanguage.Item(Session("language")&"_CustviewPastD_20")%></p>
          </div>
          <% end if %>
          <!-- end order processed check -->  
      
          <!-- if order has been shipped, show information -->
          <% 
          if (int(pOrderStatus)=4 OR int(pOrderStatus)>= 6) then %>
          <div class="pcFormItem"> 
            <div class="pcTableRowFull"><hr></div>
          </div>
		  <% if int(pOrderStatus)=8 then %>
            <div class="pcFormItem"> 
              <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_50")%></p>
            </div>
          <% else %>
			  <% if int(pOrderStatus)=7 then %>
				<div class="pcFormItem"> 
				  <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_46")%>
				</div>
			  <% else %>
				<div class="pcFormItem"> 
				  <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_19")%>
				</div>
			  <% end if %>
			  <% if pShippingFullName<>"" then %>
				<div class="pcFormItem">
				  <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_20")%></p>
				</div>
				<div class="pcFormItem"> 
				  <p><%=pShippingFullName%></p>
				</div>
				<div class="pcFormItem">
				  <p><%=pShippingCompany%></p>
				</div>
				<div class="pcFormItem">
				  <p><%=pShippingAddress%></p>
				</div>
			  <% else %>
				<div class="pcFormItem">
				  <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_20")%></p>
				</div>
				<div class="pcFormItem">
				  <p><%=pShippingAddress%></p>
				</div>
			  <% end if %>
			  
			  <% if pShippingAddress2<>"" then %>
				<div class="pcFormItem">
				  <p><%=pShippingAddress2 %></p>
				</div>
			  <% end if %>
				
			  <div class="pcFormItem"> 
				<p><% response.write pShippingCity&", "&pshippingStateCode&" "&pShippingZip %></p>
			  </div>
			  <div class="pcFormItem"> 
				<p><%=pShippingCountryCode %></p>
			  </div>
			  <div class="pcFormItem"> 
				<div class="pcSpacer">&nbsp;</div>
			  </div>
          <%end if 'OrderStatus=8%>
          <%
          ''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          '' START: Shippment Information
          ''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~       
          %>  
          <% if pshipDate="//" OR isNULL(pshipVia)=True then %>
            <%
            Dim rsShipInfo
            query="SELECT pcPackageInfo_ID, pcPackageInfo_ShipMethod,pcPackageInfo_TrackingNumber,pcPackageInfo_ShippedDate,pcPackageInfo_Comments,pcPackageInfo_UPSPackageType FROM pcPackageInfo WHERE idOrder=" & pidorder & ";"
            set rsPackages=server.CreateObject("ADODB.RecordSet")
            set rsPackages=connTemp.execute(query)
            if not rsPackages.eof then
              pcIdNum = 1               
              do while not rsPackages.eof 
                pcv_PackageID = rsPackages("pcPackageInfo_ID")
                ptrackingNum = ""
                pshipVia = ""
                tmp_ShipMethod=rsPackages("pcPackageInfo_ShipMethod")
                tmp_TrackingNumber=rsPackages("pcPackageInfo_TrackingNumber")
                tmp_ShippedDate=rsPackages("pcPackageInfo_ShippedDate")
                tmp_UPSPackageType=rsPackages("pcPackageInfo_UPSPackageType")
                ptrackingNum=tmp_TrackingNumber
                pshipVia=tmp_ShipMethod
                  
                ''// Show the Shipment Info for v3 package
                query="SELECT quantity, Description FROM Products INNER JOIN ProductsOrdered ON (products.idproduct=ProductsOrdered.idproduct) WHERE productsOrdered.pcPackageInfo_ID=" & pcv_PackageID & " ORDER BY ProductsOrdered.pcPackageInfo_ID;"
                set rsShipInfo=server.CreateObject("ADODB.RecordSet")
                set rsShipInfo=conntemp.execute(query)
                if err.number<>0 then
                  call LogErrorToDatabase()
                  set rsShipInfo=nothing
                  call closedb()
                  response.redirect "techErr.asp?err="&pcStrCustRefID
                end if
                %>
                <div class="pcSectionTitle"> 
                  <p><strong><%= dictLanguage.Item(Session("language")&"_CustviewOrd_41")%></strong></p>
                </div>
                <div class="pcFormItem"> 
                  <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_21")%>&nbsp;<%=ShowDateFrmt(tmp_ShippedDate) %></p>
                </div>
                <div class="pcFormItem"> 
                  <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_22")%>&nbsp;<%=tmp_ShipMethod %></p>
                </div>
                <% 
                ''Show Tracking Number

                if ptrackingNum<>"" then 
                  %>
                  <div class="pcFormItem"> 
                    
                    <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_25")%>
                    <% 
                    ''//  Start: Tracking Link
                    if instr(ucase(tmp_ShipMethod),"UPS:") OR tmp_UPSPackageType<>"" then %>
                      <a href="custUPSTracking.asp?itracknumber=<%=ptrackingNum%>"><%=ptrackingNum %></a>
                    <% ElseIf instr(ucase(tmp_ShipMethod),"FEDEX:") then %>
                      &nbsp;
                                                  <a href="http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=<%=ptrackingNum%>" target="_blank"><%=ptrackingNum %></a>
                    <% else 
                        response.write " " & ptrackingNum
                    end if 
                    ''//  End: Tracking Link 
                    %>
                    </p>
                    
                  </div>
                  <% 
                end if
                ''end if Tracking Number 
                %>
                <div class="pcFormItem"> 
                  <p><a href="JavaScript:pcf_ShowContents('PackageTable<%=pcIdNum%>');"><%= dictLanguage.Item(Session("language")&"_CustviewOrd_42")%></a></p>
                </div>
                <div class="pcFormItem"> 
                  <script type=text/javascript>
                  function pcf_ShowContents(obj){                   
                    if(document.getElementById){                      
                    var tablename = document.getElementById(obj);                     
                      if(tablename.style.display != ''){
                        tablename.style.display='';
                      } else {
                        tablename.style.display = 'none';
                      }
                    }
                  }
                  </script>                  
                    <div id="PackageTable<%=pcIdNum%>" class="pcTable" style="display: none;">
                      <div class="pcTableHeader" style="background-color: #ffffcc">
                        <div style="width: 90%;">Product Name</div>
                        <div style="width: 5%">Qty</div>
                      </div> 
                      <%
                      if not rsShipInfo.eof then
                        do while not rsShipInfo.eof
                          if tmp_ShipMethod<>"" OR tmp_TrackingNumber<>"" OR tmp_ShippedDate<>"" then       
                          pcv_PrdName=rsShipInfo("Description")
                          pcv_PrdQty=rsShipInfo("quantity")
                          %>                        
                          <div class="pcTableRow">
                            <div style="width: 90%;"><%=pcv_PrdName%></div>
                            <div style="width: 5%"><%=pcv_PrdQty%></div>
                          </div>             
                          <%
                          pcIdNum = pcIdNum + 1
                          end if
                          rsShipInfo.movenext
                        loop
                        set rsShipInfo = nothing
                      end if %>
                    </div>
                </div>
                <% rsPackages.movenext
              loop
            end if
            set rsPackages = nothing
            %>
            <div class="pcFormItem">
              <div class="pcTableRowFull"><hr></div>
            </div> 
            <% if pOrdPackageNum <> "" then %>
              <% if int(pOrderStatus)=7 then %>             
                <div class="pcFormItem"> 
                  <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_43")%>&nbsp;<%=pOrdPackageNum %></p>
                </div>
              <% else %>
                <div class="pcFormItem"> 
                  <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_38")%>&nbsp;<%=pOrdPackageNum %></p>
                </div>
              <% end if %>
            <% end if %>
            <% if pOrdPackageNum <> "" AND int(pOrderStatus)=7 then %>
              <div class="pcFormItem"> 
                <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_44")%><%=(pcIdNum - 1)%><%= dictLanguage.Item(Session("language")&"_CustviewOrd_45")%><%=pOrdPackageNum %></p>
              </div>
            <% end if %>
          <% else %>        
            <%
            ''// Show the Shipment Info for v2.76 package
            %>
            <div class="pcFormItem"> 
              <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_21")%><%=pshipDate %></p>
            </div>
            <div class="pcFormItem"> 
              <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_22")%><%=pshipVia %></p>
            </div>
            <div class="pcFormItem"> 
              <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_38")%><%=pOrdPackageNum %></p>
            </div>
            
            <% ''Show Tracking Number
            if ptrackingNum<>"" then %>
              <div class="pcFormItem"> 
                <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_25")%>
                <% if instr(ucase(pshipVia),"UPS") then %>
                  <a href="custUPSTracking.asp?itracknumber=<%=ptrackingNum%>"><%=ptrackingNum %></a>
                <% else 
                  if instr(ucase(pshipVia),"FEDEX") then 
                    if ucase(strFedExCountryCode)="US" then %>
                      <a href="http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=<%=ptrackingNum%>" target="_blank"><%=ptrackingNum %></a>
                    <% else %>
                      <a href="http://www.fedex.com/Tracking?cntry_code=<%=strFedExCountryCode%>" target="_blank"><%=ptrackingNum %></a>
                    <% end if %>
                  <% else 
                    response.write ptrackingNum
                  end if
                end if %>
                </p>
              </div>
            <% 
            end if
            ''end if Tracking Number 
            %>          
          
          <% end if %>
          <%
          ''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          '' END: Shippment Information
          ''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~       
          %>
            
      
          <!-- if RMA has not been issued, show link to RMA request form, otherwise show message -->
          <%
          if scHideRMA = 0 then '' START - Check if the store allows customers to request an RMA
            Dim rsRma, rmaVar, rmaNumber, rmaReturnStatus, queryrma
        
            queryrma="SELECT rmaNumber, rmaReturnStatus, rmaApproved FROM PCReturns WHERE idOrder=" &pIdOrder
            set rsRma=conntemp.execute(queryrma)
            if err.number<>0 then
              call LogErrorToDatabase()
              set rsRma=nothing
              call closedb()
              response.redirect "techErr.asp?err="&pcStrCustRefID
            end if
            if not rsRma.eof then
              rmaNumber = rsRma("rmaNumber")
              rmaReturnStatus = rsRma("rmaReturnStatus")
              rmaApproved = rsRma("rmaApproved")
              ''0=pending, 1=approved, 2=denied
              rmaVar = 1
            else
              rmaVar = 0
            end if
        
            Set rsRma = nothing
            %>        
            <div class="pcFormItem">
              <div class="pcTableRowFull"><hr></div>
            </div>
            <div class="pcSectionTitle">
              <%= dictLanguage.Item(Session("language")&"_CustviewOrd_47")%>
            </div> 
            <% 
            if rmaVar = 0 then
            '' RMA can be requested. The customer has not requested it. Show link.
            %>
              <div class="pcFormItem">
                <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_23")%><a href="rmaindex.asp?idorder=<%=pIdOrder%>"><%= dictLanguage.Item(Session("language")&"_CustviewOrd_24")%></a></p>
              </div>
            <% 
            else '' An RMA has already been requested by the customer or issued by the store manager
              if rmaApproved=0 then '' An RMA request has not yet been approved
              %>
                <div class="pcFormItem">
                  <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_30")%></p>
                </div>
              <%
              end if 
              if rmaApproved=1 then '' An RMA request has been approved
              %>
                <div class="pcFormItem"><p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_34")%></p></div>
                <div class="pcFormItem"><p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_31")%> <b><%=rmaNumber%></b></p></div>
              <%
              end if 
              if rmaApproved=2 then '' An RMA request has been denied
              %>
                <div class="pcFormItem"><div class="pcTableRowFull"><hr></div></div>
                <div class="pcFormItem"><p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_35")%></p></div>
              <%
              end if %>
              <%
              if trim(RmaReturnStatus) <> "" then '' Admin comments related to the RMA request
              %>
                <div class="pcFormItem">
                  <p><%= dictLanguage.Item(Session("language")&"_CustviewOrd_32")%>&nbsp;<%=RmaReturnStatus%></p>
                </div>
              <%
              end if '' End RMA Comments
            end if '' End RMA has already been requested
          end if '' END - Check if the store allows customers to request an RMA
          %>
          <!-- End RMA link -->
      
          <!-- end shipping info -->
        <% end if
        end if  %>
    
        <!-- START GGG Infor and Downloadable Products Information -->
        <div class="pcFormItem"> 
          <%''GGG Add-on start
          if (GCDetails<>"") then %>
          <div class="pcTableRowFull"><hr></div>
            <div class="pcTable">
              <div class="pcTableHeader">
                <div><%= dictLanguage.Item(Session("language")&"_CustviewPastD_45")%></div>
              </div>
              <%
                GCArry=split(GCDetails,"|g|")
                intArryCnt=ubound(GCArry)
      
                for k=0 to intArryCnt
        
                if GCArry(k)<>"" then
                  GCInfo = split(GCArry(k),"|s|")
                  if GCInfo(2)="" OR IsNull(GCInfo(2)) then
                  GCInfo(2)=0
                  end if
                  pGiftCode=GCInfo(0)
                  pGiftUsed=GCInfo(2)
                query="select products.IDProduct,products.Description from pcGCOrdered,Products where products.idproduct=pcGCOrdered.pcGO_idproduct and pcGCOrdered.pcGO_GcCode='"& pGiftCode & "'"
                set rsG=connTemp.execute(query)

                if not rsG.eof then
                    pIdproduct=rsG("idproduct")
                    pName=rsG("Description")
                    pCode=pGiftCode
                    %>
              <div class="pcTableRow"> 
                <div style="width: 18%;"><b><%= dictLanguage.Item(Session("language")&"_CustviewPasdiv_46")%></b></div>
                <div style="width: 82%"><b><%=pName%></b></div>
              </div>
              <div class="pcTableRow"> 
                <div style="width: 18%;">&nbsp;</div>
                <div style="width: 82%">
                    <%
                    query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_GcCode='" & pGiftCode & "'"
                    set rs19=connTemp.execute(query)
    
                    if not rs19.eof then%>
                        <%= dictLanguage.Item(Session("language")&"_CustviewPastD_47")%><b><%=rs19("pcGO_GcCode")%></b><br>
                        <%= dictLanguage.Item(Session("language")&"_CustviewPastD_48")%><%=scCurSign & money(pGiftUsed)%><br><br>
                        <%
                        pGCAmount=rs19("pcGO_Amount")
                        if cdbl(pGCAmount)<=0 then%>
                            <%= dictLanguage.Item(Session("language")&"_CustviewPastD_49")%>
                        <%else%>
                            <%= dictLanguage.Item(Session("language")&"_CustviewPastD_50")%><b><%=scCurSign & money(pGCAmount)%></b>
                            <br>
                            <%pExpDate=rs19("pcGO_ExpDate")
                            if year(pExpDate)="1900" then%>
                                <%= dictLanguage.Item(Session("language")&"_CustviewPastD_51")%>
                            <%else
                                if scDateFrmt="DD/MM/YY" then
                                    pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
                                else
                                    pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
                                end if%>
                                <%= dictLanguage.Item(Session("language")&"_CustviewPastD_52")%><font color=#ff0000><b><%=pExpDate%></b></font>
                            <%end if%>
                            <br>
                            <%
                            pGCStatus=rs19("pcGO_Status")
                            if pGCStatus="1" then%>
                                <%= dictLanguage.Item(Session("language")&"_CustviewPastD_53")%><%= dictLanguage.Item(Session("language")&"_CustviewPastD_53a")%>
                            <%else%>
                                <%= dictLanguage.Item(Session("language")&"_CustviewPastD_53")%><%= dictLanguage.Item(Session("language")&"_CustviewPastD_53b")%>
                            <%end if%>
                        <%end if%>
                        <br><br>
                    <%end if
                    set rs19=nothing%>
                </div>
              </div>
              <%end if
              set rsG=nothing
              end if
              Next%>
            </div>
            <% end if
            ''GGG Add-on end%>
                  
            <% 
            ''// we do not hide the download link on partial return
            if (int(pOrderStatus)>2 AND int(pOrderStatus)<=4) OR (int(pOrderStatus)>=7)  then
              if (pcDPs<>"") and (pcDPs="1") then %>
            <hr>
            <div class="pcTable">
              <div class="pcTableHeader">
                <div>
                  <%= dictLanguage.Item(Session("language")&"_CustviewPastD_23")%>
                </div>
              </div>
                    
              <% query="select IdProduct from DPRequests WHERE IdOrder=" & pidorder & ";"
              set rsLic=connTemp.execute(query)
              if err.number<>0 then
                  call LogErrorToDatabase()
                  set rsLic=nothing
                  call closedb()
                  response.redirect "techErr.asp?err="&pcStrCustRefID
              end if
              do while not rsLic.eof
                  pIdproduct=rsLic("idproduct")
                  query="select Description, URLExpire, ExpireDays, License, LicenseLabel1, LicenseLabel2, LicenseLabel3, LicenseLabel4, LicenseLabel5 from Products,DProducts where products.idproduct=" & pIdproduct & " and DProducts.idproduct=Products.idproduct and products.downloadable=1"
                  set rstemp=connTemp.execute(query)
                  if err.number<>0 then
                      call LogErrorToDatabase()
                      set rstemp=nothing
                      call closedb()
                      response.redirect "techErr.asp?err="&pcStrCustRefID
                  end if

                  if not rstemp.eof then
                      pName=rstemp("Description")
                      pURLExpire=rstemp("URLExpire")
                      pExpireDays=rstemp("ExpireDays")  
                      pLicense=rstemp("License")
                      pLL1=rstemp("LicenseLabel1")
                      pLL2=rstemp("LicenseLabel2")
                      pLL3=rstemp("LicenseLabel3")
                      pLL4=rstemp("LicenseLabel4")
                      pLL5=rstemp("LicenseLabel5")
                      
                      set rstemp = nothing
      
                      query="select RequestSTR,StartDate from DPRequests where idproduct=" & pIdproduct & " and idorder=" & pidorder & " and idcustomer=" & pidcustomer
                      set rstemp=connTemp.execute(query)
                      if err.number<>0 then
                          call LogErrorToDatabase()
                          set rstemp=nothing
                          call closedb()
                          response.redirect "techErr.asp?err="&pcStrCustRefID
                      end if
                      pdownloadStr=rstemp("RequestSTR")
                      pStartDate=rstemp("StartDate")
                      SPath1=split(Request.ServerVariables("PATH_INFO"),"/pc/")
                      
                      if SPath1(0)<>"" then
                      else
                          SPath1(0)="/"
                      end if
                      
                      if SPath1(0)<>"/" then
                          if Left(SPath1(0),1)<>"/" then
                              SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & "/" & SPath1(0)
                          else
                              SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1(0)
                          end if
                      else
                          SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & "/"
                      end if
  
                      if Right(SPathInfo,1)="/" then
                          pdownloadStr=SPathInfo & "pc/pcdownload.asp?id=" & pdownloadStr         
                      else
                          pdownloadStr=SPathInfo & "/pc/pcdownload.asp?id=" & pdownloadStr
                      end if
  
                      set rstemp=nothing %>
              <div class="pcTableRow"> 
                <div style="width: 18%;">
                  <p><%= dictLanguage.Item(Session("language")&"_CustviewPastD_24")%></p>
                </div>
                <div style="width: 82%;"><p><b><%=pName%></b></p></div>
              </div>
              <div class="pcTableRow"> 
                <div style="width: 18%;">
                  <p><%= dictLanguage.Item(Session("language")&"_CustviewPastD_25")%></p>
                </div>
                <div style="width: 82%;">
                  <p><a href="<%=pdownloadStr%>" target="_blank"><%=pdownloadStr%></a></p>
                  <p>
                    <% if (pURLExpire<>"") and (pURLExpire="1") then
                      if date()-(CDate(pStartDate)+pExpireDays)<0 then%>
                        <%= dictLanguage.Item(Session("language")&"_CustviewPastD_26")%><%=(CDate(pStartDate)+pExpireDays)-date()%><%= dictLanguage.Item(Session("language")&"_CustviewPastD_27")%>
                      <%else
                        if date()-(CDate(pStartDate)+pExpireDays)=0 then%>
                          <p><%= dictLanguage.Item(Session("language")&"_CustviewPastD_28")%></p>
                        <%else%>
                          <p><%= dictLanguage.Item(Session("language")&"_CustviewPastD_29")%></p>
                        <%end if
                      end if
                    end if%>
                  </p>
                </div>
              </div>
              <%if (pLicense<>"") and (pLicense="1") then %>
              <div class="pcTableRow"> 
                <div style="width: 18%;">
                  <p><%= dictLanguage.Item(Session("language")&"_CustviewPastD_30")%></p>
                </div>
                <div style="width: 82%;">
                  <% query="select Lic1, Lic2, Lic3, Lic4, Lic5 from DPLicenses where idproduct=" & pIdproduct & " and idorder=" & pidorder
                  set rstemp=connTemp.execute(query)
                  if err.number<>0 then
                      call LogErrorToDatabase()
                      set rstemp=nothing
                      call closedb()
                      response.redirect "techErr.asp?err="&pcStrCustRefID
                  end if
                  do while not rstemp.eof
                      Lic1=rstemp("Lic1")
                      Lic2=rstemp("Lic2")
                      Lic3=rstemp("Lic3")
                      Lic4=rstemp("Lic4")
                      Lic5=rstemp("Lic5")
                      %>
                    <div class="pcShowContent">
                      <% if Lic1<>"" then%>
                          <div class="pcFormItem">
                            <div><p><%=pLL1%>:</p></div>
                            <div><p><%=Lic1%></p></div>
                          </div>
                      <%end if
                      if Lic2<>"" then%>
                          <div class="pcFormItem">
                            <div><p><%=pLL2%>:</p></div>
                            <div><p><%=Lic2%></p></div>
                          </div>
                      <%end if
                      if Lic3<>"" then%>
                          <div class="pcFormItem">
                            <div><p><%=pLL3%>:</p></div>
                            <div><p><%=Lic3%></p></div>
                          </div>
                      <%end if
                      if Lic4<>"" then%>
                          <div class="pcFormItem">
                            <div><p><%=pLL4%>:</p></div>
                            <div><p><%=Lic4%></p></div>
                          </div>
                      <%end if
                      if Lic5<>"" then%>
                          <div class="pcFormItem">
                            <div><p><%=pLL5%>:</p></div>
                            <div><p><%=Lic5%></p></div>
                          </div>
                      <%end if%>
                    </div>
                    <%rstemp.movenext
                  loop
                  set rstemp=nothing
                  %>
                </div>
              </div>
              <%end if
              end if
                rsLic.MoveNext
                loop
                set rsLic=nothing
              %>
            </div>
            <!-- END Downloadable products -->

            <% 
              end if
              end if ''Order Status 3, 4
            %>
                              
            <%''GGG Add-on start
              if (int(pOrderStatus)>2 AND int(pOrderStatus)<=4) OR (int(pOrderStatus)>=7) then
              if (pGCs<>"") and (pGCs="1") then
               %>
            <hr>
            <div class="pcTable">
              <div class="pcTableHeader">
                <div><%= dictLanguage.Item(Session("language")&"_CustviewPastD_33")%></div>
              </div>
              <%
              query="select * from ProductsOrdered WHERE idOrder="& pidorder
              set rs11=connTemp.execute(query)
              do while not rs11.eof
                  query="select products.Description,pcGCOrdered.pcGO_GcCode from Products,pcGCOrdered where products.idproduct=" & rs11("idproduct") & " and pcGCOrdered.pcGO_idproduct=Products.idproduct and products.pcprod_GC=1 and pcGCOrdered.pcGO_idOrder="& pidorder
                  set rsG=connTemp.execute(query)

                  if not rsG.eof then
                      pIdproduct=rs11("idproduct")
                      pGCName=rsG("Description")
                      pCode=rsG("pcGO_GcCode")
                      %>
                      <div class="pcTableRow">
                        <div style="width:18%;"><b><%= dictLanguage.Item(Session("language")&"_CustviewPastD_34")%></b></div>
                        <div style="width:82%;"><b><%=pGCName%></b></div>
                      </div>
                      <div class="pcTableRow"> 
                        <div style="width:18%;">&nbsp;</div>
                        <div style="width:82%;">
                        <%
                        query="select pcGO_GcCode,pcGO_ExpDate,pcGO_Amount,pcGO_Status from pcGCOrdered where pcGO_idproduct=" & rs11("idproduct") & " and pcGO_idorder=" & pidorder
                        set rs19=connTemp.execute(query)

                        do while not rs19.eof%>
                            <%= dictLanguage.Item(Session("language")&"_CustviewPastD_35")%>&nbsp;<b><%=rs19("pcGO_GcCode")%></b><br>
                            <%pExpDate=rs19("pcGO_ExpDate")
                            if year(pExpDate)="1900" then%>
                                <%= dictLanguage.Item(Session("language")&"_CustviewPastD_36b")%>
                            <%else
                                if scDateFrmt="DD/MM/YY" then
                                    pExpDate=day(pExpDate) & "/" & month(pExpDate) & "/" & year(pExpDate)
                                else
                                    pExpDate=month(pExpDate) & "/" & day(pExpDate) & "/" & year(pExpDate)
                                end if%>
                                <%= dictLanguage.Item(Session("language")&"_CustviewPastD_36")%>&nbsp;<font color=#ff0000><b><%=pExpDate%></b></font>
                            <%end if%>
                            <br>
                            <%
                            pGCAmount=rs19("pcGO_Amount")
                            if cdbl(pGCAmount)<=0 then%>
                                <%= dictLanguage.Item(Session("language")&"_CustviewPastD_37b")%>
                            <%else%>
                                <%= dictLanguage.Item(Session("language")&"_CustviewPastD_37")%>&nbsp;<b><%=scCurSign & money(pGCAmount)%></b>
                            <%end if%><br>
                            <%
                            pGCStatus=rs19("pcGO_Status")
                            if pGCStatus="1" then%>
                                <%= dictLanguage.Item(Session("language")&"_CustviewPastD_38")%>&nbsp;<%= dictLanguage.Item(Session("language")&"_CustviewPastD_38a")%>
                            <%else%>
                                <%= dictLanguage.Item(Session("language")&"_CustviewPastD_38")%>&nbsp;<%= dictLanguage.Item(Session("language")&"_CustviewPastD_38b")%>
                            <%end if%>
                            <br><br>
                            <%  rs19.movenext
                        loop
                        set rs19=nothing
                        %>
                        </div>
                      </div>
                  <%end if
                  set rsG=nothing
                  rs11.MoveNext
              loop
              set rs11=nothing
              %>
            </div>
          <% end if
          end if
          ''GGG Add-on end%>
        </div>
        <%'' ------------------------------------------------------
        ''Start SDBA - Notify Drop-Shipping
        '' ------------------------------------------------------
        if scShipNotifySeparate="1" then
            
            tmp_showmsg=0
            query="SELECT products.pcProd_IsDropShipped FROM products INNER JOIN productsOrdered ON (products.idproduct=productsOrdered.idproduct AND products.pcProd_IsDropShipped=1) WHERE ProductsOrdered.idOrder=" & pIdOrder & ";"
            set rs=connTemp.execute(query)
            if err.number<>0 then
                call LogErrorToDatabase()
                set rs=nothing
                call closedb()
                response.redirect "techErr.asp?err="&pcStrCustRefID
            end if
            if not rs.eof then
                tmp_showmsg=1
            end if
            set rs=nothing
            if tmp_showmsg=1 then%>
            <div class="pcFormItem"> 
              <div class="pcSpacer">&nbsp;</div>
            </div>
            <div class="pcFormItem">
              <div class="pcTextMessage"><%= ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%></div>
            </div>
        <%end if
          
        end if
        '' ------------------------------------------------------
        ''End SDBA - Notify Drop-Shipping
        '' ------------------------------------------------------%>
        </div>
      </form>

      <%
      ''Gift Certificates Recipient Information
      ''Show only if the order has been processed
      if int(pOrderStatus)>2 AND int(pOrderStatus)<>5 AND int(pOrderStatus)<>6 then
        
        query="SELECT pcOrd_GcReName,pcOrd_GcReEmail,pcOrd_GcReMsg FROM Orders WHERE idOrder="& pidorder &" AND pcOrd_GcReEmail<>'';"
        SET rsGCObj=server.CreateObject("ADODB.RecordSet")
        SET rsGCObj=connTemp.execute(query)
        if not rsGCObj.eof then
          Gc_ReName=rsGCObj("pcOrd_GcReName")
          Gc_ReEmail=rsGCObj("pcOrd_GcReEmail")
          Gc_ReMsg=rsGCObj("pcOrd_GcReMsg")
          %>
      <form method="post" name="form3" action="CustViewPastD.asp?action=resend" class="pcForms">
        <div class="pcShowContent">
          <div class="pcFormItem">
            <div class="pcCPspacer">&nbsp;</div>
          </div>
          <div class="pcFormItem">
            <div class="pcTableHeader"><%= dictLanguage.Item(Session("language")&"_GCRecipient_1")%></div>
          </div>
          <div class="pcFormItem">
            <div class="pcCPspacer">&nbsp;</div>
          </div>
          <div class="pcFormItem"> 
            <div class="pcFormLabel"><b><%= dictLanguage.Item(Session("language")&"_NotifyRe_3")%></b></div>
            <div class="pcFormField"><input type="text" name="GC_RecName" size="30" value="<%=Gc_ReName%>"></b></div>
          </div>
          <div class="pcFormItem"> 
            <div class="pcFormLabel"><b><%= dictLanguage.Item(Session("language")&"_NotifyRe_4")%></b></div>
            <div class="pcFormField"><input type="text" name="GC_RecEmail" size="30" value="<%=Gc_ReEmail%>"></b></div>
          </div>
          <div class="pcFormItem"> 
            <div class="pcFormLabel"><b><%= dictLanguage.Item(Session("language")&"_NotifyRe_5")%></b></div>
            <div class="pcFormField"><textarea name="GC_RecMsg" cols="60" rows="5" wrap="VIRTUAL"><%=GC_ReMsg%></textarea></div>
          </div>
          <div class="pcFormItem"> 
            <div class="pcFormLabel">&nbsp;</div>
            <div class="pcFormField">
              <input type="hidden" name="idOrder" value="<%=int(pIdOrder)+scpre%>">
              <input type="submit" name="submitReSendGCRec" value="<%= dictLanguage.Item(Session("language")&"_GCRecipient_2")%>">
            </div>
          </div>
        </div>
      </form>
      <%
        end if
        set rsGCObj=nothing                 
                  
        end if  
      %>
      <%''SHW-S
        on error goto 0
        call GetSHWSettings()
        if shwOnOff=1 then
          queryQ="SELECT idOrder,pcSWO_ShipwireID,pcSWO_ShipwireDetails FROM pcShipwireOrders WHERE idOrder=" & pIdOrder & ";"
          set rsQ=connTemp.execute(queryQ)
          if not rsQ.eof then
            tmpArr=rsQ.getRows()
            intCountQ=ubound(tmpArr,2)
            set rsQ=nothing%>
      <div class="pcShowContent">
        <div class="pcTableHeader">
          <div><strong>SHIPWIRE SHIPPING INFORMATION</strong></div>
        </div>
        <%For iQ=0 to intCountQ
        %>
        <div class="pcFormItem">
          <b>Shipwire Package ID#: <%=tmpArr(1,iQ)%></b>
        </div>
        <div class="pcFormItem">
          <b>Package Details</b>
        </div>
        <div class="pcFormItem">
          <%=tmpArr(2,iQ)%>
        </div>
        <%=SHWGetPackStatusCust(tmpArr(1,iQ))%>
        <div class="pcFormItem">
            <div class="pcTableRowFull"><hr></div>
        </div>
        <div class="pcFormItem">
          <div class="pcCPspacer"></div>
        </div>
        <%Next%>
      </div>
      <%end if
        set rsQ=nothing

      end if
      ''SHW-E%>
    </div>
    <!-- End Other Order Information -->

    <% if Session("CustomerGuest")="1" then %>
    <div id="PwdArea">
        <form id="PwdForm" name="PwdForm">

            <div class="row">
                <div class="col-xs-12">
                    <h4><%=dictLanguage.Item(Session("language")&"_opc_common_2")%><h4>
                </div>
            </div>
            <div class="row">
                <div class="col-xs-12">
					<% if (piRewardPointsCustAccrued>0) then%>
                    	<%=dictLanguage.Item(Session("language")&"_opc_common_3a")%><%=RewardsLabel%><%=dictLanguage.Item(Session("language")&"_opc_common_3b")%>
					<%else%>
						<%=dictLanguage.Item(Session("language")&"_opc_common_3")%>
					<%end if%>
                </div>
            </div>
            <div class="row">
                <div class="col-xs-2">
                    <%=dictLanguage.Item(Session("language")&"_opc_6")%>
                </div>
                <div class="col-xs-3">
                    <input type="password" name="newPass1" id="newPass1" size="20">
                </div>
            </div>
            <div class="row">
                <div class="col-xs-2">
                    <%=dictLanguage.Item(Session("language")&"_opc_38")%>
                </div>
                <div class="col-xs-3">
                    <input type="password" name="newPass2" id="newPass2" size="20">
                </div>
            </div>
            <div class="row">
                <div class="col-xs-12">
                    <div style="padding-top: 10px;"></div>
                </div>
            </div>
            <div class="row">
                <div class="col-xs-12">
                    <input type="button" name="PwdSubmit" id="PwdSubmit" value="<%=dictLanguage.Item(Session("language")&"_opc_common_4")%>" class="submit2">
                </div>
            </div>

        </form>
        <div id="PwdLoader" style="display:none"></div>
    </div>
    <% end if %>
    
    <div class="pcShowContent">
        <% ''// Account Consolidation %>
        <!--#include file="opc_inc_CustConsolidate.asp"-->
    </div>
    <div class="pcShowContent">
        <script type=text/javascript>
        $pc(document).ready(function()
        {
          jQuery.validator.setDefaults({
            success: function(element) {
              $pc(element).parent("td").children("input, textarea").addClass("success")
            }
          });

          <%if Session("CustomerGuest")="1" then
          Session("SFStrRedirectUrl")="CustPref.asp"%>
          //*Validate Password Form
          $pc("#PwdForm").validate({
            rules: {
              newPass1: 
              {
                required: true
              },
              newPass2:
              {
                required: true,
                equalTo: "#newPass1"
              }
            },
            messages: {
              newPass1: {
                required: "<%=dictLanguage.Item(Session("language")&"_opc_js_4")%>",
                minlength: "<%=dictLanguage.Item(Session("language")&"_opc_js_5")%>"
              },
              newPass2: {
                required: "<%=dictLanguage.Item(Session("language")&"_opc_js_47")%>",
                minlength: "<%=dictLanguage.Item(Session("language")&"_opc_js_5")%>",
                equalTo: "<%=dictLanguage.Item(Session("language")&"_opc_js_48")%>"
              }
            }
          })
          
          $pc('#PwdSubmit').click(function(){
            if ($pc('#PwdForm').validate().form())
            {
              $pc("#PwdLoader").html('<img src="<%=pcf_getImagePath("images","ajax-loader1.gif")%>" width="20" height="20" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_5")%>');
              $pc("#PwdLoader").show(); 
              $pc.ajax({
                type: "POST",
                url: "opc_createacc.asp",
                data: $pc('#PwdForm').formSerialize() + "&action=create",
                timeout: 5000,
                success: function(data, textStatus){
                  if (data=="SECURITY")
                  {
                    $pc("#PwdArea").html("");
                    $pc("#PwdArea").hide();
                    $pc("#PwdLoader").html('<img src="<%=pcf_getImagePath("images","pc_icon_error_small.png")%>" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_6")%>');
                    var callbackPwd=function (){setTimeout(function(){$pc("#PwdLoader").hide();},1000);}
                    $pc("#PwdLoader").effect('pulsate',{},500,callbackPwd);
                  }
                  else
                  {
                  if ((data=="OK") || (data=="REG") || (data=="OKA") || (data=="REGA"))
                  {

                    if ((data=="OK") || (data=="OKA"))
                    {
                    	$pc("#PwdLoader").html('<img src="<%=pcf_getImagePath("images","pc_icon_success_small.png")%>" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_7")%>');
                    }
                    else
                    {
                    	$pc("#PwdLoader").html('<img src="<%=pcf_getImagePath("images","pc_icon_success_small.png")%>" align="absmiddle"><%=dictLanguage.Item(Session("language")&"_opc_common_8")%>');
                    }
                    var callbackPwd=function (){}
                    $pc("#PwdLoader").effect('pulsate',{},500,callbackPwd);
                    $pc("#PwdArea").html("");
                    $pc("#PwdArea").hide();
                    if (data=="OKA")
                    {
                      $pc("#ConArea").show();
                    }
                    else
                    {
                      location="login.asp?lmode=2";
                    }
                  }
                  else
                  {
                    $pc("#PwdLoader").html('<img src="<%=pcf_getImagePath("images","pc_icon_error_small.png")%>" align="absmiddle"> '+data);
                    var callbackPwd=function (){setTimeout(function(){$pc("#PwdLoader").hide();},1000);}
                    $pc("#PwdLoader").effect('pulsate',{},500,callbackPwd);
                  }
                  }
                }
              });
              return(false);
            }
            return(false);
          });
          <%end if%>


        });
        </script>
    </div>
    <%if (Session("CustomerGuest")="0") AND (Session("idCustomer")>"0") then%>
    <div class="pcShowContent">
      <div class="pcSpacer">&nbsp;</div>
    </div>
    <div class="pcFormButtons">   
     	<a class="pcButton pcButtonBack" href="custViewPast.asp">
      	<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
      	<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
			</a>
    </div>
    <%end if%>
    
    
    
    
    
  
    </div>
  </div>
</div>
				</div>
		    </div>
    </section>	
    <!-- /Section: Welcome -->

<!--#include file="footer_wrapper.asp"-->
