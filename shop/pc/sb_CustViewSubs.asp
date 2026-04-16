<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
'Check to see if ARB has been turned off by admin, then display message
If scSBStatus="0" then
	response.redirect "msg.asp?message=212"
End If 

if scSSL = "1" then
	If (Request.ServerVariables("HTTPS") = "off") Then
		Dim xredir__, xqstr__
		xredir__ = "https://" & Request.ServerVariables("SERVER_NAME") & _
				   Request.ServerVariables("SCRIPT_NAME")
		xqstr__ = Request.ServerVariables("QUERY_STRING")
		if xqstr__ <> "" Then xredir__ = xredir__ & "?" & xqstr__
		Response.redirect xredir__
	End if
end if

'SB S
query="SELECT orders.idOrder, orders.orderDate, orders.total, orders.ord_OrderName, ProductsOrdered.idProductOrdered,ProductsOrdered.UnitPrice,ProductsOrdered.quantity, ProductsOrdered.pcSubscription_ID, ProductsOrdered.pcPO_SubAmount, ProductsOrdered.pcPO_SubActive, ProductsOrdered.pcPO_IsTrial, ProductsOrdered.pcPO_SubTrialAmount, ProductsOrdered.pcPO_SubStartDate, ProductsOrdered.pcPO_SubType FROM orders, productsordered WHERE orders.idCustomer=" & Session("idcustomer") &" AND ((orders.OrderStatus>1) OR (ProductsOrdered.pcPO_SubActive=3))  And orders.idOrder = ProductsOrdered.idOrder and ProductsOrdered.pcSubscription_ID >0  ORDER BY ProductsOrdered.idProductOrdered DESC"
'SB E
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)
if err.number <> 0 then
    call LogErrorToDatabase()
    set rstemp = Nothing
    call closeDb()
    response.redirect "techErr.asp?err="&pcStrCustRefID
end If
if rstemp.eof then
	set rstemp=nothing
	call closeDb()
 	response.redirect "msg.asp?message=34"
end if
%> 

<!--#include file="header_wrapper.asp"-->
<div id="pcMain">
	<div class="pcMainContent">
		<h1>
		<%
		if session("pcStrCustName") <> "" then
			response.write(session("pcStrCustName") & " -  Subscriptions")
			else
			response.write "Subscriptions"
		end if
		%>
		</h1>
		<div class="pcTable">   
		<%if ((request("u")="1") OR (request("u")="2")) AND (Session("SBEditOrder")<>"") AND (Session("SBEditOrderID")<>"") then%>
		<div class="pcTableRowFull">
				<%if (request("u")="1") then%>
				<div class="pcSuccessMessage">
					<%response.write dictLanguage.Item(Session("language")&"_SB_35A")%><%=Session("SBEditOrderID")%><%response.write dictLanguage.Item(Session("language")&"_SB_35B")%>
				</div>
				<%end if%>
				<%if (request("u")="2") then%>
				<div class="pcInfoMessage">
					<%response.write dictLanguage.Item(Session("language")&"_SB_36")%>
				</div>
				<%end if%>
		</div>
		<%Session("SBEditOrder")=""
		Session("SBEditOrderID")=""

		'// CLEAR SESSIONS
		session("reqCardNumber")=""
		session("reqExpMonth")=""
		session("reqExpYear")=""
		session("reqCardType")=""
		session("reqCVV")=""		
		session("x_bank_acct_name") = ""
		session("x_bank_aba_code") = ""
		session("x_bank_acct_num") =  ""
		session("x_bank_acct_type") = ""
		session("x_customer_organization_type") = ""
		session("x_bank_name") = ""
		session("x_customer_tax_id") = "" 					
		session("x_drivers_license_num") = ""
		session("x_drivers_license_state") =  ""
		session("x_drivers_license_dob") = ""				
		session("pcIsSubscription") = ""
		session("pcIsSubTrial") = ""
		session("pcAgreeAll") = ""
		
		' clear cart data
		dim pcCartArray2(100,45)
		Session("pcCartSession")=pcCartArray2
		Session("pcCartIndex")=Cint(0)
		session("iOrderTotal")=""
		session("continueRef")=""
		session("pcSFCartRewards")=Cint(0)
		session("pcSFUseRewards")=Cint(0)
		session("IDRefer")=""
		session("specialdiscount")=""
		session("EPN_idOrder")=""
		session("pc_pidOrder")=""
		session("GWAuthCode")=""
		session("GWTransId")=""
		session("Entered-" & session("GWPaymentId"))=""
		session("admin-" & session("GWPaymentId") & "-pCardType")=""
		session("admin-" & session("GWPaymentId") & "-pCardNumber")=""
		session("admin-" & session("GWPaymentId") & "-expMonth")=""
		session("admin-" & session("GWPaymentId") & "-expYear")=""
		session("GWPaymentId")=""
		session("GWTransType")=""
		session("GWOrderId")=""
		session("GWSessionID")=""
		session("GWOrderDone")=""
		session("idGWSubmit")=""
		session("idGWSubmit2")=""
		session("idGWSubmit3")=""
		session("Gateway")=""
		session("SaveOrder")=""
		session("RefRewardPointsTest")=""
		'GGG Add-on start
		session("Cust_BuyGift")=""
		session("Cust_IDEvent")=""
		'GGG Add-on end
		
		session("idOrder")=""
		session("idOrderConfirm")=""
		session("GWOrderId")=""
		
		' clear cart data
		if len(session("pcSFIdDbSession"))>0 then
			query="DELETE FROM pcCustomerSessions WHERE idDbSession="&session("pcSFIdDbSession")&" AND randomKey="&session("pcSFRandomKey")&" AND idCustomer="&session("idCustomer")&";"
			set rsQ=conntemp.execute(query)
			set rsQ=nothing
		end if

		session("pcSFIdDbSession")=""
		session("pcSFRandomKey")=""
		session("iOrderTotal")=""
		session("pcSFCartRewards")=Cint(0)
		session("pcSFUseRewards")=Cint(0)
		session("IDRefer")=""
		session("specialdiscount")=""
		Session("ContinueRef")=""
		session("TF1")=""
		session("DF1")=""
		session("shippingFullName")=""
		session("shippingCompany")=""
		session("shippingAddress")=""
		session("shippingAddress2")=""
		session("shippingStateCode")=""
		session("shippingState")=""
		session("shippingZip")=""
		session("shippingPhone")=""
		session("shippingCity")=""
		session("shippingCountryCode")=""
		session("DCODE")=""
		session("idOrderSaved")=""
		session("ExpressCheckoutPayment")=""
		session("GWOrderDone")=""
		session("redirectPage")=""
		Session("SFStrRedirectUrl")=""
		session("idGWSubmit")=""
		session("idGWSubmit2")=""
		session("idGWSubmit3")=""
		session("Gateway")=""
		session("SaveOrder")=""
		Session("pcPromoSession")=""
		Session("pcPromoIndex")=""
		session("Entered-" & session("GWPaymentId"))=""
		session("SF_DiscountTotal")=""
		session("SF_RewardPointTotal")=""
		
		end if%>
		<div class="pcTableHeader">
			<div style="width:15%"><%response.write dictLanguage.Item(Session("language")&"_SB_10")%></div>
			<div style="width:40%"><%response.write dictLanguage.Item(Session("language")&"_SB_11")%></div>
			<div style="width:45%"><%response.write dictLanguage.Item(Session("language")&"_SB_12")%></div>
		</div>
		<div class="pcTableRowFull">
			<div class="pcSpacer">&nbsp;</div>
		</div>
				<%
				SaveSBGuid=""
				do while not rstemp.eof
					idorder = rstemp("idOrder")
					idProductOrdered = rstemp("idProductOrdered")
					pSubUnitPrice = rstemp("unitPrice")
					pSubQty = rstemp("quantity")
					pSubPrice = rstemp("pcPO_SubAmount")
					pSubTrial = rstemp("pcPO_IsTrial")
					pSubTrialAmount = rstemp("pcPO_SubTrialAmount")						 
					pSubStartDate = rstemp("pcPO_SubStartDate")
					pSubActive =rstemp("pcPO_SubActive") 
					
					pSubType=rstemp("pcPO_SubType")					

					'// Obtain Status
					Dim pvc_Status
					Select Case pSubActive
					Case "1":
						pcv_Status = "Active"
					Case "2":
						pcv_Status = "Pending"
					Case "3":
						pcv_Status = "Edited"
					Case Else
						pcv_Status = "<font color='#ff0000'>Not-Active</font>"
					End Select 
					
					'// Obtain GUID and Email
					Dim pcv_strCustEmail, pcv_strGUID
					pcv_strCustEmail=""
					pcv_strGUID=""

					query = "SELECT customers.email, SB_Orders.SB_GUID FROM SB_Orders "
					query = query & "INNER JOIN ( orders INNER JOIN customers On orders.idcustomer = customers.idcustomer ) ON orders.idorder = SB_Orders.idorder "
					query = query & "WHERE orders.idorder = " & idorder

					set rsSB=Server.CreateObject("ADODB.Recordset")
					set rsSB=conntemp.execute(query)
					if err.number <> 0 then
                        call LogErrorToDatabase()
                        set rsSB = Nothing
                        call closeDb()
                        response.redirect "techErr.asp?err="&pcStrCustRefID
					end If
					if NOT rsSB.eof then
						   pcv_strCustEmail = rsSb("email")
						   pcv_strGUID = rsSb("SB_GUID")
					end if  
					set rsSB=nothing 

					If len(pcv_strGUID)>0 Then

						query="SELECT Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess FROM SB_Settings;"
						set rsAPI=connTemp.execute(query)
						if not rsAPI.eof then
							Setting_APIUser=rsAPI("Setting_APIUser")
							Setting_APIPassword=enDeCrypt(rsAPI("Setting_APIPassword"), scCrypPass)
							Setting_APIKey=enDeCrypt(rsAPI("Setting_APIKey"), scCrypPass)
						end if
						set rsAPI=nothing
						
						
						Set objSB = NEW pcARBClass
						
						objSB.GUID = pcv_strGUID
						If scSBLanguageCode<>"" Then
							objSB.CartLanguageCode = scSBLanguageCode
						Else
							objSB.CartLanguageCode = "en-EN"
						End If
			
						Dim result

						result = objSB.GetSubscriptionDetailsRequest(Setting_APIUser, Setting_APIPassword, Setting_APIKey)

						If SB_ErrMsg="" Then
							
							pcv_strGUID = objSB.pcf_GetNode(result, "Guid", "//GetSubscriptionDetailsResponse/Subscription/Identifiers")
							pcv_strStatus = objSB.pcf_GetNode(result, "Status", "//GetSubscriptionDetailsResponse/Subscription/Identifiers")
							pcv_strBalance = objSB.pcf_GetNode(result, "BalanceTotal", "//GetSubscriptionDetailsResponse/Subscription")
							pcv_strBillingAgreement = objSB.pcf_GetNode(result, "Terms", "//GetSubscriptionDetailsResponse/Subscription")

							if pcv_strBalance="" then
								pcv_strBalance = 0
							end if
							
							'// Only display one and latest Order of each SB Guid
							If (len(pcv_strGUID)>0) AND (Instr(SaveSBGuid,"|" & pcv_strGUID & "|")=0) Then
							SaveSBGuid=SaveSBGuid & "|" & pcv_strGUID & "|"
							%>
							<div class="pcTableRow">
								<div style="width:15%">
								   <a href="CustviewPastD.asp?idOrder=<%response.write (scpre+int(IdOrder))%>"><%response.write (scpre+int(IdOrder))%></a>
								</div>
								<div style="width:40%">
									<a href="JavaScript:openManageSubscription('sb_CustSubDetails.asp?ID=<%=idorder%>&GUID=<%=pcv_strGUID%>')"><%=pcv_strGUID%></a>
									<br />
									<i><%response.write dictLanguage.Item(Session("language")&"_SB_31")%>: <%=pcv_strStatus%><%if pSubActive="3" then%>&nbsp;(<%=pcv_Status%>)<%end if%></i>
								</div>
								<div style="width:45%">
									<%=pcv_strBillingAgreement%>
								</div>
							</div>
							<div class="pcTableRow">
								<div>
								
										<% If pcv_strBalance>0 AND scSSL = "1" AND lcase(pcv_strStatus)="active" Then %>
										
											<div align="center" class="pcErrorMessage">
												<%response.write dictLanguage.Item(Session("language")&"_SB_16")%>&nbsp; 
												<a href="JavaScript:openManageSubscription('sb_CustOneTimePayment.asp?ID=<%=idorder%>&GUID=<%=pcv_strGUID%>')"><%response.write dictLanguage.Item(Session("language")&"_SB_17")%></a> 
											</div>
											
											
										<% End If %>
								
										<div align="left" class="pcSmallText">
										
											<a href="<%=gv_RootURL%>/CustomerCenter/AutoLogin.asp?ID=<%=pcv_strGUID%>&Email=<%=pcv_strCustEmail%>&mode=history" target="_blank"><%response.write dictLanguage.Item(Session("language")&"_SB_5")%></a>
											&nbsp;|&nbsp;
											<a href="<%=gv_RootURL%>/CustomerCenter/AutoLogin.asp?ID=<%=pcv_strGUID%>&Email=<%=pcv_strCustEmail%>&mode=details" target="_blank"><%response.write dictLanguage.Item(Session("language")&"_SB_4")%></a>
		
											<% if (lcase(pcv_strStatus)="active") OR (lcase(pcv_strStatus)="edited") Then %>
										
												  &nbsp;|&nbsp;
	
												  <!--<a href="JavaScript:openManageSubscription('sb_CustUpdatePayment.asp?ID=<%=idorder%>')"><%response.write dictLanguage.Item(Session("language")&"_SB_14")%></a>
												  &nbsp;|&nbsp; -->
												  
												  <a href="sb_EditOrder.asp?ID=<%=idorder%>&GUID=<%=pcv_strGUID%>"><%response.write dictLanguage.Item(Session("language")&"_SB_33")%></a>
												  &nbsp;|&nbsp;
	
												  <a href="JavaScript:openManageSubscription('sb_CustCancelSub.asp?ID=<%=idorder%>&GUID=<%=pcv_strGUID%>')"><%response.write dictLanguage.Item(Session("language")&"_SB_8")%></a>
												  &nbsp;
				   
				  
										   <% elseif (lcase(pcv_strStatus)="pending") then %>
												  
												  <a href="<%=gv_RootURL%>/CustomerCenter/AutoLogin.asp?ID=<%=pcv_strGUID%>&Email=<%=pcv_strCustEmail%>&mode=" target="_blank"><%response.write dictLanguage.Item(Session("language")&"_SB_9")%></a>&nbsp;
												 
										   <%else%>
												  
												  
												  
										   <%end if %>
										</div>
								
								</div>
							</div>
							<div class="pcTableRowFull"><hr></div>
							<%
							End If
						End If
						
					End If '// If len(pcv_strGUID)>0 Then
						
					rstemp.movenext
			  	loop
				%>
			<% 
			set rstemp = nothing
			%>
		<div class="pcTableRowFull">
			<div class="pcSpacer">&nbsp;</div>
		</div>
		<div class="pcTableRowFull">
			<div class="pcSpacer">&nbsp;</div>
		</div> 
		<div class="pcTableRowFull">
		<a class="pcButton pcButtonBack" href="custpref.asp">
			<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
			<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
        </a>
		</div>
</div>
</div>
</div>
<div class="modal fade" id="SBDialog" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
   <div class="modal-dialog modal-lg">
	  <div id="modal-content" class="modal-content">
	  </div>
   </div>
</div>
<script type=text/javascript>
function openMSB(url)
{
	$pc(document.body).on('hidden.bs.modal', function () {
		$pc('#SBDialog').removeData('bs.modal');
		openManageSubscription(url);
	});
	$pc('#SBDialog').modal('hide');
}
function openManageSubscription(url) {
	$pc('#SBDialog').appendTo('body').modal({
		show: false,
		remote: url
	});
	$pc('#SBDialog').on('loaded.bs.modal', function (e) {
		$pc('#SBDialog').modal('show');
	})
}

</script>
<!--#include file="footer_wrapper.asp"-->
