<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="opc_contentType.asp" -->
<!--#include file="inc_sb.asp"-->
<% 
HaveSecurity=0
if session("idCustomer")=0 OR session("idCustomer")="" then
	HaveSecurity=1
end if

'Check to see if ARB has been turned off by admin, then display message
If scSBStatus="0" then
	response.redirect "msg.asp?message=212"
End If 

Call SetContentType()

IF HaveSecurity=0 THEN

	qry_GUID = getUserInput(Request("GUID"),0)
	qry_ID = getUserInput(Request("ID"),0)
	if not validNum(qry_ID) then
	   qry_ID=0
	end if

	if request("action")="add" then

		'// Not needed

	end if
	
END IF
%>
<div class="modal-header">
    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
    <h3 class="modal-title" id="pcDialogTitle"><%response.write dictLanguage.Item(Session("language")&"_SB_21")%></h3>
</div>
<div class="modal-body">
<form method="post" name="BillingForm" id="BillingForm" action="sb_CustOneTimePayment.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<input name="ID" type="hidden" value="<%=qry_ID%>">
<input name="GUID" type="hidden" value="<%=qry_GUID%>">
	<div class="pcTable">
	<%IF HaveSecurity=1 THEN%>
        <div class="pcTableRowFull">
			<div class="pcErrorMessage"><%response.write dictLanguage.Item(Session("language")&"_SB_20")%></div>
		</div>
    <%ELSE%>
    
        <%IF UpdateSuccess="1" THEN%>
            <div class="pcTableRow">
				<div class="pcSuccessMessage"><%response.write dictLanguage.Item(Session("language")&"_SB_22")%></div>
            </div>        
		<%ELSE%>            
				<% If SB_ErrMsg <> "" Then %>
                <div class="pcTableRowFull">
					<div class="pcErrorMessage"><%=SB_ErrMsg%></div>
                </div>
				<% End If %>
				<%
				query="SELECT Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess FROM SB_Settings;"
				set rsAPI=connTemp.execute(query)
				if not rsAPI.eof then
					Setting_APIUser=rsAPI("Setting_APIUser")
					Setting_APIPassword=enDeCrypt(rsAPI("Setting_APIPassword"), scCrypPass)
					Setting_APIKey=enDeCrypt(rsAPI("Setting_APIKey"), scCrypPass)
				end if
				set rsAPI=nothing                  
				
				Set objSB = NEW pcARBClass                  
				objSB.GUID = qry_GUID
				If scSBLanguageCode<>"" Then
					objSB.CartLanguageCode = scSBLanguageCode
				Else
					objSB.CartLanguageCode = "en-EN"
				End If
			  
				result = objSB.GetSubscriptionDetailsRequest(Setting_APIUser, Setting_APIPassword, Setting_APIKey)
				
				If SB_ErrMsg="" Then 					
					pcv_strGUID = objSB.pcf_GetNode(result, "Guid", "//GetSubscriptionDetailsResponse/Subscription/Identifiers")
					pcv_strStatus = objSB.pcf_GetNode(result, "Status", "//GetSubscriptionDetailsResponse/Subscription/Identifiers")
					pcv_strBillingEmail = objSB.pcf_GetNode(result, "Email", "//GetSubscriptionDetailsResponse/Subscription/Customer")
					pcv_Description = objSB.pcf_GetNode(result, "Description", "//GetSubscriptionDetailsResponse/Subscription/SubscriptionDetails")
					pcv_NextBillingAmt = objSB.pcf_GetNode(result, "NextBillingAmt", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")
					pcv_NextBillingDate = objSB.pcf_GetNode(result, "NextBillingDate", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")
		
					pcv_PreviousPaymentDate = objSB.pcf_GetNode(result, "LastPaymentDate", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")
					pcv_PreviousPaymentAmount = objSB.pcf_GetNode(result, "LastPaymentAmount", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")
					pcv_StartDate = objSB.pcf_GetNode(result, "StartDate", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")
					pcv_EndDate = objSB.pcf_GetNode(result, "EndDate", "//GetSubscriptionDetailsResponse/Subscription/BillingDetails")			
					If instr(pcv_EndDate,"1/1/1900")>0 Then
						pcv_EndDate=dictLanguage.Item(Session("language")&"_SB_30")
					End If
					pcv_strBalance = objSB.pcf_GetNode(result, "Balance", "//GetSubscriptionDetailsResponse/Subscription/OutstandingBalance")
					pcv_strReason = objSB.pcf_GetNode(result, "Reason", "//GetSubscriptionDetailsResponse/Subscription/OutstandingBalance")                     
				End If
                %>  
                <div class="pcTableRowFull">
                    <div class="pcSpacer">&nbsp;</div>
                </div>
				<div class="pcTableHeader">
					<div style="width:25%;text-align:left;"><%response.write dictLanguage.Item(Session("language")&"_SB_11")%></div>
					<div style="width:25%;text-align:left;"><%response.write dictLanguage.Item(Session("language")&"_SB_23")%></div>
					<div style="width:25%;text-align:left;"><%response.write dictLanguage.Item(Session("language")&"_SB_24")%></div>
					<div style="width:25%;text-align:left;"><%response.write dictLanguage.Item(Session("language")&"_SB_25")%></div>
				</div> 
				<div class="pcTableRow">
					<div style="width:25%;text-align:left; white-space:normal;"><%=pcv_strGUID%></div>
					<div style="width:25%;text-align:left;"><%=pcv_Description%></div>
					<div style="width:25%;text-align:left;"><%=scCurSign & money(pcv_NextBillingAmt)%></div>
					<div style="width:25%;text-align:left;"><%=pcv_NextBillingDate%></div>
				</div>

				<div class="pcTableRowFull">
					<div class="pcSpacer">&nbsp;</div>
				</div>
	
				<div class="pcTableHeader">
					<div style="width:25%;text-align:left;"><%response.write dictLanguage.Item(Session("language")&"_SB_26")%></div>
					<div style="width:25%;text-align:left;"><%response.write dictLanguage.Item(Session("language")&"_SB_27")%></div>
					<div style="width:25%;text-align:left;"><%response.write dictLanguage.Item(Session("language")&"_SB_28")%></div>
					<div style="width:25%;text-align:left;"><%response.write dictLanguage.Item(Session("language")&"_SB_29")%></div>
				</div> 
				<div class="pcTableRow">
					<div style="width:25%;text-align:left;"><%=pcv_PreviousPaymentDate%></div>
					<div style="width:25%;text-align:left;"><%=scCurSign & money(pcv_PreviousPaymentAmount)%></div>
					<div style="width:25%;text-align:left;"><%=pcv_StartDate%></div>
					<div style="width:25%;text-align:left;"><%=pcv_EndDate%></div>
				</div>                
				
				<div class="pcTableRowFull">
					<div class="pcSpacer">&nbsp;</div>
				</div>
				
				<div class="pcTableHeader">
					<div style="width:25%;text-align:left;"><%response.write dictLanguage.Item(Session("language")&"_SB_31")%></div>
					<div style="width:25%;text-align:left;"><%response.write dictLanguage.Item(Session("language")&"_SB_32")%></div>
					<div style="width:25%;text-align:left;">&nbsp;</div>
					<div style="width:25%;text-align:left;">&nbsp;</div>
				</div> 
				<div class="pcTableRow">
					<div style="width:25%;text-align:left;"><%=pcv_strStatus%></div>
					<div style="width:25%;text-align:left;"><%=pcv_strBillingEmail%></div>
					<div style="width:25%;text-align:left;">&nbsp;</div>
					<div style="width:25%;text-align:left;">&nbsp;</div>
				</div>                
				<div class="pcTableRowFull">
					<div class="pcSpacer">&nbsp;</div>
				</div>
        <%END IF%>
    <%END IF%>
	</div>
</form>
</div>
<div class="modal-footer">
    <button class="btn btn-default" data-dismiss="modal" type="button"><%=dictLanguage.Item(Session("language")&"_AddressBook_5")%></button>
</div>
<% call closedb() %>
