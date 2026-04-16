<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'======================================================================================
'// Set redirect page
'======================================================================================
' The redirect page tells the form where to post the payment information. Most of the 
' time you will redirect the form back to this page.
'======================================================================================
session("redirectPage")="gwEmoney.asp"  'ALTER

'======================================================================================
'// DO NOT ALTER BELOW THIS LINE	
'======================================================================================
'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************

': Declare and Retrieve Customer's IP Address
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
': End Declare and Retrieve Customer's IP Address	

': Declare URL path to gwSubmit.asp	
Dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If
': End Declare URL path to gwSubmit.asp

': Get Order ID and Set to session
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if
': End Get Order ID
	
': Get customer and order data from the database for this order	
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<%
': End Get customer and order data


': Reset customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer
': End Reset customer session

 'DELETE FOR HARD CODED VARS
'======================================================================================
'// DO NOT ALTER ABOVE THIS LINE	
'======================================================================================

'======================================================================================
'// Retrieve ALL required gateway specific data from database or hard-code the variables
'// Alter the query string replacing "fields", "table" and "idField" to reference your 
'// new database table.
'//
'// If you are going to hard-code these variables, delete all lines that are
'// commented as 'DELETE FOR HARD CODED VARS' and then set variables starting
'// at line 101 below.
'======================================================================================
	query="SELECT pcPay_EM_MerchantID,pcPay_EM_MerchantPassword,pcPay_EM_cardTypes,pcPay_EM_CVC,pcPay_EM_TestMode FROM pcPay_Emoney Where pcPay_EM_ID=1;"
'ALTER :: DELETE FOR HARD CODED VARS
'======================================================================================
'// End custom query
'======================================================================================

': Create recordset and execute query
set rs=server.CreateObject("ADODB.RecordSet") 'DELETE FOR HARD CODED VARS
set rs=connTemp.execute(query) 'DELETE FOR HARD CODED VARS

': Capture any errors
if err.number<>0 then 'DELETE FOR HARD CODED VARS
	call LogErrorToDatabase() 'DELETE FOR HARD CODED VARS
	set rs=nothing 'DELETE FOR HARD CODED VARS
	call closedb() 'DELETE FOR HARD CODED VARS
	response.redirect "techErr.asp?err="&pcStrCustRefID 'DELETE FOR HARD CODED VARS
end if 'DELETE FOR HARD CODED VARS

'======================================================================================
'// Set gateway specific variables - These can be your "hard coded variables" or 
'// Variables retrieved from the database.
'======================================================================================
	pcPay_EM_MerchantID=rs("pcPay_EM_MerchantID")
    If len(pcPay_EM_MerchantID)>0 Then
	    pcPay_EM_MerchantID=enDeCrypt(pcPay_EM_MerchantID, scCrypPass)
    End If
	pcPay_EM_MerchantPassword=rs("pcPay_EM_MerchantPassword")
    If len(pcPay_EM_MerchantPassword)>0 Then
	    pcPay_EM_MerchantPassword=enDeCrypt(pcPay_EM_MerchantPassword, scCrypPass) 
    End If
	pcPay_EM_TestMode = rs("pcPay_EM_TestMode")
	pcPay_EM_cardTypes = rs("pcPay_EM_cardTypes")
	pcPay_EM_CVC = rs("pcPay_EM_CVC")
'======================================================================================
'// End gateway specific variables
'======================================================================================

': Clear recordset and close db connection
set rs=nothing 'DELETE FOR HARD CODED VARS

'======================================================================================
'// If you are posting back to this page from the gateway form, all actions will happen 
'// here. 
'======================================================================================
if request("PaymentSubmitted")="Go" then
  

	'*************************************************************************************
	'// This is where you would post and retrieve info to and from the gateway
	'// START below this line
	'*************************************************************************************
	
		    Dim objXMLHTTP, xml
		
            'Send the transaction info as part of the querystring
	        RequestData_QueryString	 = "A=AUTHORIZEONLINE" & _
			"&CID=" & server.URLEncode(pcPay_EM_MerchantID)  & _
			"&PWD="& server.URLEncode(pcPay_EM_MerchantPassword) & _
			"&PID1="& server.URLEncode(session("GWOrderID")) & _
			"&TA=" & server.URLEncode(pcBillingTotal) 			
			if pcPay_EM_CVC = 1 Then 
				RequestData_QueryString = RequestData_QueryString &"&CVV=" & server.URLEncode(request.form("CVV"))
			End if 
			
		    'Send the transaction info as part of the post
			RequestData_Post = "AN=" & server.URLEncode(Request.Form( "CardNumber" )) & _
			"&ED=" & server.URLEncode(Request.Form( "expMonth" )& Request.Form( "expYear" )) & _
			"&NOC="& server.URLEncode(pcBillingFirstName & " " & pcBillingLastName) & _
			"&AS1="& server.URLEncode(pcBillingAddress) & _
			"&AZ=" & server.URLEncode(pcBillingPostalCode)
			
			' determine where what url to send to 
			set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
			if pcPay_EM_TestMode=1 then
				xmlSend = "https://demo.emoney2k.com/ECOM3NX/main.asp?"& RequestData_QueryString &""
			else
				xmlSend = "https://pos.emoney2k.com/ECOM3NX/main.asp?"& RequestData_QueryString & ""
			end if
			'Send the request to the eMoney processor.
			xml.open "POST", xmlSend , false
			xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
			xml.send(RequestData_Post)			
			if err.number<>0 then
				pcResultErrorMsg = err.description
				'response.end
			end if
			strStatus = xml.Status
			if strStatus = 200 then 	
				'store the response
				strRetVal = xml.responseText
				'Response.write "<PRE>" & strRetVal &"</Pre>"
				'response.end 
				
				Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")    
				xmlDoc.async = False
				If xmlDoc.loadXML(strRetVal) Then
					' Get the results
					pcResultAppAmt = xmlDoc.documentElement.selectSingleNode("/responsemessage/va").Text
					pcResultApproval = xmlDoc.documentElement.selectSingleNode("/responsemessage/ac").Text
					pcResultResponseMess = xmlDoc.documentElement.selectSingleNode("/responsemessage/rt").Text
					pcResultResponseCode = xmlDoc.documentElement.selectSingleNode("/responsemessage/rc").Text
					pcResultTransRefNumber = xmlDoc.documentElement.selectSingleNode("/responsemessage/tsn").Text
					
				Else
                
					'//ERROR
					pcResultErrorMsg="An undefined processor error occurred during your transaction and your transaction was not approved.<BR>"										
                    call closeDb()
                    Session("message") = pcResultErrorMsg
                    Session("backbuttonURL") = tempURL & "?psslurl=gwemoney.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                    response.redirect "msgb.asp?back=1"
                    
				End If
				If pcResultResponseCode = "00" or pcResultResponseCode = "85" then
				 	session("GWAuthCode")=pcResultApproval
					session("GWTransId")=pcResultTransRefNumber
					response.redirect "gwReturn.asp?s=true&gw=eMoney"
				Else				
					pcResultErrorMsg = pcResultResponseMess
					if pcResultErrorMsg="" then
					  pcResultErrorMsg="An undefined error occurred during your transaction and your transaction was not approved.<BR>"              	
					end if
                    call closeDb()
                    Session("message") = pcResultErrorMsg
                    Session("backbuttonURL") = tempURL & "?psslurl=gwemoney.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
                    response.redirect "msgb.asp?back=1"	
				End if
			Else
		 
			if pcResultErrorMsg="" then
				pcResultErrorMsg="An undefined processor error occurred during your transaction and your transaction was not approved.<BR>"					
			end if
            
            call closeDb()
            Session("message") = pcResultErrorMsg
            Session("backbuttonURL") = tempURL & "?psslurl=gwemoney.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
            response.redirect "msgb.asp?back=1"	
		 
			End if 
			Dim pcv_SuccessURL
			If scSSL="" OR scSSL="0" Then
				pcv_SuccessURL=replace((scStoreURL&"/"&scPcFolder&"/pc/pcPay_TD_Receipt.asp"),"//","/")
				pcv_SuccessURL=replace(pcv_SuccessURL,"https:/","https://")
				pcv_SuccessURL=replace(pcv_SuccessURL,"http:/","http://") 
			Else
				pcv_SuccessURL=replace((scSslURL&"/"&scPcFolder&"/pc/pcPay_TD_Receipt.asp"),"//","/")
				pcv_SuccessURL=replace(pcv_SuccessURL,"https:/","https://")
				pcv_SuccessURL=replace(pcv_SuccessURL,"http:/","http://")
			End If

	

	'*************************************************************************************
	' END
	'*************************************************************************************
	
end if 
'======================================================================================
'// End post back 
'======================================================================================


'======================================================================================
'// Show customer the payment form 
'======================================================================================
%>
<div id="pcMain">
	<div class="pcMainContent">

        <form method="POST" action="<%=session("redirectPage")%>" name="PaymentForm" class="pcForms">
			<input type="hidden" name="PaymentSubmitted" value="Go">

            <% If msg<>"" Then %>
                <div class="pcErrorMessage"><%=msg%></div>
            <% End If %>
                    
            <% call pcs_showBillingAddress %>

            <div class="pcFormItem"> 
                <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_12")%></div>
                <div class="pcFormField">
                    <select name="CardType">
                        <% dim ArryCardTypes, strCardType, j
                        ArryCardTypes=split(pcPay_EM_CardTypes,", ")
                        for j=0 to ubound(ArryCardTypes) 
                            strCardType=ArryCardTypes(j) 
                            select case strCardType
                                case "VISA"
                                    response.write "<option value='VISA'>VISA</option>"
                                case "MAST"
                                    response.write "<option value='MAST'>Master Card</option>"
                                case "AMER"
                                    response.write "<option value='AMER'>American Express</option>"
                                case "DISC"
                                    response.write "<option value='DISC'>Discover Card</option>"
                                case "DINE"
                                    response.write "<option value='DINE'>Diners Club</option>"
                                case "JCB"
                                    response.write "<option value='JCB'>JCB</option>"
                            end select
                        next %>
                    </select>
                </div>
            </div>
            
            
            <div class="pcFormItem">
                <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_7")%></div>
                <div class="pcFormField"><input type="text" name="CardNumber" value="" autocomplete="off"></div>
            </div>

					<div class="pcFormItem">
						<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_8")%></div>
						<div class="pcFormField"><%=dictLanguage.Item(Session("language")&"_GateWay_9")%> 
							<select name="expMonth">
								<option value="01">1</option>
								<option value="02">2</option>
								<option value="03">3</option>
								<option value="04">4</option>
								<option value="05">5</option>
								<option value="06">6</option>
								<option value="07">7</option>
								<option value="08">8</option>
								<option value="09">9</option>
								<option value="10">10</option>
								<option value="11">11</option>
								<option value="12">12</option>
							</select>
							<% dtCurYear=Year(date()) %>
							&nbsp;<%=dictLanguage.Item(Session("language")&"_GateWay_10")%> 
							<select name="expYear">
								<option value="<%=right(dtCurYear,2)%>" selected><%=dtCurYear%></option>
								<option value="<%=right(dtCurYear+1,2)%>"><%=dtCurYear+1%></option>
								<option value="<%=right(dtCurYear+2,2)%>"><%=dtCurYear+2%></option>
								<option value="<%=right(dtCurYear+3,2)%>"><%=dtCurYear+3%></option>
								<option value="<%=right(dtCurYear+4,2)%>"><%=dtCurYear+4%></option>
								<option value="<%=right(dtCurYear+5,2)%>"><%=dtCurYear+5%></option>
								<option value="<%=right(dtCurYear+6,2)%>"><%=dtCurYear+6%></option>
								<option value="<%=right(dtCurYear+7,2)%>"><%=dtCurYear+7%></option>
								<option value="<%=right(dtCurYear+8,2)%>"><%=dtCurYear+8%></option>
								<option value="<%=right(dtCurYear+9,2)%>"><%=dtCurYear+9%></option>
								<option value="<%=right(dtCurYear+10,2)%>"><%=dtCurYear+10%></option>
							</select>
						</div>
					</div>
                    
					<% 
					'======================================================================================
					'// If your gateway supports Credit Card Security Code (such as CVV and CVV2), create
					'// a variable for it and then show the row below.
					'// NOTE :: If no Security Code support exists, delete the table row below
					'======================================================================================
					If pcPay_EM_CVC="1" Then %>
                <div class="pcFormItem">
                    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></div>
                    <div class="pcFormField"><input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4"></div>
                </div> 
                <div class="pcFormItem">
                    <div class="pcFormLabel">&nbsp;</div>
                    <div class="pcFormField"><img src="<%=pcf_getImagePath("images","CVC.gif")%>" alt="cvc code" width="212" height="155"></div>
                </div>
					<% end if
					'======================================================================================
					'// End Security Code support
					'// NOTE :: If no Security Code support exists, delete the table row above
					'======================================================================================
				 	%>
                    
            <div class="pcFormItem"> 
			    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_4")%></div>
                <div class="pcFormField"><%= scCurSign & money(pcBillingTotal)%></div> 
            </div>
					
            <div class="pcFormButtons">
                <!--#include file="inc_gatewayButtons.asp"-->
            </div>
        </form>
    </div>
</div>
<% 
'======================================================================================
'// End Show customer the payment form 
'======================================================================================
%>
<!--#include file="footer_wrapper.asp"-->
