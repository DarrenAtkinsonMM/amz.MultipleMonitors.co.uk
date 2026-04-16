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
session("redirectPage")="gwGlobalPay.asp"  'ALTER

'======================================================================================
'// DO NOT ALTER BELOW THIS LINE	
'======================================================================================
'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************

': Declare and Retrieve Customer's IP Address
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("RGPOTE_ADDR")
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

': Open Connection to the DB
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
	query="SELECT pcPay_GP_MerchantID,pcPay_GP_MerchantPassword,pcPay_GP_TransType,pcPay_GP_cardTypes,pcPay_GP_CVC,pcPay_GP_TestMode FROM pcPay_GlobalPay Where pcPay_GP_ID=1;"
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
	pcPay_GP_MerchantID=rs("pcPay_GP_MerchantID")
	pcPay_GP_MerchantID=enDeCrypt(pcPay_GP_MerchantID, scCrypPass)
	pcPay_GP_MerchantPassword=rs("pcPay_GP_MerchantPassword")
	pcPay_GP_MerchantPassword=enDeCrypt(pcPay_GP_MerchantPassword, scCrypPass) 
	pcPay_GP_TransType = rs("pcPay_GP_TransType")
	pcPay_GP_TestMode = rs("pcPay_GP_TestMode")
	pcPay_GP_cardTypes = rs("pcPay_GP_cardTypes")
	pcPay_GP_CVC = rs("pcPay_GP_CVC")
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

	Dim objXMLHTTP,RequestData_QueryString,  xml, XmlSend, strStatus
	
			
	if pcPay_GP_CVC = "1" and (not isNumeric(request.form("CVV")) or  len(request.form("CVV")) < 3 ) Then

        call closeDb()
        Session("message") = "Please supply a Security Code."
        Session("backbuttonURL") = tempURL & "?psslurl=gwGlobalPay.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
        response.redirect "msgb.asp?back=1"		

	End if 

	'Send the transaction info as part of the querystring
	RequestData_QueryString	 = "GlobalUserName=" & server.URLEncode(pcPay_GP_MerchantID)  & _
	"&GlobalPassword="& server.URLEncode(pcPay_GP_MerchantPassword) & _
	"&TransType="& server.URLEncode(pcPay_GP_TransType) & _ 
	"&InvNum=" &	server.URLEncode(session("GWOrderID")) & _
	"&Amount=" & server.URLEncode(pcBillingTotal) & _		
	"&CardNum=" & server.URLEncode(Request.Form( "CardNumber" )) & _
	"&ExpDate=" & server.URLEncode(Request.Form( "expMonth" )& Request.Form( "expYear" )) & _
	"&NameOnCard="& server.URLEncode(pcBillingFirstName & " " & pcBillingLastName) & _
	"&Street="& server.URLEncode(pcBillingAddress) & _
	"&Zip=" & server.URLEncode(pcBillingPostalCode)	& _
	"&CVNum=" & server.URLEncode(request.form("CVV")) & _
	"&MagData=&PNRef=&ExtData="
		   
	' determine where what url to send to 
	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	if pcPay_GP_TestMode=1 then
		xmlSend = "https://certapia.globalpay.com/GlobalPay/transact.asmx/ProcessCreditCard?"& RequestData_QueryString &""
	else
		xmlSend = "https://certapia.globalpay.com/GlobalPay/transact.asmx/ProcessCreditCard?"& RequestData_QueryString & ""
	end if
	'Send the request to the GlobalPay processor.
	xml.open "GET", xmlSend , false
	xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
	xml.send()			
	if err.number<>0 then
		pcResultErrorMsg = err.description
	end if
	
	strStatus = xml.Status
	
	if strStatus = 200 then 	
		'store the response
		strRetVal = xml.responseText
		
		strRetVal = replace(strRetVal,"&lt;", "<")
		strRetVal = replace(strRetVal,"&gt;", ">")
		Response.write "<PRE>" & strRetVal &"</Pre><BR>"
		Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")    
		xmlDoc.async = False
		If xmlDoc.loadXML(strRetVal) Then
			' Get the results
			 pcResultResponseCode = xmlDoc.documentElement.selectSingleNode("Result").Text
			 pcResultResponseMess  = xmlDoc.documentElement.selectSingleNode("RespMSG").Text					
			 pcResultApprovalCode = xmlDoc.documentElement.selectSingleNode("AuthCode").Text
			 pcResultTransRefNumber = xmlDoc.documentElement.selectSingleNode("ExtData/ReceiptData/Trans_Id").Text
			
		Else
	
			'//ERROR
            call closeDb()
            Session("message") = "An undefined processor error occurred during your transaction and your transaction was not approved."
            Session("backbuttonURL") = tempURL & "?psslurl=gwGlobalPay.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
            response.redirect "msgb.asp?back=1"
            
		End If
		If pcResultResponseCode = "0"  then
			session("GWAuthCode")=pcResultApprovalCode
			session("GWTransId")=pcResultTransRefNumber
			response.redirect "gwReturn.asp?s=true&gw=GlobalPay"
		Else				
			pcResultErrorMsg = pcResultResponseMess
			if pcResultErrorMsg="" then
			  pcResultErrorMsg="An undefined error occurred during your transaction and your transaction was not approved.<BR>"              	
			 end if
            call closeDb()
            Session("message") = pcResultErrorMsg
            Session("backbuttonURL") = tempURL & "?psslurl=gwGlobalPay.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
            response.redirect "msgb.asp?back=1"	
		End if
	Else		  
		if pcResultErrorMsg="" then
			pcResultErrorMsg="An undefined processor error occurred during your transaction and your transaction was not approved.<BR>"					
		end if
        call closeDb()
        Session("message") = pcResultErrorMsg
        Session("backbuttonURL") = tempURL & "?psslurl=gwGlobalPay.asp&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
        response.redirect "msgb.asp?back=1"			 
	End if 

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
                        ArryCardTypes=split(pcPay_GP_CardTypes,", ")
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
					If pcPay_GP_CVC="1" Then %>
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
