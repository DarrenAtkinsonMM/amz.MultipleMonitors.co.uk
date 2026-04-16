<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<%
Dim xmldoc
Function CheckExistTag(tagName)
Dim tmpNode
	Set tmpNode=iRoot.selectSingleNode(tagName)
	If tmpNode is Nothing Then
		CheckExistTag=False
	Else
		CheckExistTag=True
	End if
End Function

' thisd is to clean out nay bad characters for VM XML parcer
function getUserVM_XMLOutPut(input,stringLength)
 dim tempStr

 known_bad= array("*","--")
 if stringLength>0 then
  tempStr	= left(trim(input),stringLength) 
 else
  tempStr	= trim(input)
 end if
 for i=lbound(known_bad) to ubound(known_bad)
 	if (instr(1,tempStr,known_bad(i),vbTextCompare)<>0) then
 		tempStr	= replace(tempStr,known_bad(i),"")
 	end if
 next
 tempStr	= replace(tempStr,"'","''")
 tempStr	= replace(tempStr,"<","")
 tempStr	= replace(tempStr,">","")
 tempStr	= replace(tempStr,"%0d","")
 tempStr	= replace(tempStr,"%0D","")
 tempStr	= replace(tempStr,"%0a","")
 tempStr	= replace(tempStr,"%0A","")
 tempStr	= replace(tempStr,"\r\n","")
 tempStr	= replace(tempStr,"\r","")
 tempStr	= replace(tempStr,"\n","")
 tempStr	= replace(tempStr,"\R\N","")
 tempStr	= replace(tempStr,"\R","")
 tempStr	= replace(tempStr,"\N","")
 tempStr	= replace(tempStr,"&","")
 tempStr	= replace(tempStr,"#","")
 tempStr	= replace(tempStr,"%","")
	
	if tempStr<>"" then
	 	if IsNumeric(tempStr) then
	 		if InStr(Cstr(10/3),",")>0 then
				if Instr(tempStr,".")>0 then
					tempStr=FormatNumber(tempStr,,,,0)
	 				tempStr=replace(tempStr,".",",")
				end if
	 		end if
	 	end if
	end if
 
 getUserVM_XMLOutPut	= tempStr 
end function

Response.write Request.ServerVariables("LOCAL_ADDR") 
'//Set redirect page to the current file name
session("redirectPage")="gwTGSPay.asp"

'//VirtualMerchant Gateway URL
Dim pcVMPayURL

'//Declare and Retrieve Customer's IP Address	
Dim pcCustIpAddress
pcCustIpAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcCustIpAddress="" Then pcCustIpAddress = Request.ServerVariables("REMOTE_ADDR")
	
'//Declare URL path to gwSubmit.asp	
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
		
'//Get Order ID
if session("GWOrderId")="" then
	session("GWOrderId")=request("idOrder")
end if

'//Retrieve customer data from the database using the current session id		
pcGatewayDataIdOrder=session("GWOrderID")
%>
<!--#include file="pcGateWayData.asp"-->
<% '//Set customer session - we may now be on a different server where this session was lost
session("idCustomer")=pcIdCustomer

'//Retrieve any gateway specific data from database or hard-code the variables
query= "SELECT TOP 1 pcPay_TGS_AccountID,pcPay_TGS_Authkey,pcPay_TGS_Cur,pcPay_TGS_TransType,pcPay_TGS_TestMode,pcPay_TGS_CVV2 FROM pcPay_TGS"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
pcPay_TGS_MerchantID=rs("pcPay_TGS_AccountID")
pcPay_TGS_Authkey=rs("pcPay_TGS_AuthKey")
pcPay_TGS_cur=rs("pcPay_TGS_Cur")
pcPay_TGS_AuthKey=enDeCrypt(pcPay_TGS_Authkey, scCrypPass)
pcPay_TGS_TransType=rs("pcPay_TGS_TransType")
pcPay_TGS_TestMode=rs("pcPay_TGS_TestMode")
pcPay_TGS_CVV2=rs("pcPay_TGS_CVV2")

set rs=nothing

if request("PaymentSubmitted")="Go" then

   pExpiration=getUserInput(request("expMonth"),0) & "/01/" & getUserInput(request("expYear"),0)				
	' validates expiration
	if DateDiff("d", Month(Now)&"/01/"&Year(now),pExpiration)<=-1 then

        call closeDb()
        Session("message") = dictLanguage.Item(Session("language")&"_paymntb_o_6")
        Session("backbuttonURL") = tempURL & "?psslurl="&session("redirectPage")&"&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
        response.redirect "msgb.asp?back=1"

	end if

     pcVMPayURL="https://tgs.ecsuite.com/merchant_interface/XMLServlet"

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Dim SrvVMPayXmlHttp, pcVMPayXMLPostData
	pcVMPayXMLPostData=""
	pcVMPayXMLPostData = pcVMPayXMLPostData & "<TGSMerchantRequest version=""0.1"">"	

	pcVMPayXMLPostData=pcVMPayXMLPostData&"<AuthKey merchantId="""&pcPay_TGS_MerchantID&""">"&pcPay_TGS_Authkey&"</AuthKey>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<Transactions>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<CreditCardAuthRequest capture="""&pcPay_TGS_TransType&"""  merchantTransactionId="""&session("GWOrderID")&""">"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<CustomerAccountInfo>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<CardPresent>false</CardPresent>" 
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<NameOnCard>"&getUserVM_XMLOutPut(trim(pcBillingFirstName),100)&" "&getUserVM_XMLOutPut(trim(pcBillingLastName),100)&"</NameOnCard>"
	If pcPay_TGS_CVV2="1" Then
		pcVMPayXMLPostData=pcVMPayXMLPostData&"<CardNumber cvv="""&getUserVM_XMLOutPut(request("CVV"),4)&""" cvvState=""1"">"&getUserVM_XMLOutPut(request("CardNumber"),16)&"</CardNumber>"
	else
		pcVMPayXMLPostData=pcVMPayXMLPostData&"<CardNumber>"&getUserVM_XMLOutPut(request("CardNumber"),16)&"</CardNumber>"
	end if 
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ExpDate>"&getUserVM_XMLOutPut(request("expMonth")&request("expYear"),0)&"</ExpDate>"

	pcVMPayXMLPostData=pcVMPayXMLPostData&"<CardholderIDMethod>4</CardholderIDMethod>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"</CustomerAccountInfo>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<MerchantInfo>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<TerminalID>7182</TerminalID></MerchantInfo>"'<MCC></MCC><Descriptor>Processed by EC Suite</Descriptor><CustomerServiceNumber></CustomerServiceNumber><Province></Province><CountryCode></CountryCode><PostalCode></PostalCode><TZOffset></TZOffset>"
    pcVMPayXMLPostData=pcVMPayXMLPostData&"<Currency>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<Amount>"&(pcBillingTotal * 100)&"</Amount>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<Code>"&pcPay_TGS_cur&"</Code>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"</Currency>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<AddressInfo>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<IP>"&pcCustIpAddress&"</IP>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<StreetAddress>"&getUserVM_XMLOutPut(pcBillingAddress,255)& " " &getUserVM_XMLOutPut(pcBillingAddress2,0) & "</StreetAddress>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<City>"&getUserVM_XMLOutPut(pcBillingCity,30)&"</City>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<Province>"&getUserVM_XMLOutPut(pcBillingState,30)&"</Province>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<PostalCode>"&getUserVM_XMLOutPut(pcBillingPostalCode,10)&"</PostalCode>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<CountryCode>"&getUserVM_XMLOutPut(pcBillingCountryCode,0)&"</CountryCode>"		
	pcVMPayXMLPostData=pcVMPayXMLPostData&"</AddressInfo>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"<ACI>Y</ACI><ECI>00</ECI>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"</CreditCardAuthRequest>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"</Transactions>"
	pcVMPayXMLPostData=pcVMPayXMLPostData&"</TGSMerchantRequest>"
	
	'response.write "<pre>" & pcVMPayXMLPostData &"</pre>"
	'Response.end
    resolveTimeout	= 500000
	connectTimeout	= 500000
	sendTimeout		= 500000
	receiveTimeout	= 1000000
	 

	Set SrvVMPayXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	err.clear
	SrvVMPayXmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
	SrvVMPayXmlHttp.open "POST", pcVMPayURL & "" , false
   ' SrvVMPayXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
	SrvVMPayXmlHttp.send(pcVMPayXMLPostData)
	if err.number<>0 then
		response.write "ERROR: "&err.description
		response.end
	end if
	VMPayResult = SrvVMPayXmlHttp.responseText
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")    
	xmlDoc.async = False
	pcResultVMPay_Result=""
	pcResultVMPay_Msg=""
	pcResultVMPay_ErrCode=""
	pcResultVMPay_ErrName=""
	pcResultVMPay_ErrMsg=""
	pcResultVMPay_TransID=""
	pcResultVMPay_AuthCode=""
	
	
	'Response.write "<BR><BR><PRE>" & SrvVMPayXmlHttp.responseText &"</PRE>"
 
	If xmlDoc.loadXML(SrvVMPayXmlHttp.responseText) Then
	
			if instr(SrvVMPayXmlHttp.responseText, "ErrorResponse")  then 			
			 pcResultErrorMsg = xmlDoc.documentElement.selectSingleNode("//TGSMerchantResponse/Responses/ErrorResponse/ResponseText").Text
			 pcResultApproved ="N"
						 
			Else
				pcResultApproved = xmlDoc.documentElement.selectSingleNode("//TGSMerchantResponse/Responses/CCResponse/Approved").Text
				SET XMLatt = xmlDoc.documentElement.selectSingleNode("//TGSMerchantResponse/Responses/CCResponse/")				
				pcResultErrorMsg = xmlDoc.documentElement.selectSingleNode("//TGSMerchantResponse/Responses/CCResponse/ResponseText").Text				
				Set domMethodList = xmldoc.documentElement.getElementsByTagname("CCResponse")
				For Each domMethod In domMethodList
				pcResultVMPay_TransID = domMethod.getAttribute("transactionId")
				Next
								
			End if 

	Else
		'//ERROR
		response.write "Failed to process response"
		response.end
	End If
	
	if pcResultApproved="Y" then
		'create sessions
		session("GWAuthCode")=pcResultVMPay_AuthCode
		session("GWTransId")=pcResultVMPay_TransID
		session("GWTransType")=pcPay_TGS_TransType
        
		call closedb()		
		Response.redirect "gwReturn.asp?s=true&gw=TGS"
		response.end
        
	else

        call closeDb()
        Session("message") = pcResultErrorMsg
        Session("backbuttonURL") = tempURL & "?psslurl="&session("redirectPage")&"&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderId")
        response.redirect "msgb.asp?back=1"

	end if

	'*************************************************************************************
	' END
	'*************************************************************************************
end if 
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
                    
					<% If pcPay_TGS_CVV2="1" Then %>
                <div class="pcFormItem">
                    <div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_GateWay_11")%></div>
                    <div class="pcFormField"><input name="CVV" type="text" id="CVV" value="" size="4" maxlength="4"></div>
                </div> 
                <div class="pcFormItem">
                    <div class="pcFormLabel">&nbsp;</div>
                    <div class="pcFormField"><img src="<%=pcf_getImagePath("images","CVC.gif")%>" alt="cvc code" width="212" height="155"></div>
                </div>
					<% End If %>

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
<!--#include file="footer_wrapper.asp"-->
