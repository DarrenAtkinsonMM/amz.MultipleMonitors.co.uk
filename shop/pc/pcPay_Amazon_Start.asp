<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%pcStrPageName="pcPay_Amazon_Start.asp"%>
<%IF request("action")<>"r" THEN%>
<script type=text/javascript>
	location="<%=pcStrPageName%>?action=r&" + location.hash.substr(1,location.hash.length);
</script>
<%response.End()
END IF%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_AmazonHeader.asp" -->
<% response.Buffer = true %>
<%
dim Info
tmpscXML=".3.0"
session("ExpressPayMethod") = "AMZ"

'*****************************************************************************************************
' START: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*****************************************************************************************************
' END: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************


'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
ppcCartIndex=Session("pcCartIndex")

If session("customerType")=1 Then
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scWholesaleMinPurchase then  
	'Wholesale minimum not met, so customer cannot checkout -> show message
		response.redirect "msg.asp?message=205"
	end if
Else
	if calculateCartTotal(pcCartArray, ppcCartIndex)<scMinPurchase then
	'Retail minimum not met, so customer cannot checkout -> show message
		response.redirect "msg.asp?message=206"
	end if
End If

IF request("action")="r" THEN
	session("Amz_scope")=getUserInput(request("scope"),0)
	session("Amz_expires_in")=getUserInput(request("expires_in"),0)
	session("Amz_token_type")=getUserInput(request("token_type"),0)
	session("Amz_access_token")=getUserInput(request("access_token"),0)
	
	if session("Amz_access_token")="" then
		response.redirect "viewcart.asp"
		response.end
	end if
	
	QueryStr=pcAMZAPI & "auth/o2/tokeninfo?access_token=" & AmazonURLEnCode(session("Amz_access_token"))
	
	Set xml = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)
	xml.open "GET", QueryStr, False
	'xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xml.send ""
	strStatus = xml.Status
	
	'store the response
	strRetVal = xml.responseText
	
	if (strRetVal="") then
		response.redirect "viewcart.asp"
		response.end
	end if

	if InStr(strRetVal,"aud")=0 then    
        set Info = JSON.parse(strRetVal) 
        Session("message") = Info.error_description
        Session("backbuttonURL") = "viewcart.asp"
        response.redirect "msgb.asp?back=1"
	end if
	
	set Info = JSON.parse(strRetVal) 
	retClientID = Info.aud

	if UCase(retClientID)<>UCase(pcAMZClientID) then
		response.redirect "viewcart.asp"
		response.end
	end if
	
	QueryStr=pcAMZAPI & "user/profile"
	
	Set xml = server.createobject("MSXML2.serverXMLHTTP"&tmpscXML)
	xml.open "GET", QueryStr, False
	xml.setRequestHeader "Authorization", "bearer " & session("Amz_access_token")
	xml.send ""
	strStatus = xml.Status
	
	'store the response
	strRetVal = xml.responseText   

	if (strRetVal="") then
		response.redirect "viewcart.asp"
		response.end
	end if
	
	if (InStr(strRetVal,"""email""")=0) OR (InStr(strRetVal,"""name""")=0) OR (InStr(strRetVal,"""user_id""")=0) then
		response.redirect "viewcart.asp"
		response.end
	end if
	
	tmpStr1=Mid(strRetVal,3,len(strRetVal)-4)
	tmpStr1=split(tmpStr1,""",""")
	pcAMZ_CustEmail=""
	pcAMZ_CustName=""
	pcAMZ_CustID=""
	For i=lbound(tmpStr1) to ubound(tmpStr1)
		if tmpStr1(i)<>"" then
			tmpStr2=split(tmpStr1(i),""":""")
			Select Case tmpStr2(0)
			Case "email": pcAMZ_CustEmail=tmpStr2(1)
			Case "name": pcAMZ_CustName=tmpStr2(1)
			Case "user_id": pcAMZ_CustID=tmpStr2(1)
			End Select
		end if
	Next

	if (pcAMZ_CustEmail<>"") AND (pcAMZ_CustName<>"") AND (pcAMZ_CustID<>"") then
	
		'// Create a Session Flag
		session("ExpressCheckoutPayment")="YES"
		session("PayWithAmazon")="YES"
		strEmail=pcAMZ_CustEmail
		strPassword=randomNumber(9999999)
		strPassword=enDeCrypt(strPassword, scCrypPass)
		pCustomerType = 0
		pIdRefer = 0
		pRecvNews = 0
		session("AmazonFirstTime")="0"
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Update Customer Details
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
		'// Customer Logged into ProductCart
		if session("idCustomer")<>"" and session("idCustomer")<>0 then

			query="SELECT pcCust_AmazonID FROM Customers WHERE idCustomer="&session("idCustomer")&";"
			set rs=conntemp.execute(query)
			if not rs.eof then
				tmpAmazonCustID=rs("pcCust_AmazonID")
				if (IsNull(tmpAmazonCustID)) OR (tmpAmazonCustID="") OR (tmpAmazonCustID<>pcAMZ_CustID) then
					query="UPDATE Customers SET email='" & pcAMZ_CustEmail & "',pcCust_AmazonID='" & pcAMZ_CustID & "' WHERE idCustomer="&session("idCustomer")&";" 
					set rs=connTemp.execute(query)
				end if
			end if
			set rs=nothing

			response.redirect "OnePageCheckout.asp"
		
		'// Customer NOT Logged into ProductCart
		else

			'// Check if Customer Exists
			query="SELECT idCustomer, pcCust_Guest, customerType, pcCust_AmazonID FROM customers WHERE email='"&strEmail&"' AND pcCust_Guest=0;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)				
			
			'// Email Does Not Exist - Create New Customer
			if rs.eof then		
				pcv_dateCustomerRegistration=Date()
				if SQL_Format="1" then
					pcv_dateCustomerRegistration=Day(pcv_dateCustomerRegistration)&"/"&Month(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
				else
					pcv_dateCustomerRegistration=Month(pcv_dateCustomerRegistration)&"/"&Day(pcv_dateCustomerRegistration)&"/"&Year(pcv_dateCustomerRegistration)
				end if
				
				if Instr(pcAMZ_CustName," ")>0 then
					tmpAMZFName=Left(pcAMZ_CustName,Instr(pcAMZ_CustName," ")-1)
					tmpAMZLName=Mid(pcAMZ_CustName,Instr(pcAMZ_CustName," ")+1,len(pcAMZ_CustName))
				else
					tmpAMZFName=pcAMZ_CustName
					tmpAMZLName=""
				end if
							
				query="INSERT INTO customers (name, lastName, email, [password],  customerType, IDRefer, RecvNews,pcCust_DateCreated,pcCust_Guest, pcCust_AmazonID) VALUES ('" &tmpAMZFName& "', '" &tmpAMZLName& "', '" &pcAMZ_CustEmail& "', '" &strPassword&"'," &pCustomerType& ","&pIdRefer&"," &pRecvNews&",'" & pcv_dateCustomerRegistration & "',0,'" & pcAMZ_CustID & "');"
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)	
				set rstemp=nothing
				
				query="SELECT idCustomer, pcCust_Guest FROM customers WHERE email='"&strEmail&"' AND pcCust_Guest=0 ORDER BY idCustomer DESC;"
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=conntemp.execute(query)
				session("AmazonFirstTime")="1"				
				session("idCustomer")=rstemp("idCustomer")	
				session("CustomerGuest")=rstemp("pcCust_Guest")
				session("customerType")="0"
				session("isCustomerNew")="YES"				
				set rstemp=nothing				
			
			'// Email Does Exist - Login Customer
			else 
				intIdCustomer=rs("idCustomer")
				intCustomerGuest=rs("pcCust_Guest")
				session("customerType")=rs("customerType")
				tmpAmazonCustID=rs("pcCust_AmazonID")
				if (IsNull(tmpAmazonCustID)) OR (tmpAmazonCustID="") OR (tmpAmazonCustID<>pcAMZ_CustID) then
					query="UPDATE Customers SET email='" & pcAMZ_CustEmail & "',pcCust_AmazonID='" & pcAMZ_CustID & "' WHERE idCustomer="&session("idCustomer")&";" 
					set rs=connTemp.execute(query)
				end if
				session("idCustomer")=intIdCustomer	
				session("CustomerGuest")=intCustomerGuest			
				set rs=nothing
			end if

		end if	
			
		
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Update Customer Sessions
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	%>
		<!--#include file="DBsv.asp" -->
		<%
		query="SELECT payTypes.idPayment FROM payTypes WHERE (payTypes.active = - 1) AND (payTypes.gwCode = 88);"
		set rs=connTemp.execute(query)
		if not rs.eof then
			AmzidPayment=rs("idPayment")
		end if
		set rs=nothing
		
		session("pcSFIdPayment")=AmzidPayment
		
		query="UPDATE pcCustomerSessions SET pcCustSession_IdPayment=" & AmzidPayment& ", idCustomer="&session("idCustomer")&" WHERE (idDbSession="&session("pcSFIdDbSession")&") AND (randomKey="&session("pcSFRandomKey")&");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' Start: Update Customer Sessions
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	

		set rs=nothing
		call closedb()	


		If session("customerType")=1 Then
			if calculateCartTotal(pcCartArray, ppcCartIndex)<scWholesaleMinPurchase then  
			'Wholesale minimum not met, so customer cannot checkout -> show message
				response.redirect "msg.asp?message=205"
			end if
		Else
			if calculateCartTotal(pcCartArray, ppcCartIndex)<scMinPurchase then
			'Retail minimum not met, so customer cannot checkout -> show message
				response.redirect "msg.asp?message=206"
			end if
		End If
		
		
		response.redirect "OnePageCheckout.asp"
	
	end if
END IF

Function AmazonURLEnCode(tmpStr)
Dim tmp1

tmp1=tmpStr
tmp1=replace(Server.URLEncode(tmp1),"%2E",".")
tmp1=replace(tmp1,"%5F","_")
tmp1=replace(tmp1,"%2D","-")
tmp1=replace(tmp1,"%7E","~")
tmp1=replace(tmp1,"+","%20")
AmazonURLEnCode=tmp1

End Function

function randomNumber(limit)
	randomize
	randomNumber=int(rnd*limit)+2
end function

%>