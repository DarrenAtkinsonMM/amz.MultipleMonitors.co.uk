<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
'//Set redirect page to the current file name
session("redirectPage")="gwPSI.asp"
		
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
query="SELECT Config_File_Name, Config_File_Name_Full, Host, Port, psi_testmode FROM PSIGate WHERE id=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

'// Set gateway specific variables
Config_File_Name=rs("Config_File_Name")
Config_File_Name_Full=rs("Config_File_Name_Full")
Host=rs("Host")
Port=rs("Port")
psi_testmode=rs("psi_testmode")

set rs=nothing

if request("PaymentSubmitted")="Go" then

	'*************************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	Set PSIObj=CreateObject("MyServer.PsiGate")
	
	'This is supplied by PSiGate, must match configurate filename on 
	'PSiGate payment gateway

	PSIObj.Configfile=trim(Config_File_Name)
    
	'This is the location of the certificate file you download from
	'PSiGate

	PSIObj.Keyfile=Server.MapPath(trim(Config_File_Name_Full))
	
	PSIObj.Host=trim(Host)
	PSIObj.Port=trim(Port)
		
	PSIObj.Oid=session("GWOrderId")
	PSIObj.Userid="COM Sample FORM"
	PSIObj.Bname=pcBillingFirstName&" "&pcBillingLastName
	PSIObj.Bcompany=pcBillingCompany
	PSIObj.Baddr1=pcBillingAddress
	PSIObj.Baddr2=pcBillingAddress2
	PSIObj.Bcity=pcBillingCity
	PSIObj.Bstate=pcBillingState
	PSIObj.Bzip=pcBillingPostalCode
	PSIObj.Bcountry=pcBillingCountry
	PSIObj.Sname=pcShippingFirstName&" "&pcShippingLastName
	PSIObj.Saddr1=pcShippingAddress
	PSIObj.Saddr2=pcShippingAddress2
	PSIObj.Scity=pcShippingCity
	PSIObj.Sstate=pcShippingState
	PSIObj.Szip=pcShippingPostalCode
	PSIObj.Scountry=pcShippingCountryCode
	PSIObj.Phone=pcBillingPhone
	PSIObj.Fax=""
	PSIObj.Comments=""
	PSIObj.Cardnumber=Request.Form("Cardnumber")
	PSIObj.Chargetype="0"
	PSIObj.Expmonth=Request.Form("expMonth")
	PSIObj.Expyear=Request.Form("expYear")
	PSIObj.Email=pcCustomerEmail

	'Used during the testing process, 0=Live
	if psi_testmode="YES" then
	else
	PSIObj.Result=0
   end if
	'Used with AVS processing (only in the US)
	'PSIObj.Addrnum="111"
    
	'----------------------------Add items
	ItemID1=Request.Form("ItemID1")			
	Description1=Request.Form("Description1")
	Price1=Request.Form("Price1")
	Quantity1=Request.Form("Quantity1")
	SoftFile1=Request.Form("SoftFile1")
	EsdType1=Request.Form("EsdType1")
	Serial1=Request.Form("Serial1")

  intErr=0  
   
	ret_code=PSIObj.AddItem(ItemID1, Description1, Price1, Quantity1, SoftFile1, EsdType1, Serial1)
	If Not ret_code=1 Then
		Msg="ERROR   " & PSIObj.ErrMsg
		intErr=1
		Set PSIObj=Nothing
	End If


	'-------------------------------Process
	if intErr=0 then
		ret_code=PSIObj.ProcessOrder()
		If Not ret_code=1 Then
			Msg="ERROR   " & PSIObj.ErrMsg
			intErr=1
			Set PSIObj=Nothing
		End If
	end if
	
	'-------------------------------Get results
	if intErr=0 then
		pcv_Response_Approved=PSIObj.Appr
		pcv_Response_Code=PSIObj.code
		pcv_Response_TransTime=PSIObj.transtime
		pcv_Response_Refno=PSIObj.refno
		pcv_Response_Error=PSIObj.Err
		pcv_Response_Orderno=PSIObj.OrdNo
		pcv_Response_Subtotal=CStr(PSIObj.Subtotal)
		pcv_Response_Shiptotal=CStr(PSIObj.Shiptotal)
		pcv_Response_Taxtotal=CStr(PSIObj.Taxtotal)
		pcv_Response_Total=CStr(PSIObj.Total)
		
		Set PSIObj=Nothing
		
		If pcv_Response_Approved="APPROVED" then
			session("GWAuthCode")=pcv_Response_Code
			session("GWTransId")=pcv_Response_Refno
			response.redirect "gwReturn.asp?s=true&gw=PSIGate"
		Else
			Msg=pcv_Response_Error
		End if
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
					<input type="hidden" name="ItemID1" value="Online Order">
					<% If scCompanyName="" then %>
						<input type="hidden" name="Description1" value="Shopping Cart"> 
					<%else %>
						<input type="hidden" name="Description1" value="<%=scCompanyName%>"> 
					<% end if %>
					<input type="hidden" name="Price1" value="<%=pcBillingTotal%>"> 
					<input type="hidden" name="Quantity1" value="1">

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
                    
					<% If pcv_CVV="1" Then %>
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
