<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="UPS Shipping Configuration" %>
<% Section="shipOpt" %>
<%PmAdmin=4%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
on error resume next

'Let's check for XML parser
dim xml, XMLAvailable, XMLUse, XML_checked, XML_Err_reason, XML_Err_reason_2, XML3_checked, XML3_Err_reason, XML3_Err_reason_2, XML4_checked, XML4_Err_reason, XML4_Err_reason_2
xml = "<?xml version=""1.0"" encoding=""UTF-16""?><cjb></cjb>"
XMLAvailable=0
XML3=""
XMLUse=""
XML_checked = ""
XML_Err_reason = "Installed"
XML_Err_reason_2 = ""
XML3_checked = ""
XML3_Err_reason = "Installed"
XML3_Err_reason_2 = ""
XML6_checked = ""
XML6_Err_reason = "Installed"
XML6_Err_reason_2 = ""

testURL="https://onlinetools.ups.com/ups.app/xml/Rate"
 
err.clear
Set x = server.CreateObject("Msxml2.DOMDocument")
x.async = false 
if x.loadXML(xml) then
	XML_checked="checked"
end if
set x=nothing

if err.number<>0 then
	XML_Err_reason=err.description
	XML_checked=""
	err.clear
else
	Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp")
	srvXmlHttp.open "POST", testURL, false
	if err.number<>0 then
		XML_Err_reason_2=err.description
		err.clear
	else
		srvXmlHttp.send(xml)
		if err.number<>0 then
			XML_Err_reason_2=err.description
			err.clear
		else
			XMLAvailable=1
			XMLUse=""
		end if
	end if
	set srvXmlHttp=nothing
end if
									
dim intReqXML
intReqXML=0

err.clear
Set x = server.CreateObject("Msxml2.DOMDocument.3.0")
x.async = false 
if x.loadXML(xml) then
	XML3_checked="checked"
end if
set x=nothing
if err.number<>0 then
	XML3_Err_reason=err.description
	XML3_checked=""
	err.clear
else
	Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp.3.0")
	srvXmlHttp.open "POST", testURL, false
	if err.number<>0 then
		XML3_Err_reason_2=err.description
		err.clear
	else
		srvXmlHttp.send(xml)
		if err.number<>0 then
			XML3_Err_reason_2=err.description
			err.clear
		else
			XMLAvailable=1
			XML3=".3.0"
			XMLUse=".3.0"
			intReqXML=1
		end if
	end if
	set srvXmlHttp=nothing
end if

err.clear
Set x = server.CreateObject("Msxml2.DOMDocument.6.0")
x.async = false 
if x.loadXML(xml) then
	XML6_checked="checked"
end if
set x=nothing
if err.number<>0 then
	XML6_Err_reason=err.description
	XML6_checked=""
	err.clear
else
	Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp.6.0")
	srvXmlHttp.open "POST", testURL, false
	if err.number<>0 then
		XML6_Err_reason_2=err.description
		err.clear
	else
		srvXmlHttp.send(xml)
		if err.number<>0 then
			XML6_Err_reason_2=err.description
			err.clear
		else
			XMLAvailable=1
			XMLUse=".6.0"
			intReqXML=1
		end if
	end if
	set srvXmlHttp=nothing
end if

pcv_XML=""
if XML3<>"" then
	pcv_XML=XML3
else
	pcv_XML=XMLUse
end if

if scXML="" and pcv_XML<>"" then %>
	<!--#include file="pcAdminRetrieveSettings.asp"-->
	<% pcStrXML = pcv_XML

	'/////////////////////////////////////////////////////
	'// Update database with new Settings
	'/////////////////////////////////////////////////////
	%>
	<!--#include file="pcAdminSaveSettings.asp"-->
	<% call closeDb()
response.redirect "ConfigureOption1.asp"
end if


'check if UPS has been configured or not
query="SELECT ups_UserId, ups_Password, ups_AccessLicense FROM ups_license WHERE idUPS=1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if NOT rs.eof then
	if rs("ups_UserId")<>"" then
		'already setup an account
		Session("ship_UPS_userID")=EnDeCrypt(rs("ups_UserId"),scCrypPass)
		Session("ship_UPS_Password")=EnDeCrypt(rs("ups_Password"),scCrypPass)
		Session("ship_UPS_AccessLicense")=EnDeCrypt(rs("ups_AccessLicense"),scCrypPass)
		call closeDb()
response.redirect "1_Step2.asp"
	end if
else
	query="INSERT INTO ups_license (idUPS) VALUES (1);"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
end if
%>
	<script type=text/javascript>
	function win(fileName)
		{
		myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=640,height=500')
		myFloater.location.href=fileName;
		}
	</script>

<% if request.form("submit")<>"" then
	UPSAccessLicense=request.form("UPSAccessLicense")
	Session("ship_UPS_AccessLicense")=UPSAccessLicense
	UPSID=request.form("UPSID")
	Session("ship_UPS_userID")=UPSID
	UPSPassword=request.form("UPSPassword")
	Session("ship_UPS_Password")=UPSPassword
	if UPSAccessLicense="" or UPSID="" or UPSPassword="" then
		call closeDb()
response.redirect "ConfigureOption1.asp?msg="&Server.URLEncode("All fields are required.")
		response.end
	end if
	call closeDb()
response.redirect "1_Step2.asp"
	response.end
else  %>
    <form name="form1" method="post" action="ConfigureOption1.asp">
        <table width="94%" border="0" cellpadding="4" cellspacing="0" align="center">
			<% if intReqXML=0 then %>
    
                <tr> 
                    <td colspan="2" class="normal"> 
                        <table width="100%" border="0" cellspacing="0" cellpadding="4">
                        <tr> 
                        <td width="4%"><img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"></td>
                        <td width="96%" class="small"><font color="#FF9900"><b>The XML Parser version 3.0 is not installed or has returned errors while trying to connect. An XML parser version 3.0 or higher is required in order to UPS as a dynamic shipping provider. Contact your hosting provider and ask them to install or reinstall the XML Parser version 3.0.</b></font></td>
                        </tr>
                        </table></td>
                </tr>    
			<% else %>
                <% if request.querystring("msg")<>"" then %>
                    <tr> 
                        <td colspan="2" class="normal"> 
                            <table width="100%" border="0" cellspacing="0" cellpadding="4">
                            <tr> 
                            <td width="4%"><img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"></td>
                            <td width="96%" class="small"><font color="#FF9900"><b><%=request.querystring("msg")%></b></font></td>
                            </tr>
                            </table></td>
                    </tr>
                <% end if %>
                                                                
                <tr> 
                    <td colspan="2" bgcolor="e1e1e1" class="normal">Enable <b>UPS&reg; Developer Kit</b> ( <a href="javascript:win('../UPSLicense/licenseAgrRequest.asp')">Web site</a> )</td>
                </tr>
                <tr> 
                    <td colspan="2" class="normal">&nbsp; 
                        <table width="100%" border="0" cellspacing="0" cellpadding="2">
                            <tr class="normal">
                                <td width="10%" valign="top"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50"></td>
                                <td width="90%">
                                    <p>In order to use UPS&reg; Developer Kit, you need to <a href="javascript:win('../UPSLicense/licenseAgrRequest.asp')">register</a> an account with the company. Registration is free and includes access to the following UPS&reg; Developer Kit:</p>
                                        <ul>
                                            <li>UPS&reg; Developer Kit Tracking</li>
                                            <li>UPS&reg; Developer Kit Rates &amp; Service Selection</li>
                                        </ul>
                                    <p>To register an account <a href="javascript:win('../UPSLicense/licenseAgrRequest.asp')">click here</a>.</p>
                                    <p>UPS, the UPS Shield trademark, the UPS Ready mark, <br />the UPS Developer Kit mark and the Color Brown are trademarks of <br />United Parcel Service of America, Inc. All Rights Reserved.</p>
                                    <p><input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()"></p></td>
                            </tr>
                        </table>
                    </td>
                </tr>
		   <% end if %>
        </table>
    </form>
<% end if %>
<!--#include file="AdminFooter.asp"-->