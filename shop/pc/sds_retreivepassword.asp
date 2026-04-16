<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="header_wrapper.asp"-->
<%
sds_username=replace(trim(request.querystring("sds_username")),"'","''")
redirectUrl= server.HTMLEncode(Session("pcSF_redirectUrl"))
Session("pcSF_redirectUrl")=""
frURL=server.HTMLEncode(Session("pcSF_pcfrUrl"))
Session("pcSF_pcfrUrl")=""
' verify password for that username
query="SELECT pcDropShipper_ID As idsds, pcDropShipper_Username As sdsUsername,pcDropShipper_Password As sdsPassword, pcDropShipper_FirstName As FirstName, pcDropShipper_LastName As LastName, pcDropShipper_Company As Company,pcDropShipper_Email As sdsEmail FROM pcDropShippers WHERE pcDropShipper_Username='" & sds_username & "' UNION SELECT pcSupplier_ID,pcSupplier_Username,pcSupplier_Password,pcSupplier_FirstName,pcSupplier_LastName, pcSupplier_Company,pcSupplier_Email FROM pcSuppliers WHERE pcSupplier_Username='" & sds_username & "' AND pcSupplier_IsDropShipper=1"
set rs=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if not rs.eof then
	fPassword=enDeCrypt(rs("sdsPassword"),scCrypPass)
	fName=rs("FirstName") & " " & rs("LastName")
	fName=rs("Company") & "(" & fName & ")"
	fEmail=rs("sdsEmail")
	fSubject=dictLanguage.Item(Session("language")&"_forgotpasswordmailsubject")
	fBody=dictLanguage.Item(Session("language")&"_forgotpasswordmailbody2")

	fBody=replace(fBody,"#password",fPassword)	
	fBody=replace(fBody,"#name",fName)      
	
	call sendmail (scEmail, scEmail, fEmail, fSubject, fBody) 
	%>	
	<div id="pcMain">
		<div class="pcErrorMessage">
			<%= dictLanguage.Item(Session("language")&"_checkout_11")%>
    </div>
    <div class="pcFormButtons">
			<%
      if frURL<>"" then
        %>
        <a class="pcButton pcButtonContinue" href="<%=frURL&"?redirectUrl="&Server.Urlencode(redirectUrl)&"&s=1"%>">
          <img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>" />
          <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
        </a> 
      <%else
        %>        
        <a class="pcButton pcButtonContinue" href="sds_Login.asp?s=1">
          <img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>" />
          <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
        </a> 
      <%end if
      call clearLanguage()%>
    </div>
	</div>	
<% else %>
		<%
		call closeDb()
		response.redirect "msg.asp?message=203"
		%>        
<% end if %>
<!--#include file="footer_wrapper.asp"-->
