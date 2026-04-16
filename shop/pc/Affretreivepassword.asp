<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/validation.asp" --> 
<!--#include file="pcStartSession.asp"-->
<!--#include file="DBsv.asp"--> 
<!--#include file="../includes/sendmail.asp"--> 
<!--#include file="header_wrapper.asp"-->
<%
dim fName, fLastname, fEmail, fPassword, fFrom, fFromName, fSubject, fBody, ftmp,ftmp1,ftmp2

fEmail=replace(trim(request.querystring("email")),"'","''")
redirectUrl= server.HTMLEncode(Session("pcSF_redirectUrl"))
Session("pcSF_redirectUrl")=""
frURL=server.HTMLEncode(Session("pcSF_pcfrUrl"))
Session("pcSF_pcfrUrl")=""

mySQL="SELECT Affiliatename, AffiliateEmail, [pcAff_Password] from Affiliates WHERE AffiliateEmail='" &fEmail& "'"
set rs=conntemp.execute(mySQL)	
if not rs.eof then
	fName=rs("Affiliatename")
	fEmail=rs("Affiliateemail")
	fPassword=enDeCrypt(rs("pcAff_Password"),scCrypPass)		
	fSubject=dictLanguage.Item(Session("language")&"_forgotpasswordmailsubject")
	fBody=dictLanguage.Item(Session("language")&"_forgotpasswordmailbody2")

	fBody=replace(fBody,"#password",fPassword)	
	fBody=replace(fBody,"#name",fName)      
	
	call sendmail (scEmail, scEmail, fEmail, fSubject, fBody) 
	call pcs_hookAffRetrievePassEmailSent(fEmail)
%>
	<div id="pcMain">
		<div class="pcMainContent">
			<div class="pcErrorMessage">
				<%= dictLanguage.Item(Session("language")&"_checkout_11")%>
				<br /><br />
				<% if frURL<>"" then %>
					<a class="pcButton pcButtonSubmit"  href="<%=frURL&"?redirectUrl="&Server.Urlencode(redirectUrl)&"&s=1"%>">
						<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>">
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
					</a>
				<% else %>
					<a class="pcButton pcButtonSubmit" href="AffiliateLogin.asp?s=1">
						<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>">
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
					</a>
				<% end if
				call clearLanguage()
			 %>
			</div>
		</div>
	</div>
<% else %>
	<%
	call closeDb()
	response.redirect "msg.asp?message=2"
	%>
<% end if %>
<!--#include file="footer_wrapper.asp"-->
