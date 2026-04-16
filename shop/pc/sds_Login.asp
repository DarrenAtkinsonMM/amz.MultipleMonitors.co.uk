<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="pcStartSession.asp"-->

<% if request.form("SubmitCO")<>"" then
	ErrCnt=0
	EP=0
	'form is submitted
	sds_username=replace(request.form("sds_username"),"'","''")
	session("sds_username")=sds_username
	if sds_username="" then
		ErrCnt=ErrCnt+1
	End if
	sds_password=request.form("sds_password")
	if sds_password="" then
		ErrCnt=ErrCnt+1
		EP=1
	End if
	
	If ErrCnt>0 then
		If (scSecurity=1) and (scAlarmMsg=1) then
				if session("AttackCount")="" then
					session("AttackCount")=0
				end if
				session("AttackCount")=session("AttackCount")+1
				if session("AttackCount")>=scAttackCount then
				session("AttackCount")=0%>
				<!--#include file="../includes/sendAlarmEmail.asp" -->
				<%end if	
		End if
		response.redirect "sds_Login.asp?EP="&EP&"&msg=1"
	Else
		erypassword=encrypt(sds_password, 9286803311968)
		session("sds_erypassword")=erypassword
		response.redirect "sds_LoginB.asp" 
	End if

end if
%>
<% ' if Drop-Shipper already login
if (Session("pc_idsds")<>"") and (Session("pc_idsds")<>"0") then
 response.redirect "sds_MainMenu.asp"
end if
%>
<!--#include file="header_wrapper.asp"-->
<%
pcRequestRedirect=trim(getUserInput(request("redirectUrl"),250))
if len(pcRequestRedirect)>0 then
	session("redirectUrlLI")=pcRequestRedirect
end if
%>
<%
'// START - Check for SSL and redirect to SSL login if not already on HTTPS
	If scSSL="1" And scIntSSLPage="1" Then
		If (Request.ServerVariables("HTTPS") = "off") Then
		Dim xredir__, xqstr__
		xredir__ = "https://" & Request.ServerVariables("SERVER_NAME") & _
		Request.ServerVariables("SCRIPT_NAME")
		xqstr__ = Request.ServerVariables("QUERY_STRING")
		if xqstr__ <> "" Then xredir__ = xredir__ & "?" & xqstr__
		Response.redirect xredir__
		End if
	End If
'// END - check for SSL
%>
<div id="pcMain">
	<div class="pcMainContent">
		<form method="post" name="auth" action="sds_Login.asp" class="pcForms">
			<% 
				msg = ""
				code = getUserInput(Request.QueryString("msg"), 0)

				If code = "1" Then
					msg = dictLanguage.Item(Session("language")&"_Custmoda_18")
				End If

				If msg<>"" Then	
					%><div class="pcErrorMessage"><%= msg %></div><%
				End If 
			%>
		
    
    	<!-- start of login form -->
      <div class="pcShowContent"> 
      	<h2><%= dictLanguage.Item(Session("language")&"_sdsLogin_1")%></h2>
				<div class="pcFormItem">
					<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_sdsLogin_2")%></div>
				</div>
        
				<div class="pcSpacer"></div>

        <div class="pcFormItem">
        	<div class="pcFormLabel">
          	<%= dictLanguage.Item(Session("language")&"_sdsLogin_3")%>
          </div>
          <div class="pcFormField">
            <input type="text" name="sds_username" size="30" maxlength="150" value="<%=session("sds_username")%>">
            <% if msg="" then %>
              <img src="<%=pcf_getImagePath("",rsIconObj("requiredicon"))%>" alt="Required">
            <% else
               if session("email")="" then %>
               <img src="<%=pcf_getImagePath("",rsIconObj("errorfieldicon"))%>" alt="Error">
            	<% end if %>
            <% end if %>
          </div>
        </div>
        
        <div class="pcFormItem">
        	<div class="pcFormLabel">
						<%= dictLanguage.Item(Session("language")&"_sdsLogin_4") %>
          </div>
          <div class="pcFormField">
            <input type="password" name="sds_password" size="30" maxlength="150">
            <% if msg="" then %>
            <img src="<%=pcf_getImagePath("",rsIconObj("requiredicon"))%>" alt="Required">
            <% end if %>
          </div>
        </div>
        
				<div class="pcSpacer"></div>

        <div class="pcFormButtons">
        	<button class="pcButton pcButtonLogin" name="SubmitCO" value="Submit" id="submit">
          	<img src="<%=pcf_getImagePath("",rslayout("login"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_login") %>"/>
            <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_login") %></span>
          </button>
        </div>
        <!-- end of login table -->
      </div>
		</form>
		<hr>
  
		<!-- start of password request -->
    <% if request.querystring("s")="1" then %>
      <div class="pcErrorMessage">
        <%= dictLanguage.Item(Session("language")&"_sdsLogin_5")%>
      </div>
    <% else %>
      <p><%= dictLanguage.Item(Session("language")&"_sdsLogin_6")%>
      <a href="sds_fpass.asp?redirectUrl=<%if request("redirectUrl")<>"" then%><%=Server.URLEnCode(session("redirectUrlLI"))%><%else%><%=Server.URLEncode(session("redirectUrlLI"))%><%end if%>&frURL=sds_Login.asp"><%= dictLanguage.Item(Session("language")&"_sdsLogin_7")%></a></p>
    <% end if %>
		<!-- end of password request -->
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
<%call clearLanguage()%>
