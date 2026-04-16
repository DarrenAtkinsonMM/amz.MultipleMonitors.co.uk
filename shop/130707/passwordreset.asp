<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% 
pageTitle="Reset Password" 
pageIcon="pcv4_keys.png"
Section="" 
%>
<%PmAdmin=19%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<!--#include file="../includes/pcFormHelpers.asp" -->
<%
pcStrPageName = "passwordreset.asp"

tmpAdminID = getUserInput(request("aid"),0)
tmpGUID = getUserInput(request("GUID"),0)

tmpResult = pcf_CheckPRGuidAdmin(tmpAdminID, tmpGuid)
if tmpResult>0 then
	call closedb()
	response.redirect "login_1.asp?s=1&msg=" & dictLanguage.Item(Session("language")&"_security_2") 
end if

if request.form("updatemode")="1" then

	pcvErrMsg=""
	
	pcv_NewPass1 = getUserInput(request("NewPass1"), 0)
	pcv_NewPass2 = getUserInput(request("NewPass2"), 0)

	if (pcv_NewPass1="") OR (pcv_NewPass2="") OR (pcv_NewPass1<>pcv_NewPass2) then
		pcvErrMsg = dictLanguage.Item(Session("language")&"_newpass_15")
	end if
	
	If pcvErrMsg<>"" Then
		Session("message") = pcvErrMsg
		response.redirect pcStrPageName
	else
	
		pcv_NewPass1 = pcf_PasswordHash(pcv_NewPass1)
		query="UPDATE [admins] SET [adminpassword]='" & pcv_NewPass1 & "' WHERE [idadmin] = " & tmpAdminID & ";"
		set query=connTemp.execute(query)
		set rs=nothing
		
		call pcs_UpdatePRGuid(tmpGUID, "1")
		
		'//call pcs_SaveUsedPass(tmpAdminID, pcStrCustomerPassword)
				
		'//call pcs_SendResetPassMail(tmpAdminID, "")
		
		pcv_intSuccess=1
	End If

End If
%>
 
<div id="pcMain" class="container-fluid">		
    <div class="row">
  
        <% if pcv_intSuccess<>1 then %>

			<%
			msg = ""
			If Session("message") <> "" Then
				msg = Session("message")
				Session("message") = ""
			End If

			If msg <> "" Then
				%><div class="pcErrorMessage"><%= msg %></div><%
			End If
			%>

			<form method="post" id="resetpass" name="resetpass" action="<%=pcStrPageName%>" class="form" role="form">
				<input type="hidden" name="updatemode" value="1">
				<input type="hidden" id="aid" name="aid" value="<%=tmpAdminID%>">
				<input type="hidden" name="GUID" value="<%=tmpGUID%>">
				
				<div class="form-group">
					<label for="NewPass1"><%= dictLanguage.Item(Session("language")&"_order_H")%></label>
					<input type="password" class="form-control" id="NewPass1" name="NewPass1" autocomplete="off">
				</div>
		
				  
				<div class="form-group">
					<label for="NewPass2"><%= dictLanguage.Item(Session("language")&"_order_I")%></label>
					<input type="password" class="form-control" id="NewPass2" name="NewPass2" autocomplete="off">
				</div>
				
				<script>
					var validator0
					$pc(document).ready(function () {
					var validator0 = $pc("#resetpass").validate({
					rules: {
						NewPass1: {
							required: true,
							remote: {
								type: 'POST',
								url: "../pc/checkPass.asp",
								data: {
									passtype: "Cp",
									pass: function () {
										return $pc("#NewPass1").val();
									},
									email: function () {
										return $pc("#email").val();
									}
									},
								dataFilter: function(data) {
									var myjson = JSON.parse(data);
									if(myjson.isError == "true") {
										return "\"" + myjson.errorMessage + "\"";
									} else {
										return true;
									}
								}
							}
						},
						NewPass2: {
							required: true,
							equalTo: "#NewPass1"
						}
					},
					messages: {
						pcCustomerPassword: {
							required: "<%=dictLanguage.Item(Session("language")&"_opc_js_4")%>"
						},
						pcCustomerConfirmPassword: {
							required: "<%=dictLanguage.Item(Session("language")&"_opc_js_47")%>",
							equalTo: "<%=dictLanguage.Item(Session("language")&"_opc_js_48")%>"
						}
					}
					});
					});
				</script>
          
				<div class="form-group">
					<button class="btn btn-default" id="FormSubmit" name="FormSubmit">						
						<%= dictLanguage.Item(Session("language")&"_css_submit") %>
					</button>                       
				</div>
      	</form>

    <% else %>

        <p>
            <%= dictLanguage.Item(Session("language")&"_newpass_17")%>
        </p>     
   
        <div class="form-group">
      	    <a class="btn btn-default" href="menu.asp">
                <%= dictLanguage.Item(Session("language")&"_css_submit") %>
            </a>
        </div>
        
    <% end if %>
  </div>
</div>
<!--#include file="AdminFooter.asp"-->
