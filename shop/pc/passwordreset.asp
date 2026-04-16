<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="header_wrapper.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<!--#include file="../includes/pcFormHelpers.asp" -->
<%
pcStrPageName = "passwordreset.asp"

tmpEmail=getUserInput(request("email"),0)
tmpGUID=getUserInput(request("GUID"),0)

tmpResult=pcf_CheckPRGuid(tmpemail,tmpGuid)
if tmpResult>0 then
	call closedb()
	response.redirect "msg.asp?message=315"
end if

if request.form("updatemode")="1" then

	pcvErrMsg=""
	
	pcv_NewPass1=getUserInput(request("NewPass1"),0)
	pcv_NewPass2=getUserInput(request("NewPass2"),0)

	if (pcv_NewPass1="") OR (pcv_NewPass2="") OR (pcv_NewPass1<>pcv_NewPass2) then
		pcvErrMsg=dictLanguage.Item(Session("language")&"_newpass_15")
	end if
	
	If pcvErrMsg<>"" Then
		Session("message") = pcvErrMsg
		response.redirect pcStrPageName
	else
	
		pcv_NewPass1=pcf_PasswordHash(pcv_NewPass1)
		query="UPDATE Customers SET [password]='" & pcv_NewPass1 & "' WHERE [email] LIKE '" & tmpEmail & "';"
		set query=connTemp.execute(query)
		set rs=nothing
		
		query="SELECT idCustomer FROM Customers WHERE [email] LIKE '" & tmpEmail & "';"
		set rs=connTemp.execute(query)
		if not rs.eof then
			pIdCustomer=rs("idCustomer")
		end if
		set rs=nothing
		
		call pcs_UpdatePRGuid(tmpGUID,"1")
		
		call pcs_SaveUsedPass(pIdCustomer,pcStrCustomerPassword)
				
		call pcs_SendResetPassMail(pIdCustomer,"")
		
		pcv_intSuccess=1
	End If

End If


%>
<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="Contact Us">Reset Password</h3>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
	<!-- /Header: pagetitle -->

	<section id="intWarranties" class="intWarranties paddingtop-30 paddingbot-70">	
           <div class="container">
				<div class="row">
                	<div class="col-sm-12 warrantyHeading wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s"><div id="pcMain" class="container-fluid">		
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
				<input type="hidden" id="email" name="email" value="<%=tmpEmail%>">
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
								url: "checkPass.asp",
								data: {
									passtype: "Rs",
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
					<button class="pcButton pcButtonContinue btn btn-skin btn-wc btn-contact" id="FormSubmit" name="FormSubmit">
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
					</button>                       
				</div>
      	</form>

    <% else %>

        <p>
            <%= dictLanguage.Item(Session("language")&"_newpass_17")%>
        </p>
      
        <div class="pcFormItem"><hr></div>
              
        <div class="pcFormButtons">
      	    <a class="pcButton pcButtonContinue btn btn-skin btn-wc btn-contact" href="<% if pIdCustomer>0 then %>custpref.asp<% else %>default.asp<% end if %>">
                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
            </a>
        </div>
        
    <% end if %>
  </div>
</div>
				</div>
		    </div>
    </section>	
    <!-- /Section: Welcome -->

<!--#include file="footer_wrapper.asp"-->
