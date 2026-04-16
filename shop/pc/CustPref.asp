<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="CustLIv.asp"-->
<%
'// START - Check for SSL and redirect to SSL login if not already on HTTPS
call storeSSLRedirect("1")
'// END - check for SSL
	
if session("customerCategory")<>0 then
	query="SELECT pcCC_Name, pcCC_Description FROM pcCustomerCategories WHERE idCustomerCategory="&session("customerCategory")&";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	strpcCC_Name=rs("pcCC_Name")
	strpcCC_Description=rs("pcCC_Description")
	SET rs=nothing
end if

' START - Retrieve customer name
if session("pcStrCustName") = "" OR session("pcStrCustEmail") = "" then
	pcIntCustomerId = session("idCustomer")
	if not validNum(pcIntCustomerId) then
		session("idCustomer") = Cdbl(0)
		response.Redirect("default.asp")
	end if	
	query = "SELECT name, lastName, email FROM customers WHERE idCustomer = " & pcIntCustomerId
	set rs = Server.CreateObject("ADODB.Recordset")
	set rs = conntemp.execute(query)
	pcStrCustName = rs("name") & " " & rs("lastName")
	session("pcStrCustName") = pcStrCustName
	pcStrCustEmail = rs("email")
	session("pcStrCustEmail") = pcStrCustEmail
	set rs = nothing	
	pEmail = pcStrCustEmail
else
	pcIntCustomerId = session("idCustomer")
	if not validNum(pcIntCustomerId) then
		session("idCustomer") = Cdbl(0)
		response.Redirect("default.asp")
	end if	
	query = "SELECT email FROM customers WHERE idCustomer = " & pcIntCustomerId
	set rs = Server.CreateObject("ADODB.Recordset")
	set rs = conntemp.execute(query)
	pEmail = rs("email")
	set rs = nothing
end if
' END - Retrieve customer name
%>
<!--#include file="header_wrapper.asp"-->
<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="Contact Us">Customer Service Area</h3>
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
                	<div class="col-sm-12 warrantyHeading wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s">
                    <div id="pcMain" class="pcCustPref">
	<div class="pcMainContent">
  	<% 'Erorr Message %>
		<%
			msg = Session("message")
			Session("message") = ""	
		%>
		<% If msg <> "" then %>
      <div class="pcErrorMessage"><%= msg %></div>
    <% end if %>
    
    <% 'Success Message %>
    <% If request.querystring("mode")="new" then %>
      <div class="pcSuccessMessage"><%= dictLanguage.Item(Session("language")&"_RegThankyou_1")%></div>
    <% end if %>
    
    <% 'Page Title %>
		<h1><%=(session("pcStrCustName") & " - " & dictLanguage.Item(Session("language")&"_CustPref_1"))%></h1>
    
    <% 'Customer Type %>
		<%if session("customerType")="1" then%>
      <p><%= dictLanguage.Item(Session("language")&"_CustPref_6")%></p>
    <%else%>
      <p><%= dictLanguage.Item(Session("language")&"_CustPref_7")%></p>
    <%end if%>
    <% if session("customerCategory")<>0 then%>
      <p><%= dictLanguage.Item(Session("language")&"_CustPref_15") & strpcCC_Name %></p>
    <%end if%>
    
    <div class="pcSpacer"></div>
    
    <% 'Welcome Message %>
		<p><%=(dictLanguage.Item(Session("language")&"_CustPref_10") & session("pcStrCustName") & "!")%></p>
    
		<ul>
			<% 'Start Shopping %>
      <li><a href="default.asp"><%= dictLanguage.Item(Session("language")&"_CustPref_9")%></a></li>
      
			<%
			'// GGG Add-on start		
			if scDisableGiftRegistry <> "1" then
			%>  
				<li><a href="ggg_manageGRs.asp"><%= dictLanguage.Item(Session("language")&"_CustPref_13")%></a></li>
			<%
			end if
			'//GGG Add-on end 
			%>
      
      <% 'View Past Orders %>
			<li><a href="CustviewPast.asp"><%= dictLanguage.Item(Session("language")&"_CustPref_8")%></a></li>
      
			<% 'Start Rewards Points %>
			<% If RewardsActive <> 0 AND session("customerType")<>"1" then %>
				<li><a href="CustRewards.asp"><%= dictRewardsLanguage.Item(Session("language")&"_CustPref_11")%><%=RewardsLabel%></a></li>
			<% End If %> 
			<% If RewardsActive <> 0 AND session("customerType")="1" AND RewardsIncludeWholesale=1 then %>
				<li><a href="CustRewards.asp"><%= dictRewardsLanguage.Item(Session("language")&"_CustPref_11")%><%=RewardsLabel%></a></li>
			<% End If %>
			<% 'End Reward Points %> 
      
      <% 'Modify Personal Info %>
			<li><a href="login.asp?lmode=1"><%= dictLanguage.Item(Session("language")&"_CustPref_3")%></a></li>
      
      <% 'Manage Shipping Addresses %>
			<li><a href="CustSAmanage.asp"><%= dictLanguage.Item(Session("language")&"_CustPref_11")%></a></li>
      
      <% 'View Saved Products %>
			<% If (scWL=-1) or ((scBTO=1) and (iBTOQuote=1)) then %>
				<li><a href="Custquotesview.asp"><%= dictLanguage.Item(Session("language")&"_CustPref_5")%></a></li>
			<% End If %>
      
      <% 'View Saved Carts %>
			<li><a href="CustSavedCarts.asp"><%= dictLanguage.Item(Session("language")&"_CustPref_16")%></a></li>
			
			<%
			'SB S 
			if scSBStatus="1" then
			%>
			<li><a href="sb_CustViewSubs.asp"><%= dictLanguage.Item(Session("language")&"_SB_3")%></a></li>
			<%
			end if
			'SB E 
			%>
      
			<%
				query="SELECT pcPay_EIG_Vault_ID, pcPay_EIG_Vault_CardNum, pcPay_EIG_Vault_CardExp FROM pcPay_EIG_Vault WHERE idCustomer="& Session("idCustomer") &""
				set rs=Server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)		
				if NOT rs.eof then
      	%>
        	<% 'View/Modify Payment Methods %>
        	<li><a href="CustviewPayment.asp"><%= dictLanguage.Item(Session("language")&"_EIG_10")%></a></li>
        <%
				end if
				set rs=nothing
			%>
      
      <% 'Contact Us %>
			<li><a href="contact.asp"><%= dictLanguage.Item(Session("language")&"_CustPref_12")%></a></li>
      
      <% 'Logout %>
			<li><a href="CustLO.asp"><%= dictLanguage.Item(Session("language")&"_CustPref_4")%></a></li>
		</ul>

		<% '// Account Consolidation %>
		<%  %>
		<!--#include file="opc_inc_CustConsolidate.asp"-->            
    <%
		'// START - Check Gift Certificate Balance
		'// Check to see if there are active Gift Certificates		
		Dim pcvIntGCExist
		pcvIntGCExist=0
		query="SELECT pcGO_GcCode FROM pcGCOrdered WHERE pcGO_Status = 1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)

			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			if not rs.eof then
				pcvIntGCExist=1
			end if
		set rs=nothing

		IF pcvIntGCExist<>0 THEN '// START - There are gift certificates
			pGiftCode = getUserInput(request.Form("pcGCcode"),100)
			IF pGiftCode<>"" THEN
				
				query="SELECT products.IDProduct, products.Description, pcGO_GcCode, pcGO_ExpDate, pcGO_Amount, pcGO_Status FROM Products,pcGCOrdered WHERE products.idproduct=pcGCOrdered.pcGO_idproduct AND pcGCOrdered.pcGO_GcCode='"& pGiftCode & "'"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
	
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
	
					IF NOT rs.eof THEN
					
						pcvGiftCertName=rs("Description")
						pcvGiftCertExp=rs("pcGO_ExpDate")
						pcIntGiftCertStatus=rs("pcGO_Status")
						
							if year(pcvGiftCertExp)="1900" then
								pcvGiftCertExp = dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_15")
							else
								if scDateFrmt="DD/MM/YY" then
									pcvGiftCertExp=day(pcvGiftCertExp) & "/" & month(pcvGiftCertExp) & "/" & year(pcvGiftCertExp)
								else
									pcvGiftCertExp=month(pcvGiftCertExp) & "/" & day(pcvGiftCertExp) & "/" & year(pcvGiftCertExp)
								end if
							If datediff("d", Now(), pcvGiftCertExp) <= 0 Then pcIntGiftCertExpired = 1
							end if
							
						pcvGiftCertAmount=rs("pcGO_Amount")
							if pcvGiftCertAmount<0 then pcvGiftCertAmount=0
	
						set rs = nothing
						%>
						<br /><br />
            <form name="checkGC2" action="" method="" class="pcForms">
							<h2><%=pcvGiftCertName%></h2>
							<div class="pcFormItem">
								<%
								if pcIntGiftCertStatus<>0 and pcIntGiftCertExpired<>1 then
									'// Gift Certificate is active
									%>
										<div class="pcFormItemFull">
											<%= dictLanguage.Item(Session("language")&"_CustPref_21") %>
										</div>
									<%
									else
									'// Gift Certificate is inactive
									%>
										<div class="pcFormItemFull">
											<img src="<%=pcf_getImagePath("images","pc_icon_error.png")%>" alt="<%=dictLanguage.Item(Session("language")&"_CustPref_22")%>">
											<%= dictLanguage.Item(Session("language")&"_CustPref_22") %>
										</div>
									<%
								end if
								%>
								<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_11")%><strong><%=pGiftCode%></strong></div>
								<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_16")%><%=pcvGiftCertExp%></div>
								<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_14")%><strong><%=scCurSign & money(pcvGiftCertAmount)%></strong></div>
							</div>
            </form>
						<%
					
					Else
					
						%>
						<br /><br />
						
						<div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_CustPref_18")%></div>                 	
						
			<%
					End If '// Retrieving information from the database
				Else
				%>
					<br /><br />
				<%
				End If '// GC Check form has been submitted
			'// END
			%>
	
			<form name="checkGC" action="custPref.asp" method="post" class="pcForms">
				<h2><%=dictLanguage.Item(Session("language")&"_CustPref_17")%></h2>
				<div class="pcFormItem">
					<div class="pcFormItemFull">
						<%=dictLanguage.Item(Session("language")&"_CustPref_19")%><input type="text" size="20" name="pcGCcode">
					</div>
				</div>
				<div class="pcSpacer"></div>
				<div class="pcFormButtons">
					<button class="pcButton pcButtonContinue" id="submit" name="submitGCcheck" value="<%=dictLanguage.Item(Session("language")&"_CustPref_20")%>">
						<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>" />
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
					</button>
				</div>
			</form>
            
        <%
		END IF '// END - There are gift certificates
		%>

	</div>
</div>
				</div>
		    </div>
    </section>	
    <!-- /Section: Welcome -->

<!--#include file="footer_wrapper.asp"-->
