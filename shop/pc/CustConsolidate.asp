<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%'Allow Guest Account
AllowGuestAccess=1
%>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<%
err.number=0
dim pIdOrder
%>
<!--#include file="header_wrapper.asp"-->
<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="Contact Us">Customer Service Message</h3>
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
                    <div id="pcMain">
	<div class="pcMainContent"> 
		<h1>
			<%=dictLanguage.Item(Session("language")&"_opc_cons_title")%>
		</h1>
		<%
		query = "SELECT email FROM customers WHERE idCustomer = " & Session("idCustomer")
		set rs = Server.CreateObject("ADODB.Recordset")
		set rs = conntemp.execute(query)
		If Not rs.EOF Then
			pEmail = rs("email")
		End If 
		set rs = nothing
		%>
		<!--#include file="opc_inc_CustConsolidate.asp"-->
	</div>
</div>
				</div>
		    </div>
    </section>	
    <!-- /Section: Welcome -->

<!--#include file="footer_wrapper.asp"-->
