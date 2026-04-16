<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%pcStrPageName="msg.asp"%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
ButtonType = "LINK"

FUNCTION GetButtonLink(buttonname)
	If ButtonType = "LINK" Then
		select case buttonname
			case "back"
				tmpButtonName = "Back"
			case "continueshop"
				tmpButtonName = "Continue Shopping"
		end select
		GetButtonLink = tmpButtonName
	Else
		GetButtonLink = "<img src="""& pcf_getImagePath("",rslayout(""&tmpButtonName&"")) & """>"
	End If
End Function
%>
<%
	msg=request.querystring("message")
	'Check that msg is a number
	if not validNum(msg) then
			msg = 0
			response.write dictLanguage.Item(Session("language")&"_techErr_1")
	end if
on error resume next

Dim pcStrClass
pcStrClass="pcErrorMessage"			

select case msg
case 1
pcStrClass="pcInfoMessage"
end select				
%>
<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="Contact Us">Customer Message</h3>
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
<div id="pcMain" class="pcMsg">
  <div class="pcMainContent">
    <div class="<%=pcStrClass%>">
			<%
				response.write pcf_getStoreMsg(msg)
			%>
		</div>
  </div>
</div>
				</div>
		    </div>
    </section>	
    <!-- /Section: Welcome -->
<!--#include file="footer_wrapper.asp"-->
