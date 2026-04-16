<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="header_wrapper.asp"-->

	<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi">Customer Support Message</h3>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
	<!-- /Header: pagetitle -->
    
    <!-- Section: product-detail -->
    <section id="admin-list" class="paddingtop-60 paddingbot-40 blog-page">
		<div class="container">
			<div class="row">
				<div class="col-xs-12 col-sm-7 blog-listing">
					<!-- Post: Start -->
					<div class="admin-detail-single">
						<div class="admin-content admin-list-content">
                            <p>
        <%If Session("message")<>"" then
			msg = Session("message")
		Else
			msg=request.querystring("message")
		end if
		
		if msg<>"" then
			%>
			<%'= msg %>
<%
					'DA Edit to present nicer errors if possible
					boolNiceMsg = 0
					
					if not instr(msg,"3051") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The credit / debit card number has been entered incorrectly, please click 'Back to Payment Page' to return to the previous page and try again."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if

					if not instr(msg,"3120") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The first name field is required, please click 'Back to Payment Page' to return to the previous page and try again."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if
					
					if not instr(msg,"3048") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The credit / debit card number has been entered incorrectly, please click 'Back to Payment Page' to return to the previous page and try again."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if

					if not instr(msg,"3049") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The credit / debit card start date has been entered incorrectly, please click 'Back to Payment Page' to return to the previous page and try again."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if
										
					if not instr(msg,"4021") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The credit / debit card number entered is not accepted by our payment processor, please click 'Back to Payment Page' to return to the previous page and try again with a different credit / debit card."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if
					
					if not instr(msg,"5011") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The credit / debit card number has been entered incorrectly, please click 'Back to Payment Page' to return to the previous page and try again."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if
					
					if not instr(msg,"4022") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The credit / debit card number entered does not match the card type you selected from the drop down list, please click 'Back to Payment Page' to return to the previous page and try again."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if
					
					if not instr(msg,"3025") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The delivery post code entered is not in the correct format, please click 'Back to Payment Page' to return to the previous page and try again."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if

					if not instr(msg,"5038") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The delivery phone number entered is not in the correct format, please click 'Back to Payment Page' to return to the previous page and try again. Please try removing any non numeric characters."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if

					if not instr(msg,"3139") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The delivery address state / county entered is not in the correct format, please click 'Back to Payment Page' to return to the previous page and try again."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if

					if not instr(msg,"5055") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The post code entered is not in the correct format, please click 'Back to Payment Page' to return to the previous page and try again."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if

					if not instr(msg,"4026") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "The 3D Secure check has failed, we require all payments to pass the 3D Secure system, please click 'Back to Payment Page' to return to the previous page and try again."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if

					if not instr(msg,"4042") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "There has been an error with processing this order and payment. This can happen if the checkout process has taken a long time or if you are attempting to make multiple purchases over a short period of time. Please return to the 'Shoping Cart' page and try again."
						strNiceMsg = strNiceMSg & "<br /><br /><a href=""/shop/pc/viewCart.asp"" class=""btn product-action btn-skin pg-blue-btn"">Back to Shopping Cart</a>"
					end if

					if not instr(msg,"5036") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "There has been an error with processing this order and payment. This can happen if the checkout process has taken a long time or if you are attempting to make multiple purchases over a short period of time. Please return to the 'Shoping Cart' page and try again."
						strNiceMsg = strNiceMSg & "<br /><br /><a href=""/shop/pc/viewCart.asp"" class=""btn product-action btn-skin pg-blue-btn"">Back to Shopping Cart</a>"
					end if

					if not instr(msg,"2000") = 0 then
						boolNiceMsg = 1
						strNiceMsg = "Unfortunately this transaction has been declined by your bank. This can be due to either:<br /><br />- An error in the address or card details<br />- Insufficient funds in your account<br />- A bank security check<br /><br /> We recommend that you return to the payment page using the link below and check that all the details are correct before attempting the transaction again.<br /><br />If you have checked your details and the payment still fails then you should contact your bank directly to check whether this is a security check, if so they will be able to instantly clear this and allow you to process the transaction."
						strNiceMsg = strNiceMSg & "<br /><br />"
					end if

					'Fixes for current strings
					msg=getUserInput(request.querystring("message"),0)
					msg=replace(msg, "&lt;BR&gt;", "<BR>")
					msg=replace(msg, "&lt;br&gt;", "<br>")
					msg=replace(msg, "&lt;b&gt;", "<b>")
					msg=replace(msg, "&lt;/b&gt;", "</b>")
					msg=replace(msg, "&lt;/font&gt;", "</font>")
					msg=replace(msg, "&lt;a href", "<a href")
					msg=replace(msg, "&gt;Back&lt;/a&gt;", ">Back</a>")
					msg=replace(msg, "&lt;font", "<font")
					msg=replace(msg, "&gt;<b>Error&nbsp;</b>:", "><b>Error&nbsp;</b>:")
					msg=replace(msg, "&gt;&lt;img src=", "><img src=")
					msg=replace(msg, "&gt;&lt;/a&gt;", "></a>")
					msg=replace(msg, "&gt;<b>", "><b>")
					msg=replace(msg, "&lt;/a&gt;", "</a>")
					msg=replace(msg, "&gt;View Cart", ">View Cart")
					msg=replace(msg, "&gt;Continue", ">Continue")
					msg=replace(msg, "&lt;u>", "<u>")
					msg=replace(msg, "&lt;/u>", "</u>")
					msg=replace(msg, "&lt;ul&gt;", "<ul>")
					msg=replace(msg, "&lt;/ul&gt;", "</ul>")
					msg=replace(msg, "&lt;li&gt;", "<li>")
					msg=replace(msg, "&lt;/li&gt;", "</li>")
					msg=replace(msg, "&gt;", ">") 
					msg=replace(msg, "DAAND", "&") 

					
					if boolNiceMsg=1 then
					strNiceMsg = Replace(strNiceMsg,"DAAND","&")
					response.write(strNiceMsg)
					else
					response.write(msg)
					end if
					%></p>
   			<% If Request("back")="1" Then %>
				<br /><br /><br />
				<a href="<% If len(Session("backbuttonURL"))>0 Then %><%=Session("backbuttonURL") %><% Else %>javascript:history.go(-1)<% End If %>" class="btn product-action pg-green-btn">
					
					<span class="btn product-action pg-green-btn">Back To Payment Page</span>
				</a>
			<% End If
		Else%>
			<%=dictLanguage.Item(Session("language")&"_msg_note")%>
			<br /><br /><br />
			<a class="btn product-action pg-green-btn" href="/">
				<img src="<%=pcf_getImagePath("",RSlayout("continueshop"))%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_11")%>">
				<span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_continueshop")%></span>
			</a>&nbsp;
			<a href="<% If len(Session("backbuttonURL"))>0 Then %><%=Session("backbuttonURL") %><% Else %>javascript:history.go(-1)<% End If %>" class="btn product-action pg-green-btn">
				<span class="btn product-action pg-green-btn">Back To Payment Page</span>
			</a>
		<%End if%>					
				<a href="<% If len(Session("backbuttonURL"))>0 Then %><%=Session("backbuttonURL") %><% Else %>javascript:history.go(-1)<% End If %>" class="btn product-action pg-green-btn">
					
					<span class="btn product-action pg-green-btn">Back To Payment Page</span>
				</a>
						</div>
					</div>
					<!-- Post: End -->
				</div>
						</div>
					</div>
        
					</div>
				</div>
						</div>
			</div>
		</div>
	</section>
<%
Session("backbuttonURL") = ""
Session("message") = ""
%>
<!--#include file="footer_wrapper.asp"-->
