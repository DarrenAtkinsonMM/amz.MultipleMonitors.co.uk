<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
	<%
	on error resume next

	query="SELECT recalculate, continueshop, checkout, submit, morebtn, viewcartbtn, checkoutbtn, addtocart, addtowl, register, cancel, remove, add2, login, login_checkout, back, register_checkout"
	'CONFIGURATOR ADDON-S
	if scBTO=1 then
		query=query&", customize, [reconfigure], resetdefault, savequote,revorder,submitquote,pcLO_requestQuote"
	end if
	'CONFIGURATOR ADDON-E
	query=query&", ID,pcLO_placeOrder,pcLO_checkoutWR,pcLO_processShip,pcLO_finalShip,pcLO_backtoOrder,pcLO_Previous,pcLO_Next,CreRegistry,DelRegistry,AddToRegistry,UpdRegistry,SendMsgs,RetRegistry,pcLO_Update, pcLO_Savecart FROM layout WHERE (((ID)=2));"

	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)

	if err.number <> 0 then
		response.write "Error in AdminButtons: "&Err.Description
			set rstemp=nothing
	end If

	precalculate=rstemp("recalculate")
	pcontinueshop=rstemp("continueshop")
	pcheckout=rstemp("checkout")
	psubmit=rstemp("submit")
	pmorebtn=rstemp("morebtn")
	pviewcartbtn=rstemp("viewcartbtn")
	pcheckoutbtn=rstemp("checkoutbtn")
	paddtocart=rstemp("addtocart")
	paddtowl=rstemp("addtowl")
	pregister=rstemp("register")
	pcancel=rstemp("cancel")
	premove=rstemp("remove")
	padd2=rstemp("add2")
	plogin=rstemp("login")
	plogin_checkout=rstemp("login_checkout")
	pback=rstemp("back")
	pregister_checkout=rstemp("register_checkout")
	'CONFIGURATOR ADDON-S
	If scBTO=1 then
		pcustomize=rstemp("customize")
		preconfigure=rstemp("reconfigure")
		presetdefault=rstemp("resetdefault")
		psavequote=rstemp("savequote")
		prevorder=rstemp("revorder")
		psubmitquote=rstemp("submitquote")
		pcv_requestQuote=rstemp("pcLO_requestQuote")
	End If
	'CONFIGURATOR ADDON-E
	pcv_placeOrder=rstemp("pcLO_placeOrder")
	pcv_checkoutWR=rstemp("pcLO_checkoutWR")
	pcv_processShip=rstemp("pcLO_processShip")
	pcv_finalShip=rstemp("pcLO_finalShip")
	pcv_backtoOrder=rstemp("pcLO_backtoOrder")
	pcv_previous=rstemp("pcLO_Previous")
	pcv_next=rstemp("pcLO_Next")

	'GGG Add-on start

		pcrereg=rstemp("CreRegistry")
		pdelreg=rstemp("DelRegistry")
		paddreg=rstemp("AddToRegistry")
		pupdreg=rstemp("UpdRegistry")
		psendmsgs=rstemp("SendMsgs")
		pretreg=rstemp("RetRegistry")

	'GGG Add-on end

	yellowupd=rstemp("pcLO_Update")
	pcv_strSaveCart=rstemp("pcLO_Savecart")

	set rstemp=nothing
	%>

<!DOCTYPE html>
<html>
	<head>
		<title>Upload Store Buttons</title>
		<link href="../pc/css/pcStorefront.css" rel="stylesheet" type="text/css" />
		<link href="../pc/<%= scThemePath %>/css/theme.css" rel="stylesheet" type="text/css" />
		<!--#include file="inc_header.asp"-->
	</head>
	<body style="background-image: none; background-color: transparent; width: 100%; min-width: inherit;">

		<form method="post" enctype="multipart/form-data" action="buttonupl.asp" class="pcForms" target="_top">
			<table class="pcCPcontent" style="width: 100%">
			<tr>
				<td colspan="3" class="pcCPspacer">
					<!--#include file="pcv4_showMessage.asp"-->
				</td>
			</tr>
			<tr>
				<td colspan="3">
					<div class="bs-callout bs-callout-warning">
						<h4>Store Buttons & Themes</h4>
						<ul>
							<li>You are currently using the "<%= Replace(scThemePath, "theme/", "") %>" theme. We have created a preview of the theme's buttons below. To review/switch your current theme, visit the <a href="AdminSettings.asp?tab=4" target="_blank">display settings</a>.</li>
							<li id="cssButtonsDesc" style="display: none;">The theme you have selected on your store uses <strong>CSS Buttons</strong>. As a result, the image uploader is not available for this theme.</li>
							<li id="imageButtonsDesc" style="display: none;">The theme you have selected on your store uses <strong>Image Buttons</strong>. If you would like to customize any of these buttons, use the upload areas below.</a></li>
							<li id="hybridButtonsDesc" style="display: none;">The theme you have selected on your store uses a combination of <strong>Image Buttons</strong> and <strong>CSS Buttons</strong>. You may customize <strong>Image Buttons</strong> from this page, but you will need to edit the theme to customize <strong>CSS Buttons</strong>.</li>
							<li>For instructions on how to customize your theme, please consult the <a href="https://productcart.desk.com/customer/en/portal/articles/1543649-theming-guide" target="_blank">theme guide</a> on our Wiki.</li>
						</ul>
					 </div>
				</td>
			</tr>
			<tr>
				<th>Button Name</th>
				<th>Upload</th>
				<th style="width: 31%; text-align: center">Button Preview</th>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td width="25%">Small Add to Cart:</td>
				<td width="44%">
					<input class=ibtng type="file" name="add2" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonAddToCartSmall">
							<img src="../pc/<%=padd2%>" alt="<%= dictLanguage.Item(Session("language")&"_css_add2") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_add2") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Add To Cart:</td>
				<td width="44%">
					<input class=ibtng type="file" name="addtocart" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonAddToCart">
							<img src="../pc/<%=paddtocart%>" alt="<%= dictLanguage.Item(Session("language")&"_css_addtocart") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_addtocart") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">View Cart:</td>
				<td width="44%">
					<input class=ibtng type="file" name="viewcartbtn" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonViewCart">
							<img src="../pc/<%=pviewcartbtn%>" alt="<%= dictLanguage.Item(Session("language")&"_css_viewcartbtn") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_viewcartbtn") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Wish List:</td>
				<td width="44%">
					<input class=ibtng type="file" name="addtowl" size="30">
				</td>
				<td width="31%">
					<div align="center">						
						<a href="#" onclick="return false;" class="pcButton pcButtonAddToWishlist">
							<img src="../pc/<%=paddtowl%>" alt="<%= dictLanguage.Item(Session("language")&"_css_addtowl") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_addtowl") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Checkout:</td>
				<td width="44%">
					<input class=ibtng type="file" name="checkout" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonCheckout">
							<img src="../pc/<%=pcheckout%>" alt="<%= dictLanguage.Item(Session("language")&"_css_checkout") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_checkout") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Cancel: </td>
				<td width="44%">
					<input class=ibtng type="file" name="cancel" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonCancel">
							<img src="../pc/<%=pcancel%>" alt="<%= dictLanguage.Item(Session("language")&"_css_cancel") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_cancel") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Back:</td>
				<td width="44%">
					<input class=ibtng type="file" name="back" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonBack">
							<img src="../pc/<%=pback%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Continue Shopping:</td>
				<td width="44%">
					<input class=ibtng type="file" name="continueshop" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonContinueShopping">
							<img src="../pc/<%=pcontinueshop%>" alt="<%= dictLanguage.Item(Session("language")&"_css_continueshop") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_continueshop") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Small More Info: </td>
				<td width="44%">
					<input class=ibtng type="file" name="morebtn" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonMoreDetails">
							<img src="../pc/<%=pmorebtn%>" alt="<%= dictLanguage.Item(Session("language")&"_css_morebtn") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_morebtn") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Login:</td>
				<td width="44%">
					<input class=ibtng type="file" name="login" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonLogin">
							<img src="../pc/<%=plogin%>" alt="<%= dictLanguage.Item(Session("language")&"_css_login") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_login") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Recalculate:</td>
				<td width="44%">
					<input class=ibtng type="file" name="recalculate" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonRecalculate">
							<img src="../pc/<%=precalculate%>" alt="<%= dictLanguage.Item(Session("language")&"_css_recalculate") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_recalculate") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td>Register:</td>
				<td>
					<input class=ibtng type="file" name="register" size="30">
				</td>
				<td>
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonRegister">
							<img src="../pc/<%=pregister%>" alt="<%= dictLanguage.Item(Session("language")&"_css_register") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_register") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Remove from Cart: </td>
				<td width="44%">
					<input class=ibtng type="file" name="remove" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonRemove">
							<img src="../pc/<%=premove%>" alt="<%= dictLanguage.Item(Session("language")&"_css_remove") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_remove") %></span>
						</a>
					</div>
				</td>
			</tr>

			<% 'CONFIGURATOR ADDON-S
			If scBTO=1 then %>
			<tr>
				<td width="25%">Customize:</td>
				<td width="44%">
					<input class=ibtng type="file" name="customize" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonCustomize">
							<img src="../pc/<%=pcustomize%>" alt="<%= dictLanguage.Item(Session("language")&"_css_customize") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_customize") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Reconfigure:</td>
				<td width="44%">
					<input class=ibtng type="file" name="reconfigure" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<% 'NOT USED IN STOREFRONT %>
						<a href="#" onclick="return false;" class="pcButton pcButtonRecustomize">
							<img src="../pc/<%=preconfigure%>" alt="<%= dictLanguage.Item(Session("language")&"_css_reconfigure") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_reconfigure") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Reset to Default:</td>
				<td width="44%">
					<input class=ibtng type="file" name="resetdefault" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<% 'NOT USED IN STOREFRONT %>
						<a href="#" onclick="return false;" class="pcButton pcButtonResetDefault">
							<img src="../pc/<%=presetdefault%>" alt="<%= dictLanguage.Item(Session("language")&"_css_resetdefault") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_resetdefault") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Save Quote:</td>
				<td width="44%">
					<input class=ibtng type="file" name="savequote" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonSaveQuote">
							<img src="../pc/<%=psavequote%>" alt="<%= dictLanguage.Item(Session("language")&"_css_savequote") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_savequote") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Review &amp; Order:</td>
				<td width="44%">
					<input class=ibtng type="file" name="revorder" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<% 'NOT USED IN STOREFRONT %>
						<a href="#" onclick="return false;" class="pcButton pcButtonReviewOrder">
							<img src="../pc/<%=prevorder%>" alt="<%= dictLanguage.Item(Session("language")&"_css_revorder") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_revorder") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Submit Quote:</td>
				<td width="44%">
					<input class=ibtng type="file" name="submitquote" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<% 'NOT USED IN STOREFRONT %>
						<a href="#" onclick="return false;" class="pcButton pcButtonSubmitQuote">
							<img src="../pc/<%=psubmitquote%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submitquote") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submitquote") %></span>
						</a>
					</div>
				</td>
			</tr>
			<%'New Button for ProductCart v3%>
			<tr>
				<td width="25%">Request a Quote:</td>
				<td width="44%">
					<input class=ibtng type="file" name="pcv_requestQuote" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<% 'NOT USED IN STOREFRONT %>
						<a href="#" onclick="return false;" class="pcButton pcButtonRequestQuote">
							<img src="../pc/<%=pcv_requestQuote%>" alt="<%= dictLanguage.Item(Session("language")&"_css_pcLO_requestquote") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_requestquote") %></span>
						</a>
					</div>
				</td>
			</tr>
			<%'End of New Button for ProductCart v3%>
			<% End If
			'CONFIGURATOR ADDON-E %>
			<%'New Button for ProductCart v3%>
			<tr>
				<td width="25%">Place Order:</td>
				<td width="44%">
					<input class=ibtng type="file" name="pcv_placeOrder" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonPlaceOrder">
							<img src="../pc/<%=pcv_placeOrder%>" alt="<%= dictLanguage.Item(Session("language")&"_css_pcLO_placeorder") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_placeorder") %></span>
						</a>
					</div>
				</td>
			</tr>
			<%'End of New Button for ProductCart v3%>
			<%'New Button for ProductCart v3%>
			<!--<tr>
				<td width="25%">Checkout Without Registering:</td>
				<td width="44%">-->
					<input class=ibtng type="hidden" name="pcv_checkoutWR" size="30">
			<!--</td>
				<td width="31%">
					<div align="center"><img src="../pc/<%=pcv_checkoutWR%>"></div>
				</td>
			</tr>-->
			<%'End of New Button for ProductCart v3%>
			<%'New Button for ProductCart v3%>
			<tr>
				<td width="25%">Process Shipment:</td>
				<td width="44%">
					<input class=ibtng type="file" name="pcv_processShip" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<% 'NOT USED IN STOREFRONT %>
						<a href="#" onclick="return false;" class="pcButton pcButtonProcessShipment">
							<img src="../pc/<%=pcv_processShip%>" alt="<%= dictLanguage.Item(Session("language")&"_css_pcLO_processship") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_processship") %></span>
						</a>
					</div>
				</td>
			</tr>
			<%'End of New Button for ProductCart v3%>
			<%'New Button for ProductCart v3%>
			<tr>
				<td width="25%">Finalize Shipment:</td>
				<td width="44%">
					<input class=ibtng type="file" name="pcv_finalShip" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<% 'NOT USED IN STOREFRONT %>
						<a href="#" onclick="return false;" class="pcButton pcButtonFinalizeShipment">
							<img src="../pc/<%=pcv_finalShip%>" alt="<%= dictLanguage.Item(Session("language")&"_css_pcLO_finalship") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_finalship") %></span>
						</a>
					</div>
				</td>
			</tr>
			<%'End of New Button for ProductCart v3%>
			<%'New Button for ProductCart v3%>
			<tr>
				<td width="25%">Back to Order Details:</td>
				<td width="44%">
					<input class=ibtng type="file" name="pcv_backtoOrder" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonBackToOrder">
							<img src="../pc/<%=pcv_backtoOrder%>" alt="<%= dictLanguage.Item(Session("language")&"_css_pcLO_backtoorder") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_backtoorder") %></span>
						</a>
					</div>
				</td>
			</tr>
			<%'End of New Button for ProductCart v3%>
			<%'New Button for ProductCart v3%>
			<tr>
				<td width="25%">Previous:</td>
				<td width="44%">
					<input class=ibtng type="file" name="pcv_previous" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonPrevious">
							<img src="../pc/<%=pcv_previous%>" alt="<%= dictLanguage.Item(Session("language")&"_css_pcLO_previous") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_previous") %></span>
						</a>
					</div>
				</td>
			</tr>
			<%'End of New Button for ProductCart v3%>
			<%'New Button for ProductCart v3%>
			<tr>
				<td width="25%">Next:</td>
				<td width="44%">
					<input class=ibtng type="file" name="pcv_next" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonNext">
							<img src="../pc/<%=pcv_next%>" alt="<%= dictLanguage.Item(Session("language")&"_css_pcLO_next") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_next") %></span>
						</a>
					</div>
				</td>
			</tr>
			<%'End of New Button for ProductCart v3%>
			<%'GGG Add-on start%>
			<tr>
				<td width="25%">Create New Registry:</td>
				<td width="44%">
					<input class=ibtng type="file" name="crereg" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonCreateRegistry">
							<img src="../pc/<%=pcrereg%>" alt="<%= dictLanguage.Item(Session("language")&"_css_creregistry") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_creregistry") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Delete Registry:</td>
				<td width="44%">
					<input class=ibtng type="file" name="delreg" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonDeleteRegistry">
							<img src="../pc/<%=pdelreg%>" alt="<%= dictLanguage.Item(Session("language")&"_css_delregistry") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_delregistry") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Add To Registry:</td>
				<td width="44%">
					<input class=ibtng type="file" name="addreg" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonAddToRegistry">
							<img src="../pc/<%=paddreg%>" alt="<%= dictLanguage.Item(Session("language")&"_css_addtoregistry") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_addtoregistry") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Update Registry:</td>
				<td width="44%">
					<input class=ibtng type="file" name="updreg" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonUpdateRegistry">
							<img src="../pc/<%=pupdreg%>" alt="<%= dictLanguage.Item(Session("language")&"_css_updregistry") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_updregistry") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Send Messages:</td>
				<td width="44%">
					<input class=ibtng type="file" name="sendmsgs" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonSendMessages">
							<img src="../pc/<%=psendmsgs%>" alt="<%= dictLanguage.Item(Session("language")&"_css_sendmsgs") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_sendmsgs") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%">Return to Registry:</td>
				<td width="44%">
					<input class=ibtng type="file" name="retreg" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonReturnToRegistry">
							<img src="../pc/<%=pretreg%>" alt="<%= dictLanguage.Item(Session("language")&"_css_retregistry") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_retregistry") %></span>
						</a>
					</div>
				</td>
			</tr>
			<%'GGG Add-on end%>
			<tr>
				<td width="25%" nowrap>Continue/Next Step/Update:<br><span class="pcSmallText">Used on One Page Checkout</span></td>
				<td width="44%">
					<input class=ibtng type="file" name="yellowupd" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonContinue">
							<img src="../pc/<%=yellowupd%>" alt="<%= dictLanguage.Item(Session("language")&"_css_pcLO_update") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_update") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td width="25%" nowrap>Save Cart:<br><span class="pcSmallText">Used on View Cart</span></td>
				<td width="44%">
					<input class=ibtng type="file" name="savecart" size="30">
				</td>
				<td width="31%">
					<div align="center">
						<a href="#" onclick="return false;" class="pcButton pcButtonSaveCart">
							<img src="../pc/<%=pcv_strSaveCart%>" alt="<%= dictLanguage.Item(Session("language")&"_css_pcLO_savecart") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_savecart") %></span>
						</a>
					</div>
				</td>
			</tr>
			<tr>
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="3"><hr></td>
			</tr>
			<tr>
				<td colspan="3" align="center">
					<script type="text/javascript">
						var imageButtonCount = 0;
						var cssButtonCount = 0;

						$pc(window).on('load', function() {
							$pc(".pcButton").each(function () {
								var imageVisible = $pc(this).find("img").is(":visible");

								if (!imageVisible) {
									var upload = $pc(this).parents("tr").find("input[type='file']")
									upload.hide();

									upload.parent().append("Visit theme guide for editing instructions.");

									cssButtonCount++;
								} else {
									imageButtonCount++;
								}
							});

							if (imageButtonCount < 1) {
								$pc("#UpdateButton").hide();
								$pc("#DefaultSettingsButton").hide();
							}

							if (cssButtonCount > 0 && imageButtonCount > 0) {
								$pc("#hybridButtonsDesc").show();
							} else if (cssButtonCount > 0) {
								$pc("#cssButtonsDesc").show();
							} else if (imageButtonCount > 0) {
								$pc("#imageButtonsDesc").show();
							}
						});
					</script>
					<button id="UpdateButton" name="submit" class="btn btn-primary">Update</button>
					&nbsp;
					<button id="DefaultSettingsButton" name="default" class="btn btn-default"  onClick="parent.location.href = 'setBtnDefault.asp'">Set back to default settings</button>
					&nbsp;
					<button class="btn btn-default" name="Button" onClick="javascript: history.back()">Back</button>
				</td>
			</tr>
			</table>
		</form>
	</body>
</html>
<% call closeDb() %>
