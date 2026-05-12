<!DOCTYPE html>
<html lang="en" <%if lcase(pcStrPageName)="viewprd.asp" then%>prefix="og: http://ogp.me/ns#"<%end if%>>
<head itemscope itemtype="http://schema.org/WebSite">
<%
'// Help prevent XSS attacks
If session("Facebook")<>"1" Then
    Response.AddHeader "x-frame-options","DENY"
End If
If scMobileOn = "1" Then
    Response.AddHeader "Vary","User-Agent"
End If
%>
<!--#include file="inc_headerV5.asp" -->

<% If Len(Session("scFaviconTag")) > 0 Then %>
    <%=Session("scFaviconTag") %>
<% End If %>

<% If session("Mobile")="1" Then %>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <%=Session("scViewPortMobile") %>
<% Else %>
    <%=Session("scViewPort") %>
<% End If %>
<%
'// Layout Icons
viewcartbtn = RSlayout("viewcartbtn")
pcv_strRequiredIcon = rsIconObj("requiredicon")
pcv_strErrorIcon = rsIconObj("errorfieldicon")
%>
<!--#include file="inc_headerDAJS.asp" -->
<% if scGAType="2" then %>
<!-- Google Tag Manager -->
<noscript><iframe src="//www.googletagmanager.com/ns.html?id=<%=scGoogleTagManager%>"
height="0" width="0" style="display:none;visibility:hidden"></iframe></noscript>
<script>(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
j=d.createElement(s),dl=l!='dataLayer'?'&l='+l:'';j.async=true;j.src=
'//www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
})(window,document,'script','dataLayer','<%=scGoogleTagManager%>');</script>
<!-- End Google Tag Manager -->
<% end if %>
<% '// START: DO NOT REMOVE THIS LINE %>
<div id="pcMainService" data-ng-controller="serviceCtrl"></div> 
<% '// END: DO NOT REMOVE THIS LINE %>
<%
'-- DA - EDIT
'Code to figure out which menu item is active
'Response.write(Request.ServerVariables("SCRIPT_NAME"))
Select Case lcase(Request.ServerVariables("SCRIPT_NAME"))
	Case "/shop/pc/viewprd.asp"
		'Determine which product page type we are on
		Select Case pcv_strViewPrdStyle
			Case "monitor"
				if request.querystring("arr")="1" Then
					topmenuArrays = " class=""is-trader"""
				else
					topmenuBundles = " class=""is-trader"""
				end if
			Case "stand"
				if request.querystring("arr")="1" Then
					topmenuArrays = " class=""is-trader"""
				elseif request.querystring("sid")<>"" then
					topmenuBundles = " class=""is-trader"""
				else
					topmenuStands = " class=""is-trader"""
				end if
			Case "computer", "traderpc", "charterpc", "traderpropc"
				if request.querystring("sid")="" Then
					topmenuComputers = " class=""is-trader"""
				else
					topmenuBundles = " class=""is-trader"""
				end if
		End Select
	Case "/shop/pc/customcat-stands.asp"
		topmenuStands = " class=""is-trader"""
	Case "/shop/pc/customcat-computers.asp","/shop/pc/customcat-tradingcomputers.asp","/shop/pc/viewprd-traderpc.asp"
		topmenuComputers = " class=""is-trader"""
	Case "/shop/pc/bundlebuilder.asp","/shop/pc/viewprd-traderpc-bundle.asp"
		topmenuBundles = " class=""is-trader"""
	Case "/shop/pc/customcat-arrays1.asp","/shop/pc/customcat-arrays2.asp","/shop/pc/customcat-arrays3.asp"
		topmenuArrays = " class=""is-trader"""
	Case "/default.asp"
		topmenuHome = " class=""is-trader"""
	Case "/new-blog/default.asp","/shop/pc/viewcontent.asp"
		topmenuBlog = " class=""is-trader"""
End Select
%>


<div id="wrapper">
	<div class="site-header">
		<div class="topbar">
			<div class="mm-container inner">
				<div>
					<a href="tel:03302236655"><i class="fa fa-phone"></i>0330 223 66 55</a>
					<span class="sep hide-xs">|</span>
					<a class="hide-xs" href="mailto:sales@multiplemonitors.co.uk"><i class="fa fa-envelope-o"></i>sales@multiplemonitors.co.uk</a>
				</div>
				<!--#include file="smallQuickCart.asp"-->
			</div>
		</div>

		<div class="navwrap">
			<div class="mm-container nav-inner">
				<a href="/" class="brand-mark" aria-label="Multiple Monitors home">
					<img src="/images/mm-logo-trans.png" alt="Multiple Monitors Ltd" />
					<span class="brand-est">Est 2008<b>UK Specialist</b></span>
				</a>

				<nav class="mainnav">
					<a href="/"<%=topmenuHome%>>Home</a>
					<a href="/computers/"<%=topmenuComputers%>>Computers</a>
					<a href="/bundles/"<%=topmenuBundles%>>Bundles</a>
					<a href="/stands/"<%=topmenuStands%>>Stands</a>
					<a href="/display-systems/"<%=topmenuArrays%>>Monitor Arrays</a>
				</nav>

				<div class="nav-actions">
					<div class="nav-cta">
						<!--#include file="smallCartButton.asp"-->
					</div>
					<button class="nav-toggle" aria-label="Open menu" onclick="document.getElementById('mobnav').classList.toggle('is-open')"><i class="fa fa-bars"></i></button>
				</div>
			</div>
			<div class="mobnav" id="mobnav">
				<a href="/"<%=topmenuHome%>>Home</a>
				<a href="/computers/"<%=topmenuComputers%>>Computers</a>
				<a href="/bundles/"<%=topmenuBundles%>>Bundles</a>
				<a href="/stands/"<%=topmenuStands%>>Stands</a>
				<a href="/display-systems/"<%=topmenuArrays%>>Monitor Arrays</a>
				<a href="/shop/pc/custPref.asp">Existing Customer Login</a>
			</div>
		</div>
	</div><!-- /.site-header -->
<%
	if scURLredirect = "" then
		homepageurl = scStoreURL & "/" & scPcFolder & "/home.asp"
	else
		homepageurl = scURLredirect
	end if
%>