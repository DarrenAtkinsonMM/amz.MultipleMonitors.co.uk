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
					topmenuArrays = " class=""active"""
				else
					topmenuBundles = " class=""active"""
				end if
			Case "stand"
				if request.querystring("arr")="1" Then
					topmenuArrays = " class=""active"""
				elseif request.querystring("sid")<>"" then
					topmenuBundles = " class=""active"""
				else
					topmenuStands = " class=""active"""
				end if
			Case "computer", "traderpc", "charterpc", "traderpropc"
				if request.querystring("sid")="" Then
					topmenuComputers = " class=""active"""
				else
					topmenuBundles = " class=""active"""
				end if
		End Select
	Case "/shop/pc/customcat-stands.asp"
		topmenuStands = " class=""active"""
	Case "/shop/pc/customcat-computers.asp","/shop/pc/customcat-tradingcomputers.asp"
		topmenuComputers = " class=""active"""
	Case "/shop/pc/customcat-bundles1.asp","/shop/pc/customcat-bundles2.asp","/shop/pc/customcat-bundles3.asp"
		topmenuBundles = " class=""active"""
	Case "/shop/pc/customcat-arrays1.asp","/shop/pc/customcat-arrays2.asp","/shop/pc/customcat-arrays3.asp"
		topmenuArrays = " class=""active"""
	Case "/default.asp"
		topmenuHome = " class=""active"""
	Case "/new-blog/default.asp","/shop/pc/viewcontent.asp"
		topmenuBlog = " class=""active"""
End Select
%>


<div id="wrapper">
    <nav class="navbar navbar-custom navbar-fixed-top" role="navigation">
		<div class="top-area">
			<div class="container">
				<div class="row">
					<div class="col-sm-6 col-md-6 topbar-connects">
					<p class="text-left">
						<span class="tb-contact-bx tb-phone"><a class="text-white" href="tel:03302236655"><i class="fa fa-phone"></i>0330 223 66 55</a></span>
						<a class="tb-contact-bx tb-mail" href="mailto:sales@multiplemonitors.co.uk"><i class="fa fa-envelope"></i>sales@multiplemonitors.co.uk</a>
					</p>
					</div>
					<div class="col-sm-6 col-md-6 text-right top-user-box">
                        <!-- #include file="smallQuickCart.asp" -->
					</div>
				</div>
			</div>
		</div>
        <div class="container navigation">
			<div class="row">
				<div class="navbar-header page-scroll col-lg-4 col-md-3 col-xs-12">
					<button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-main-collapse">
						<i class="fa fa-bars"></i>
					</button>
					<a class="navbar-brand" href="/">
						<img src="/images/logo.png" alt="" width="353" height="44" />
					</a>
				</div>
			<!-- Collect the nav links, forms, and other content for toggling -->
				<div class="collapse navbar-collapse navbar-right navbar-main-collapse">
				  <ul class="nav navbar-nav">
					<li<%=topmenuHome%>><a href="/">Home</a></li>
					<li<%=topmenuComputers%>><a href="/computers/">Computers</a></li>
					<li<%=topmenuArrays%>><a href="/display-systems/">Monitor Arrays</a></li>
					<li<%=topmenuBundles%>><a href="/bundles/">Bundles</a></li>
					<li<%=topmenuStands%>><a href="/stands/">Stands</a></li>
					<li<%=topmenuBlog%>><a href="/blog/">Blog</a></li>
				  </ul>
				</div>
				<!-- /.navbar-collapse -->
 </div>
        </div>
        <!-- /.container -->
		<div class="top-tagline-bar">
			<div class="container top-tagline">
				<div class="row">
					<p class="tagline-txt col-lg-8 col-sm-12">Multi-screen displays & computers. As seen on BBC's 'Traders: Millions by the Minute'</p>
					<p class="subscribe-txt col-lg-4 hidden-ss color">Get Exclusive Special Offers! <a href="/pages/email-signup/">SIGN UP NOW</a></p>
				</div>
				<!-- /.container -->
			</div>
		</div>
    </nav>
<%
	if scURLredirect = "" then
		homepageurl = scStoreURL & "/" & scPcFolder & "/home.asp"
	else
		homepageurl = scURLredirect
	end if
%>