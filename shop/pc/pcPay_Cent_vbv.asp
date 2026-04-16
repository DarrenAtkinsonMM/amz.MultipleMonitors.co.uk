<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<% Response.Buffer=True%> 
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->

<HTML>
	<HEAD>
		<TITLE>Verified by Visa</TITLE>
		<link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath("css","pcStorefront.css")%>" />
        <link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath("css","bootstrap.css")%>" />
	</HEAD>
<body>
<div id="pcMain">
	<div class="container-fluid">
	    <div class="row">
            <div class="col-xs-9">
                <h2>Same card, added safety online.</h2>
            </div>             
            <div class="col-xs-3">
		        <IMG height=84 src="<%=pcf_getImagePath("images","pc_logo_vbv.gif")%>" width=143 border=0>
            </div>
        </div>
	    <div class="row">
            <div class="col-xs-12">
                <p>Verified by Visa is a new service that lets you shop online with added confidence.
                Through a simple checkout process, Verified by Visa confirms your identity when you make purchases at participating online stores. It's convenient, and it works with your existing Visa card.</p>
                <p>Plus, Verified by Visa is a snap to use. You register your card just once, and create your own password. Then, when you make purchases at participating online stores, a Verified by Visa window will appear. Simply enter your password and click submit. Your identity is verified and the purchase is secure.</p>
                <p>To activate Verified by Visa in your Visa card, or to learn more, contact the financial institution that issued your Visa card or visit <A href="http://www.visa.com/verified" target=_blank>www.visa.com</A>.</p>
            </div>
        </div>
        <div class="row"> 
            <div class="col-xs-11"></div>
            <div class="col-xs-1"><a href="javascript:window.close();"><img src="<%=pcf_getImagePath("images","close.gif")%>" alt="Close" name="button" width="32" height="25" hspace=0 border=0></a></div>
        </div>        
    </div>
</div>
</body>
</html>
<% call closeDb() %>
