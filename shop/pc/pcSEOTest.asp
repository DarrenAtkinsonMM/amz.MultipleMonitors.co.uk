<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="pcStartSession.asp" -->
<!--#include file="header_wrapper.asp"-->

<div id="pcMain" class="container">
    <div class="row">
        <div class="col-sm-12">
            <h1>SEO Test</h1>
            <div class="well">
                When you click the button below  you will be directed toward an SEO Friendly URL. If the page loads, then your server is setup correctly. If the page does not load, or displays a generic 404 error, then your server is not setup for SEO URLs.
                <br />
                <br />
                <a href="SEOCheck-f0.htm" class="btn btn-info" role="button">Begin SEO Test</a>
            </div>            
        </div>
    </div>
</div>

<!--#include file="footer_wrapper.asp"-->
