<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
pageTitle = "ProductCart Twitter Updates - Latest Tweets"
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
    <tr>
    	<td>
            <style type="text/css">
                iframe[id^='twitter-widget-']{ width:100% !important;}
            </style>
            <a class="twitter-timeline"  href="https://twitter.com/productcart"  data-widget-id="388668860498325504" width="520" height="600">Tweets by @productcart</a>
            <script type=text/javascript>
                try {
                    !function (d, s, id) { var js, fjs = d.getElementsByTagName(s)[0], p = /^http:/.test(d.location) ? 'http' : 'https'; if (!d.getElementById(id)) { js = d.createElement(s); js.id = id; js.src = p + "://platform.twitter.com/widgets.js"; fjs.parentNode.insertBefore(js, fjs); } }(document, "script", "twitter-wjs");
                } catch (err) { }
            </script>
        </td>
    </tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->
