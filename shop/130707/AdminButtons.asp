<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Edit Store Buttons" %>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->

<script type="text/javascript">
	function autoResize(id) {
		var newheight;
		var newwidth;

		if (document.getElementById) {
			newheight = document.getElementById(id).contentWindow.document.body.scrollHeight;
			newwidth = document.getElementById(id).contentWindow.document.body.scrollWidth;
		}

		document.getElementById(id).height = (newheight) + "px";
		document.getElementById(id).width = (newwidth) + "px";
	}
</script>

<iframe src="AdminButtonsForm.asp" id="buttonsFrame" style="border: none; width: 100%;" onload="autoResize('buttonsFrame');">
	You must enable Javascript to view and edit store buttons!
</iframe>

<!--#include file="AdminFooter.asp"-->
