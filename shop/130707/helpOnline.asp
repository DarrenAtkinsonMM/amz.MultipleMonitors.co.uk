
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

	on error resume next
	' Retrieve Help Tooltip ID
	pcIntHelpTip = cint(request("ref"))
	pcAjaxTip = request("ajax")
	
	' Define Help Messages %>
<!--#include file="helpOnlineMessages.asp"-->

<% If pcAjaxTip = "1" Then %>
	<h3 class="pcCPhelpTitle"><%=pcStrTitle%></h3>
  <div class="pcCPhelpDesc">
  	<%
			Dim pcStrDetailsArr : pcStrDetailsArr=split(pcStrDetails, "<br><br>")
			
			If UBound(pcStrDetailsArr) > 2 Then
				Response.Write pcStrDetailsArr(0)
				Response.Write "<br><br>"
				Response.Write pcStrDetailsArr(1)
			Else
				Response.Write pcStrDetails
			End If
		%>
  </div>
  <p>Click for more details.</p>
<% Else %>
  <!DOCTYPE html>
	<html>
  <head>
  <title>ProductCart shopping cart software - Control Panel - Online Help</title>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta name="description" content="ProductCart asp shopping cart software is published by NetSource Commerce. ProductCart's Control Panel allows you to manage every aspect of your ecommerce store. For more information and for technical support, please visit NetSource Commerce at http://www.productcart.com">
  <style>
    ul {
      padding-top: 5px;
    }
    li {
      padding-bottom: 5px;
    }
</style>
  </head>
  <body style="font-family: Arial, sans-serif;">
    <table cellpadding="4" cellspacing="2" align="center" style="border: 1px solid #e1e1e1;">
      <tr>
        <td><h1 style="background-color:#e1e1e1; font-size:18px; padding:4px; margin:0;">Online Help&nbsp;&nbsp;<a href="javascript:print()"><img src="images/print_small.gif" alt="Print" border="0"></a></h1></td>
      </tr>
      <tr>
        <td><strong><span style="border-bottom: 1px solid #e1e1e1;"><%=pcStrTitle%></span></strong></td>
      </tr>
      <tr>
        <td style="font-size:12px;"><%=pcStrDetails%></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td style="font-size:12px; color:#666666">For more information on this and all other features, we encourage you to visit our constantly updated, WIKI-style <a href="http://wiki.productcart.com/" target="_blank">User Guide</a> or consult our <a href="help.asp" target="_blank">other support resources</a>.</td>
      </tr>
      <tr>
        <td align="right"><input type="button" class="btn btn-default"  onClick="self.close()" value="Close"></td>
      </tr>
    </table>
  </body>
  </html>
<% End If %>
