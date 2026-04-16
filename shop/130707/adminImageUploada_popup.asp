<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->

<%
pcv_intMaxUploads = 6
%>

<html>
<title>Upload Images</title>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" enctype="multipart/form-data" action="adminimageupl_popup.asp">
    <table width="400" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
        <tr> 
            <td> 
                <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
                    <tr> 
                    	<td colspan="3" bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Upload
                      Images</font></b></font></td>
                    </tr>
                    <tr> 
                    	<td height="10" colspan="2"></td>
                    </tr>
                    <tr> 
                    	<td height="18" colspan="2"><font face="Arial, Helvetica, sans-serif" size="2"> 
                          Select an image using the &quot;Browse&quot; button, then 
                          click on &quot;Upload&quot;. All images are automatically uploaded
                          to the &quot;<b>/store/pc/images</b>&quot; subfolder.</font>
                        </td>
                    </tr>
                    <tr> 
                    	<td colspan="2" height="10"></td>
                    </tr>
                    
                    <% For i = 1 To pcv_intMaxUploads %>
                    <tr> 
                        <td width="20%">
                        	<div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Image <%= i %>: </font></div>
                        </td>
                    	<td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2">
                      		<input class="ibtng" type="file" name="image_<%= i %>" size="25">
                      		</font></b>
                        </td>
                    </tr>
                    <% Next %>
                    
                    <tr> 
                    	<td colspan="2" height="15"></td>
                    </tr>
                    <tr> 
                        <td colspan="2"> 
                        	<div align="left"> 
                            	<p><font face="Arial, Helvetica, sans-serif" size="2"> 
                              	<input type="submit" name="Submit" value="Upload">
                              	<input type="button" class="btn btn-default"  value="Close Window" onClick="javascript:window.close();">
                              	</font></p>
                          	</div>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</form>
</body>
</html>
<% call closeDb() %>