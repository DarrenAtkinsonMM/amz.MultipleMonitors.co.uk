<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Upload Images"
pageIcon="pcv4_icon_upload.png"
Section="products" 
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<% ' START show message, if any %>
<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<% pcv_intMaxUploads = 6 %>

<form method="post" enctype="multipart/form-data" name="hForm" action="imageupl.asp" class="pcForms">
	<input type="hidden" name="smallImageUrl" value="">
    <table class="pcCPcontent" style="width:auto;">
        <tr>
            <td colspan="2">
                All images are uploaded to the &quot;<b><%=scPcFolder%>/pc/catalog</b>&quot; folder on your Web server (<a href="javascript:chgWin('../pc/imageDir.asp?ffid=smallImageUrl&fid=hForm&ref=ImageUpload','window2')">Manage Uploaded Images</a>).<br>Of course, you may also use your favorite FTP program to upload images to the same location.
            </td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <% For i = 1 To pcv_intMaxUploads %>
        <tr>
            <td width="20%" align="right">Image <%= i %>:</td>
            <td width="80%"><input class="ibtng" type="file" name="image_<%= i %>" size="30"></td>
        </tr>
        <% Next %>
        <tr>
            <td colspan="2" class="pcCPspacer"><hr></td>
        </tr>
        <tr> 
            <td colspan="2" align="center"> 
                <input type="submit" name="Submit" value="Upload" class="btn btn-primary">&nbsp;
                <input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
            </td>
        </tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->