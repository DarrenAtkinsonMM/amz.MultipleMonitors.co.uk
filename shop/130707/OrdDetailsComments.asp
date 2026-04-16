<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Dim pageTitle, Section
pageTitle="Order Details - Administrator Comments"
Section="orders" %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../pc/checkdate.asp" -->
<!--#include file="AdminHeader.asp" -->
<!--#include file="../htmleditor/editor.asp"-->
<%
Dim intIdOrder

'// Get order ID
intIdOrder=getUserInput(request("IDOrder"),0)
if not validNum(intIdOrder) or intIdOrder="0" then
   call closeDb()
    response.redirect "msg.asp?message=45"
end if

'Update or Retrieve Admin Comments
if (request("action")="add") then
	adminComments=replace(request("adminComments"),"'","''")
	adminComments=replace(adminComments,"&nbsp;"," ")
	adminComments=replace(adminComments,"  "," ")
		
	query="UPDATE orders SET adminComments=N'"&adminComments&"' WHERE idOrder="& intIdOrder
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)	
	
	call closeDb()
response.redirect "Orddetails.asp?id="&intIdOrder&"&s=1&msg="&Server.URLEncode("Administrator comments updated successfully.")
else
		
	query="SELECT adminComments FROM orders WHERE idOrder="& intIdOrder
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=conntemp.execute(query)	
	padminComments=rs("adminComments")
		
end if
%>
<script type=text/javascript>
function Form1_Validator(theForm)
{
	// InnovaStudio HTML Editor Workaround for this keyword
  theForm = document.hForm;

			if (theForm.adminComments.value == "")
 	{
		    alert("Please enter Administrator Comments for this order.");
		    theForm.Details.focus();
		    return (false);
	}
return (true);
}
function newWindow(file,window) {
		msgWindow=open(file,window,'resizable=no,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
}
</script>

<form name="hForm" method="post" action="ordDetailsComments.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<input type="hidden" value="<%=intIdOrder%>" name="IDOrder">
    <table class="pcCPcontent">
        <tr>
            <td colspan="2"><h2>You are editing Administrator Comments for <strong>Order #<%=clng(scpre)+intIdOrder%></strong></h2></td>
        </tr>
        <tr>
            <td align="right" valign="top">
            </td>
            <td>
              <textarea class="htmleditor" name="adminComments" id="adminComments" cols="70" rows="15" id="adminComments"><%=padminComments%></textarea>
            </td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td colspan="2" align="center">
                <input type="submit" name="Submit" value="Add or Edit Comments" class="btn btn-primary">
                <input type="button" class="btn btn-default"  name="Back" value="Back" onClick="document.location.href='ordDetails.asp?id=<%=intIdOrder%>'">
            </td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
    </table>
</form>
<!--#include file="Adminfooter.asp" -->
