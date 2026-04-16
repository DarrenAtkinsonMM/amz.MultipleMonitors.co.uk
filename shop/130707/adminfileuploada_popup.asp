<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
on error resume next
session("UIDFeedback")=request("IDFeedback")
session("Relink")=request("ReLink")

pcv_intMaxUploads = 6
%>
<html>
<head>
<title>Upload Data File(s)</title>
<script type=text/javascript>
	
function isCSV(s)
	{
		var test=""+s ;
		test2="";
		for (var k=test.length-4; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			test2 += c
		}
		test1=test2.toLowerCase()
		if (test1==".txt"||test1==".gif"||test1==".jpg"||test1==".htm"||test1==".zip"||test1==".pdf"||test1==".doc"||test1==".png")
			{
				return (true);
			}
		test2="";
		for (var k=test.length-5; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			test2 += c
		}
		test1=test2.toLowerCase()
		if (test1==".html"||test1==".jpeg")
			{
				return (true);
			}			
		
		return (false);
	}
	
function containsComma(s)
	{
	var pos=s.indexOf(",");
	if (pos>=0)
	{
		return(true);
	}
	return(false);
}
	

function Form1_Validator(theForm)
{
	if (<% For i = 1 To pcv_intMaxUploads %> theForm.file_<%= i %>.value == "" <% if i < pcv_intMaxUploads then %> && <% end if %> <% next %>)
	{
		alert("You need to supply at least one file to upload.");
		theForm.file_1.focus();
		return (false);
	}
	else
	{
		<% For i = 1 To pcv_intMaxUploads %>
		if (theForm.file_<%= i %>.value != "")
		{
			if (isCSV(theForm.file_<%= i %>.value) == false)
			{
				alert("File type not allowed. The file cannot be uploaded to the server.");
				theForm.file_<%= i %>.focus();
				return (false);
				}
			if (containsComma(theForm.file_<%= i %>.value)==true)
			{
				alert("The file name cannot contain a comma.");
				theForm.file_<%= i %>.focus();
				return (false);
			}
		}
		<% Next %>
	}
	return (true);
}
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form method="post" enctype="multipart/form-data" action="adminfileupl_popup.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
        <table width="100%" border="0" cellspacing="0" cellpadding="4" align="center">
          <tr> 
            <td colspan="3" bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Upload
              Data File(s)</font></b></font></td>
          </tr>
          <tr> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td colspan="2"><font face="Arial, Helvetica, sans-serif" size="2">Only *.txt, *.htm, *.html, *.gif, *.jpg, *.pdf, *.doc and *.zip file types may be uploaded.</font></td>
          </tr>
          <tr> 
            <td colspan="2" height="10"></td>
          </tr>
          
          <% For i = 1 To pcv_intMaxUploads %>
          <tr> 
            <td width="20%">
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">File <%= i %>: </font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2">
              <input class="ibtng" type="file" name="file_<%= i %>" size="25">
              </font></b></td>
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
	</form>
</body>
</html>
<% call closeDb() %>