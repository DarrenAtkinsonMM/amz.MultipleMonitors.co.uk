<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->
<% 
on error resume next

session("UIDFeedback")=request("IDFeedback")
session("Relink")=request("ReLink")

If session("UIDFeedback")&""="" OR session("Relink")&""="" OR scShowHD = 0 Then
	response.redirect "custPref.asp"
	response.end
End If

pcv_intMaxUploads = 6

%>
<!DOCTYPE html>
<html>
<head>
<title>Upload Data File(s)</title>
<!--#include file="inc_headerv5.asp" -->
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
			if (isCSV(theForm.file_<%= i %>.value)==false)
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
<body id="pcPopup">
<div id="pcMain">
<form method="post" enctype="multipart/form-data" action="userfileupl_popup.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
	<div class="pcMainContent">
		<h2>Upload Data File(s)</h2>
    
    <p>
    	Select a file using the &quot;Browse&quot; button, then click on &quot;Upload&quot;. Only *.txt, *.htm, *.html, *.gif, *.jpg, *.png, *.pdf, *.doc and *.zip file types were accepted.
    </p>
    
    <div class="pcSpacer"></div>
    
    <div class="pcShowContent">
    	<div class="pcFormItem">
      	<% For i = 1 To pcv_intMaxUploads %>
      	<div class="pcFormLabel pcFormLabelRight">
			 		File <%= i %>:
        </div>
        <div class="pcFormField">
        
			 	<input type="file" name="file_<%= i %>" size="25">
        </div>
        <% Next %>
      </div>
      <div class="pcSpacer"></div>
      
    	<div class="pcFormItem">
      	<div class="pcFormLabel"></div>
        <div class="pcFormField">
          <div class="pcFormButtons">
          	<button class="pcButtonUpload" name="Submit">Upload</button>
            <button class="pcButtonCloseWindow" onClick="javascript:window.close(); return false;">Close Window</button>
          </div>
        </div>
      </div>
    </div>
	</form>
</div>
</body>
</html>
<% call closeDb() %>
