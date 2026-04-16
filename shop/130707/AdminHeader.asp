<!--#include file="smallRecentProducts-header.asp"-->
<!DOCTYPE html>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

' If your store is using a dedicated SSL certificate (e.g. https://www.yourstore.com)
' you can use the following code to force all Control Panel pages to load securely
' using the HTTPS protocol. Remove the apostrophe from the beginning of each of the following
' 8 lines of code to use this feature. This code will not work with shared SSL certificates.

If scSSL="1" And scIntSSLPage="1" Then
  If (Request.ServerVariables("HTTPS") = "off") Then
      Dim xredir__, xqstr__
      xredir__ = "https://" & Request.ServerVariables("SERVER_NAME") & _
      Request.ServerVariables("SCRIPT_NAME")
      xqstr__ = Request.ServerVariables("QUERY_STRING")
      If xqstr__ <> "" Then xredir__ = xredir__ & "?" & xqstr__
      Response.redirect xredir__
  End If
End If
    
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
if err.number <> 0 then
	call closeDb()
    response.redirect "dbError.asp"
	response.End()
end if
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
Dim pcv_strAdminPrefix
pcv_strAdminPrefix="1"

If InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "msgb.asp") = 0 Then
	Session("message") = ""
End If

'// Load admin resources
pcv_strRequiredIcon = rsIconObj("requiredicon")
pcv_strErrorIcon = rsIconObj("errorfieldicon")

%>
<html>
<head>
<title>ProductCart v5 - Control Panel</title>
<meta charset="UTF-8" />
<meta name="description" content="ProductCart asp shopping cart software is published by NetSource Commerce. ProductCart's Control Panel allows you to manage every aspect of your ecommerce store. For more information and for technical support, please visit NetSource Commerce at http://www.productcart.com">
<META NAME="ROBOTS" CONTENT="NOINDEX, NOFOLLOW">

<!--#include file="inc_header.asp" -->

</head>
<body ng-app="productcart">
<% '// START: DO NOT REMOVE THIS LINE %>
<div id="pcMainService" data-ng-controller="serviceCtrl"></div> 
<% '// END: DO NOT REMOVE THIS LINE %>
<script type="text/javascript" src="../includes/pcjscolorchooser.js"></script>
<div id="pcCPmain">
	<div id="pcCPheader">
    	<div id="pcCPstoreName">
		<% '// Prepare and show company name
		Dim pcvStrCompanyName
		pcvStrCompanyName=scCompanyName
		if Len(pcvStrCompanyName)>34 then
		 pcvStrCompanyName=Left(pcvStrCompanyName,31) & "..."
		end if
		if pcvStrCompanyName="" or IsNull(pcvStrCompanyName) then
			pcvStrCompanyName="ProductCart v5"
		end if
		response.write pcvStrCompanyName
		%>
        </div>
        
        <div id="pcCPheaderNav">
            
            <a href="../pc/default.asp" target="_blank"><img src="images/cp11/cp11-storefront.png" width="14" height="14" alt="Storefront"> Storefront</a> [<% if scStoreOff="0" then %><span style="color: #090;">OPEN</span><% else %><span style="color: #F30;">CLOSED</span><% end if %>]
            <a href="http://wiki.productcart.com" target="_blank"><img src="images/cp11/cp11-docs.png" width="14" height="14" alt="ProductCart Wiki"> Wiki</a>
            <a href="http://blog.productcart.com" target="_blank"><img src="images/cp11/cp11-blog.png" width="14" height="14" alt="NetSource Commerce Blog"> Blog</a>
            <a href="http://twitter.com/productcart" target="_blank"><img src="images/cp11/cp11-twitter.png" width="14" height="14" alt="ProductCart on Twitter"> Twitter</a>
            <a href="https://www.facebook.com/productcartsoftware" target="_blank"><img src="images/cp11/cp11-facebook.png" width="14" height="14" alt="ProductCart on Facebook"> Facebook</a>
            <a href="about.asp"><img src="images/cp11/cp11-company.png" width="14" height="14" alt="About ProductCart"> About</a>
            
		</div>
        
        <div id="pcCPversion">
        
            <% If statusPCL="1" Then %>
            <a href="account.asp" class="btn btn-default btn-xs">Manage Hosting</a>&nbsp;&nbsp;
            <% End If %>
            
            ProductCart <strong>v<%=scVersion&scSubVersion%><% if scSP<>"" and scSP<>"0" then Response.Write(" SP " & scSP) end if %><% if PPD="1" then Response.Write(" PPD") end if %></strong>
            
    	</div>
        
        <div id="pcCPtopNav">
						<div id="smoothmenu1" class="ddsmoothmenu">
							<!--#include file="pcv4_navigation_links.asp"-->
						</div>
        </div>
		 
    </div>
    
    <div id="pcCPmainArea">
    
        <% If pcv_strDisplayType <> "1" Then %>
            <div id="pcCPmainLeft">
        <% Else %>
            <div id="pcCPmainCenter">          
        <% End If %>
				<%
					Set FS = Server.CreateObject("Scripting.FileSystemObject")
					pcv_tmpImagePath = Server.MapPath("images/") & "/"
					headerCSS = ""
					If FS.FileExists(pcv_tmpImagePath & pageIcon) Then
						headerCSS = "background-image: url(images/" & pageIcon & "); background-position: 700px 0px; background-repeat:no-repeat;"
					End If
					Set FS = Nothing
				%>
                <% If Not pcv_strDisplayType = "1" Then %>
				    <h1 style="<%= headerCSS %>"><%=pageTitle%></h1>
                <% End If %>

				<% If scUpgrade = 1 Then %>
					<div id="upgradeNotes">
						This store is currently in upgrade mode. Use caution as ALL settings are shared with the live store.  <a data-href="upddb_v50_complete.asp?status=1" data-toggle="modal" data-target="#confirm-delete" href="#" class="btn btn-default btn-xs">Switch to Live Mode</a>
					</div>
                    <div class="modal fade" id="confirm-delete" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
                        <div class="modal-dialog">
                            <div class="modal-content">
                                <div class="modal-header">
                                    <strong>Switch to Live Mode</strong>
                                </div>
                                <div class="modal-body">
                                    Click the "Continue" button to exit upgrade mode and switch to live mode. 
                                    This is the final step in the <a href="https://www.productcart.com/support/v5/article.asp?id=1" target="_blank">Upgrade Guide</a>. 
                                    Do not complete this step until you have completed all previous steps.
                                </div>
                                <div class="modal-footer">
                                    <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
                                    <a href="#" class="btn btn-danger danger">Continue</a>
                                </div>
                            </div>
                        </div>
                    </div>
                    <script>
                        $('#confirm-delete').on('show.bs.modal', function(e) {
                            $(this).find('.danger').attr('href', $(e.relatedTarget).data('href'));
                        });
                    </script>
                    <style>
                        #upgradeNotes {
                            position:fixed;                            
                            top:0px;                            
                            left:0px;                            
                            right:0px;                            
                            min-height:50px;                            
                            background-color:#D9534F;                                                       
                            color:#DEDEDE;                            
                            padding:10px;                            
                            font-size:20px;                            
                            box-shadow:0px 0px 10px #000;
                            text-align:center;
                            z-index: 999999;
                        }
                        body {
                            margin-top: 60px;
                            background-position: 0px 60px;  
                        }
                        .modal {
                            z-index: 999999;
                        }
                    </style>
				<% End If %>
                
                           
