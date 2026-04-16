<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
pageTitle = "Developer Console" 'dictLanguage.Item(Session("language")&"_pcAppBtnSettings")
pageIcon = "pcv4_icon_settings.png"
%>
<!--#include file="AdminHeader.asp"-->
<%
pcPageName="pcws_DevCon.asp"
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<div class="apps">
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-12">

                <ul class="nav nav-tabs" role="tablist">
                    <li role="presentation" class="active"><a href="#home" aria-controls="home" role="tab" data-toggle="tab">Sandbox Apps</a></li>
                    <li role="presentation"><a href="#widgets" aria-controls="widgets" role="tab" data-toggle="tab">Widgets</a></li>
                    <li role="presentation"><a href="#hooks" aria-controls="hooks" role="tab" data-toggle="tab">Event Hooks</a></li>
                </ul>
                <div class="tab-content">
                    <div role="tabpanel" class="tab-pane active" id="home">
                        
                        <br />
                        <div class="row" data-ng-controller="appsCtrl">
                            <div class="col-sm-12">
                            
                            
<%
query="SELECT pcPCWS_Uid, pcPCWS_AuthToken, pcPCWS_Username, pcPCWS_Password FROM pcWebServiceSettings;"
Set rs=connTemp.execute(query)
If Not rs.eof Then
    pcv_strUid = rs("pcPCWS_Uid")
    pcv_AuthToken = rs("pcPCWS_AuthToken")  
    pcv_strUsername = rs("pcPCWS_Username")  
    pcv_strPassword = enDeCrypt(rs("pcPCWS_Password"), scCrypPass)          
End If
Set rs=nothing
%>                            
                            
<div class="row">                           
<%
Dim SandboxPath

If PPD="1" Then
    SandboxPath = Server.MapPath("/" & scPcFolder & "/includes/apps/") & "/"
Else
    SandboxPath = Server.MapPath("../includes/apps/") & "/"
End If

Dim SandboxFS, SandboxDir, SandboxId, SandboxName, SandboxStatus
Set SandboxFS = Server.CreateObject("Scripting.FileSystemObject")
Set SandboxDir = SandboxFS.GetFolder(SandboxPath)


If SandboxDir.SubFolders.Count > 0 Then

    For Each Folder in SandboxDir.SubFolders
    
        'If (instr(Folder.name, "__") > 0) Then

            %>
            <div class="col-sm-4">
                <%=Folder.name %>
            </div>
            <div class="col-sm-8">
                    <a class="btn  btn-success" href="" data-ng-click="Install('/MyApps/<%=pcv_strUid %>', '<%=Folder.name %>');" ng-disabled="isDisabled"><span class="glyphicon glyphicon-cloud-download" aria-hidden="true"></span> Install</a>

                    <a class="btn btn-sm btn-default" href="pcws_Settings.asp?fc=<%=Folder.name %>"><span class="glyphicon glyphicon-cog" aria-hidden="true"></span> Settings</a>
            </div>
            <%

        'End If
        
    Next

End If

Set SandboxFS = Nothing
Set SandboxDir = Nothing
%>                        
</div>                            
                            
                            

                            </div>
                        </div>            
                        
                    </div>
                    <div role="tabpanel" class="tab-pane" id="widgets">

                        <div data-ng-controller="apiCtrl" ng-init="load('/<%=scAdminFolderName %>/service/api/widgets.asp')">
                            <form class="form" id="widgetsForm" name="widgetsForm">
                                <table class="table table-condensed">
                                    <tr>
                                        <td class="col-sm-3"><input name="widget_Desc" placeholder="Description" type="text" class="form-control" /></td>
                                        <td class="col-sm-2"><input name="widget_Shortcode" placeholder="Short Code" type="text" class="form-control" /></td>
                                        <td class="col-sm-4">
                                            <select name="widget_Type" placeholder="Type" type="text" class="form-control">
                                                <option value="Execute">Execute</option>
                                            <select>
                                        </td>
                                        <td colspan="2" class="col-sm-3"><input name="widget_Method" placeholder="Method" type="text" class="form-control" /></td>
                                    </tr>
                                    <tr>
                                        <td class="col-sm-11" colspan="4">
                                            <input name="widget_Uri" placeholder="Enter the interface Uri (e.g. '../includes/apps/pcHelloWordASP/interface.asp')" type="text" class="form-control" />
                                        </td>
                                        <td class="col-sm-1">
                                            <button data-ng-click="create('/<%=scAdminFolderName %>/service/api/widgets.asp', 'widgetsForm')" class="btn btn-default">Create</button>
                                            <input type="hidden" name="action" value="createWidget" />
                                        </td>
                                    </tr>
                                </table>
                                
                                <table class="table table-hover table-condensed">
                                    <tr>
                                        <th class="col-sm-3">Description</th>
                                        <th class="col-sm-2">Short Code</th>
                                        <th class="col-sm-4">Type</th>
                                        <th class="col-sm-2">Method</th>
                                        <th class="col-sm-1"></th>
                                    </tr>        
                                    <tr data-ng-repeat="widget in data.widgets">
                                        <td><samp>{{widget.Desc}}</samp></td>
                                        <td><kbd>{{widget.ShortCode}}</kbd></td>
                                        <td>
                                            <samp>{{widget.Type}}</samp><br />
                                            <small>{{widget.Uri}}</small>
                                        </td>
                                        <td><code>{{widget.Method}}</code></td>
                                        <td>
                                            <a class="btn btn-default btn-xs" data-ng-click="delete('/<%=scAdminFolderName %>/service/api/widgets.asp', 'wid', widget.ID)"><i class="fa fa-times" aria-hidden="true"></i></a>
                                        </td>
                                    </tr>
                                </table>
                                
                            </form>
                        </div>
                    </div>
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    <div role="tabpanel" class="tab-pane" id="hooks" autocomplete="off">
                        <div data-ng-controller="apiCtrl" ng-init="load('/<%=scAdminFolderName %>/service/api/widgets.asp')"> 
                            <form class="form" id="hooksForm" name="hooksForm">
                                <table class="table table-condensed">
                                    <tr>
                                        <td class="col-sm-3"><input name="hook_Desc" placeholder="Description" type="text" class="form-control" /></td>
                                        <td class="col-sm-2"><input name="hook_Shortcode" placeholder="Short Code" type="text" class="form-control" /></td>
                                        <td class="col-sm-3">
                                            <select name="hook_Event" placeholder="Event" class="form-control">
                                                <option value="" selected>Select Event</option>
                                                <option value="ProductModified">Product Modified</option>
                                                <option value="ProductRemoved">Product Removed</option>
                                                <option value="ProductPurged">Product Purged</option>
                                                <option value="StockChanged">Stock Changed</option>
                                                <option value="OrderProcessed">Order Processed</option>
                                                <option value="OrderCompleted">Order Completed</option>
                                                <option value="CustResetPassEmailSent">Customer Reset Password Email</option>
                                                <option value="SendAlarmEmailSent">Alarm Email Sent</option>
                                                <option value="NewCustEmailSent">New Customer Email</option>
                                                <option value="AffRetrievePassEmailSent">Affiliate Retrieve Password Email</option>
                                                <option value="ForgotOrderCodeEmailSent">Forgot Order Code Email</option>
                                                <option value="ContactUsEmailSent">Contact Email</option>
                                                <option value="NewOrderEmailSent">New Order Email</option>
                                                <option value="OrderConfirmationEmailSent">Order Confirmation Email</option>
                                                <option value="OrderReceivedEmailSent">Order Received Email</option>
                                                <option value="Order Shipped Email">OrderShippedEmailSent</option>
                                                <option value="OrderPartShippedEmailSent">Order PartShipped Email</option>
                                                <option value="GROrderEmailSent">GR Order Email</option>
                                                <option value="GCOrderEmailSent">GC Order Email</option>
                                                <option value="AffOrderEmailSent">Aff Order Email</option>
                                                <option value="HelpDeskEmailSent">Help Desk Email</option>
                                                <option value="ProductReviewEmailSent">Product Review Email</option>
                                            
                                            <select>
                                        </td>
                                        <td colspan="2" class="col-sm-4"><input name="hook_Method" placeholder="Method" type="text" class="form-control" /></td>
                                    </tr>
                                    <tr>
                                        <td class="col-sm-11" colspan="4">
                                            <input name="hook_Uri" placeholder="Enter the interface Uri (e.g. '../includes/apps/pcHelloWordASP/interface.asp')" type="text" class="form-control" />
                                        </td>
                                        <td class="col-sm-1">
                                            <button data-ng-click="create('/<%=scAdminFolderName %>/service/api/widgets.asp', 'hooksForm')" class="btn btn-default">Create</button>
                                            <input type="hidden" name="action" value="createHook" />
                                        </td>
                                    </tr>
                                </table>
    
                                <table class="table table-striped table-hover table-condensed">
                                    <tr>
                                        <th class="col-sm-3">Description</th>
                                        <th class="col-sm-2">Short Code</th>
                                        <th class="col-sm-3">Event</th>
                                        <th class="col-sm-3">Method</th>
                                        <th class="col-sm-1"></th>
                                    </tr>
                                    <tr data-ng-repeat="hook in data.hooks">
                                        <td><samp>{{hook.Desc}}</samp></td>
                                        <td><kbd>{{hook.ShortCode}}</kbd></td>
                                        <td><samp>{{hook.Event}}</samp></td>
                                        <td><code>{{hook.Method}}</code></td>
                                        <td>
                                            <a class="btn btn-default btn-xs" data-ng-click="delete('/<%=scAdminFolderName %>/service/api/widgets.asp', 'hid', hook.ID)"><i class="fa fa-times" aria-hidden="true"></i></a>
                                        </td>
                                    </tr>
                                </table>
                            
                            </form>
                            
                        </div>
                    </div>
                </div>

            </div>
        </div>
    </div>
</div>

<!--#include file="AdminFooter.asp"-->