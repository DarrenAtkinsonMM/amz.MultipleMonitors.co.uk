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
pageTitle=dictLanguage.Item(Session("language")&"_pcAppBtnMyAccount")
pageIcon="pcv4_icon_settings.png"
%>
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<% 
pcPageName="pcws_MyAccount.asp"

'/////////////////////////////////////////////////////
'// Retrieve current database data
'/////////////////////////////////////////////////////
%>
<!--#include file="pcAdminRetrieveSettings.asp"-->
<%
'// START: Refresh Token
If request("refresh")<>"" Then
    pcf_UpdateToken()
End If
'// END: Refresh Token

'// START: Activate ProductCart WebServices
If request("activatePCWS")<>"" Then


    '// 1) Get the Info
    pcv_strEmail = getUserInput(request("Email"), 0)
    pcv_strUsername = pcv_strEmail '// scCrypPass      
    pcv_strFirstName = getUserInput(request("FirstName"), 0)
    pcv_strLastName = getUserInput(request("LastName"), 0)  
    pcv_strAccountNum = getUserInput(request("AccountNum"), 0)
    pcv_strPassword = getUserInput(request("AuthPassword"), 0)
    pcv_strConfirmPassword = pcv_strPassword    


    '// 2) Gen Request
    dim jsonService : set jsonService = JSON.parse("{}")    
    Dim jsonObj
    If len(pcv_strAccountNum)=0 Then

        jsonService.set "Email", pcv_strEmail 
        jsonService.set "Username", pcv_strUsername 
        jsonService.set "Password", pcv_strPassword 
        jsonService.set "ConfirmPassword", pcv_strConfirmPassword 
        jsonService.set "FirstName", pcv_strFirstName 
        jsonService.set "LastName", pcv_strLastName     
        jsonService.set "LicenseKey", scCrypPass 
        
        jsonService.set "storeVersion", scVersion 
        jsonService.set "subVersion", scSubVersion 
        jsonService.set "servicePack", scSP 
        jsonService.set "qbVersion", qbsv 
        jsonService.set "referrer", currentURL
        jsonService.set "themeFolder", scThemePath

        jsonObj = JSON.stringify(jsonService, null, 2)
        
        pcv_baseURL = pcv_baseURL & "/accounts/create/"
    
    Else
     
        pcv_baseURL = pcv_baseURL & "/accounts/user/" & pcv_strAccountNum
    
    End If


    '// 3) Send off the Registration
    If len(pcv_strAccountNum)=0 Then
        cfuResult = pcf_PostRequest(jsonObj, pcv_baseURL, "")
    Else
        pcv_strAuthToken = pcf_GetToken(pcv_strUsername, pcv_strPassword)     
        cfuResult = pcf_GetRequest(pcv_baseURL, pcv_strAuthToken)
    End If


    '// 4) Parse Response    
    pcv_strMessage = ""
    If len(cfuResult)=0 Then
        pcv_strMessage = "Sorry, there was a problem. Please contact support. "
    End If
    
    dim Info : set Info = JSON.parse(cfuResult)
    
    pcv_strUrl = ""
    If instr(cfuResult, "url")>0 Then
        pcv_strUrl = Info.url
    End If

    If instr(cfuResult, "message")>0 Then
        pcv_strMessage = Info.message
    End If
    
    If instr(cfuResult, "Message")>0 Then
        pcv_strMessage = Info.Message
    End If


    '// 5) If Error Redirect with Message
    If pcv_strMessage="The request is invalid." Then 
        dim Info2 : set Info2 = Info.modelState
        for each key in Info2.keys()
            pcv_strReason = Info2.get(key)
        next
        msg = "The request is invalid. "
        If len(pcv_strReason)>0 Then
            msg = pcv_strReason
        End If
        response.Redirect(pcPageName & "?msg=" & msg & "&s=2")
    End If
    
    If pcv_strMessage="An error has occurred." Then   
        msg = "Oops!  There was an error processing the activation. Please contact support."
        response.Redirect(pcPageName & "?msg=" & msg & "&s=2")
    End If
    
    If len(pcv_strMessage)>0 Then   
        msg = "Oops! " & pcv_strMessage & " Please contact support."
        response.Redirect(pcPageName & "?msg=" & msg & "&s=2")
    End If


    '// 6) If Active Update the Db
    If pcv_strMessage="" And len(pcv_strUrl)>0 Then 

        pcv_strId = ""
        If instr(cfuResult, "id")>0 Then
            pcv_strId = Info.id
        End If
        
        pcv_strUsername = ""
        If instr(cfuResult, "userName")>0 Then
            pcv_strUsername = Info.userName
        End If
        
        pcv_strFullname = ""
        If instr(cfuResult, "fullName")>0 Then
            pcv_strFullname = Info.fullName
        End If
        
        pcv_strEmail = ""
        If instr(cfuResult, "email")>0 Then
            pcv_strEmail = Info.email
        End If
        
        pcv_strEmailConfirmed = ""
        If instr(cfuResult, "emailConfirmed")>0 Then
            pcv_strEmailConfirmed = Info.emailConfirmed
            If pcv_strEmailConfirmed="True" Then
                pcv_strEmailConfirmed = 1
            Else
                pcv_strEmailConfirmed = 0
            End If
        End If
        
        pcv_strLevel = ""
        If instr(cfuResult, "level")>0 Then
            pcv_strLevel = Info.level
        End If
        
        pcv_strJoinDate = ""
        If instr(cfuResult, "joinDate")>0 Then
            pcv_strJoinDate = Info.joinDate
        End If
        
        pcv_strLicenseKey = ""
        If instr(cfuResult, "licenseKey")>0 Then
            pcv_strLicenseKey = Info.licenseKey
        End If
        
        pcv_strAuthToken = pcf_GetToken(pcv_strUsername, pcv_strPassword) 

        query="UPDATE pcWebServiceSettings SET "
        query=query&"pcPCWS_TurnOnOff=1, "
        query=query&"pcPCWS_IsActive=1, "
        query=query&"pcPCWS_Url='" & pcv_strUrl & "', "
        query=query&"pcPCWS_Uid='" & pcv_strId & "', "
        query=query&"pcPCWS_Username='" & pcv_strUsername & "', "
        query=query&"pcPCWS_Password='" & enDeCrypt(pcv_strPassword, scCrypPass) & "', "
        query=query&"pcPCWS_Fullname='" & pcv_strFullname & "', "
        query=query&"pcPCWS_Email='" & pcv_strEmail & "', "
        query=query&"pcPCWS_EmailConfirmed=" & pcv_strEmailConfirmed & ", "
        query=query&"pcPCWS_Level=" & pcv_strLevel & ", "
        query=query&"pcPCWS_AuthToken='" & pcv_strAuthToken & "', "
        query=query&"pcPCWS_JoinDate='" & pcv_strJoinDate & "', "
        query=query&"pcPCWS_LicenseKey='" & pcv_strLicenseKey & "' "
        Set rs2 = server.CreateObject("ADODB.RecordSet")
        Set rs2 = connTemp.execute(query)
        Set rs2 = Nothing  

    End If


    '// Write Constants
    pcs_GenGlobalWebServiceSettings()

	msg = "Thank you for registering! " '& pcv_strPassword
    response.Redirect(pcPageName & "?msg=" & msg & "&s=1")
    
End If
'// END: Activate ProductCart WebServices
%>


<%
'// START: Update Settings
If request("updateSettings")<>"" Then
	
    pcv_intStoreOn = getUserInput(request("StoreOn"), 1)

    '// Turn On / Off
    query="UPDATE pcWebServiceSettings SET pcPCWS_TurnOnOff=" & pcv_intStoreOn & ";"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing  
    
    '// Write Constants
    pcs_GenGlobalWebServiceSettings()

	msg = "Updated PCWS Settings successfully!"
    response.Redirect(pcPageName & "?msg=" & msg & "&s=1")
    
End If
'// END: Update Settings
%>

<%
'// START: Page Load
query="SELECT pcPCWS_TurnOnOff, pcPCWS_IsActive FROM pcWebServiceSettings"
Set rs = server.CreateObject("ADODB.RecordSet")
Set rs=connTemp.execute(query)
If rs.eof Then

    '// Generate Table for first time use
    query="INSERT INTO pcWebServiceSettings (pcPCWS_TurnOnOff, pcPCWS_IsActive) VALUES (0, 0);"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing  
          
End If
Set rs=nothing
'// END: Page Load
%>
<%
Session("pcSupportPlanEligible") = "1"
'response.Write("Support Plan Eligible: " & Session("pcSupportPlanEligible") & "<br />")
'Session("pcSupportPlanEligible") = "1"
'response.Write("scPCWS_IsActive:  " & scPCWS_IsActive)
%>
<form name="form1" method="post" action="<%=pcPageName%>" class="pcForms">

    <% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
    
    <% If scPCWS_IsActive = "0" Then %>
    
        <% If Session("pcSupportPlanEligible") = "0" Then %>
            <p class="bs-callout bs-callout-warning">
                <h4><%=dictLanguage.Item(Session("language")&"_pcAppSupportH4") %></h4>
                <%=dictLanguage.Item(Session("language")&"_pcAppSupport") %>
            </p>  
        <% End If %>
    
        <div class="panel panel-default">                        
            
            <div class="panel-heading">
                <h3 class="panel-title"><%=dictLanguage.Item(Session("language")&"_pcAppTagLine") %></h3>
            </div>
            <div class="panel-body">
                
                <br />
             
                <div class="form-group">
                    <input type="text" class="form-control" name="Firstname" id="Firstname" placeholder="Firstname" value="<%=Server.HTMLEncode(pcv_strFirstname)%>" />
                </div>
                
                <div class="form-group">
                    <input type="text" class="form-control" name="Lastname" id="Lastname" placeholder="Lastname" value="<%=Server.HTMLEncode(pcv_strLastname)%>" />
                </div>
                
                <div class="form-group">
                    <input type="email" class="form-control" name="Email" id="Email" placeholder="Email" value="<%=Server.HTMLEncode(pcv_strEmail)%>" />
                </div>
                
                <div class="form-group">
                    <input type="password" class="form-control" name="AuthPassword" placeholder="Password">
                </div>
                
                <div class="collapse" id="collapseExample">
    
                    <div class="form-group">
                        <input type="text" class="form-control" name="AccountNum" placeholder="Account Number">
                    </div>
    
                </div>
                
                <input type="submit" name="activatePCWS" value="<%=dictLanguage.Item(Session("language")&"_pcAppBtnCreate") %>" <% If Session("pcSupportPlanEligible") = "1" Then %>class="btn btn-primary"<% Else %>class="btn btn-primary disabled" disabled<% End If %> />
                
                <button class="btn btn-default" type="button" data-toggle="collapse" data-target="#collapseExample" aria-expanded="false" aria-controls="collapseExample">
                  <%=dictLanguage.Item(Session("language")&"_pcAppLinkAlready") %>
                </button>
    
                <input type="hidden" name="activatePCWS" value="1" />
    
                <br />
                <br />
                
            </div>  
            <div class="panel-footer">
                <a href="http://www.productcart.com/policies.asp#security" title="Privacy Policy" target="_blank"><%=dictLanguage.Item(Session("language")&"_pcAppLinkPrivacy") %></a>
                 &nbsp;|&nbsp;
                 <a href="http://wiki.productcart.com/productcart/eula" title="End User License Agreement" target="_blank"><%=dictLanguage.Item(Session("language")&"_pcAppLinkTerms") %></a>
            </div> 
    
        </div> 
    
    <% Else %>
    
        <% If Session("pcSupportPlanEligible") = "0" Then %>
            <p class="bs-callout bs-callout-info">
                <%=dictLanguage.Item(Session("language")&"_pcAppSupportExpired") %>
            </p>
        <% End If %>
        
        <!--#include file="pcws_Navigation.asp"-->
        
        <br />
        
        <div class="container-fluid">
            <div class="row">
                <div class="col-md-12">
                    
                    <div class="well">
                        <p>
                            <strong><%=dictLanguage.Item(Session("language")&"_pcAppActStatus") %></strong>  
                            <% if scPCWS_StoreOn="1" then%>
                                <span class="label label-success">Active</span> <%'=scPCWS_StoreOn %>
                            <% else %>
                                <span class="label label-warning">Suspended</span> <%'=scPCWS_StoreOn %>
                            <% end if %>
                        </p>
    
    
                        <div style="display: none">
                        
                            <div class="well well-sm"><strong>NOTE:</strong> Turning off ProductCart Apps is not the same as cancelling. In order to stop charges you need to cancel all of your services. To cancel billable services visit the "My Account" page.</div>
                        
                            <input type="radio" name="StoreOn" value="1" checked class="clearBorder">Turn On ProductCart Apps
                            <input type="radio" name="StoreOn" value="0" <% if scPCWS_StoreOn="0" then%>checked<% end if %> class="clearBorder">Turn Off ProductCart Apps
                        </div>
                        
                        <hr />
                        
                        <p>
                            <strong><%=dictLanguage.Item(Session("language")&"_pcAppActNum") %></strong>  <%=scPCWS_Uid %>
                        </p>
                        
                        <a href="pcws_MyAccount.asp?refresh=1" title="<%=dictLanguage.Item(Session("language")&"_pcAppBtnRefresh") %>" class="btn btn-default">
                            <span class="glyphicon glyphicon-refresh" aria-hidden="true"></span>
                            <%=dictLanguage.Item(Session("language")&"_pcAppBtnRefresh") %>
                        </a>
                        
                    </div>                            
                    
                </div>
            </div>
        </div>
    
    <% End If %>

</form>
<!--#include file="AdminFooter.asp"-->
