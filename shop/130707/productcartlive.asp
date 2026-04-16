<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Check for Updates" %>
<% Section="updatees" %>
<%PmAdmin=1%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/utilities.asp"-->
<%
pcPageName="productcartlive.asp"
%>
<%
IsApparel = False
IsConfig = False
IsConfigPlus = False

pcv_baseURL = "http://service.productcartlive.com/v1/api/Clients"

'// START: JSON
dim jsonService : set jsonService = JSON.parse("{}")

    jsonService.set "packageId", pcl_PackageId 
    jsonService.set "EmailPartner", pcl_EmailPartner 

Dim jsonObj
jsonObj = JSON.stringify(jsonService, null, 2)
'// END: JSON


'// START: POST
Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
objXMLhttp.open "POST", pcv_baseURL, false
objXMLhttp.setRequestHeader "Content-type","application/json"
objXMLhttp.setRequestHeader "Accept","application/json"
objXMLhttp.send jsonObj
cfuResult = objXMLhttp.responseText
Set objXMLhttp = Nothing
'// END: POST


'// START: PARSE RESULTS
dim Info : set Info = JSON.parse(cfuResult)
          
for each key in Info.configoptions.configoption.keys()
    
    pcv_strOptionName = Info.configoptions.configoption.get(key).option 
    pcv_strOptionValue = Info.configoptions.configoption.get(key).value  
    
    If pcv_strOptionName="Apparel Add-On" Then
        
        '// Update Apparel Add-On
        If pcv_strOptionValue = 1 Then
            IsApparel = True
        Else
            IsApparel = False     
        End If     

    End If
    
    If pcv_strOptionName="QuickBooks Add-On" Then
        
        '// Update QuickBooks Add-On
        If pcv_strOptionValue = 1 Then
            IsQBWC = True
        Else
            IsQBWC = False      
        End If      

    End If
    
    If pcv_strOptionName="Configurator Add-On" Then
        
        '// Update Configurator Add-On
        If pcv_strOptionValue = "Base" Then
            IsConfig = True
            IsConfigPlus = False
        ElseIf pcv_strOptionValue = "Base w/ Conflict Management" Then
            IsConfig = True 
            IsConfigPlus = True
        Else
            IsConfig = False 
            IsConfigPlus = False    
        End If      

    End If

next
call pcf_upgradeDowngrade(IsApparel, IsConfig, IsConfigPlus, IsQBWC)
%>
<!--#include file="pcAdminRetrieveSettings.asp"-->
<!--#include file="pcAdminSaveSettings.asp"-->
<%
'// END: PARSE RESULTS


Function pcf_upgradeDowngrade(IsApparel, IsConfig, IsConfigPlus, IsQBWC)

    Dim strtext1
    
	if PPD="1" then
		pcStrFolder = "/"&scPcFolder&"/includes"
	else
		pcStrFolder = "../includes"
	end if

    strtext1 = ""
    If IsApparel Then
        strtext1 = "<" & Chr(37) & " Dim statusAPP" & vbNewLine & "statusAPP=""1""" & vbNewLine & Chr(37) & ">"
    Else
        strtext1 = "<" & Chr(37) & " Dim statusAPP" & vbNewLine & "statusAPP=""0""" & vbNewLine & Chr(37) & ">"
    End IF
    call pcs_SaveUTF8(pcStrFolder & "\statusAPP.inc", pcStrFolder & "\statusAPP.inc", strtext1)

    strtext1 = ""
    If IsConfig Then
        strtext1 = "<" & Chr(37) & " Dim statusBTO" & vbNewLine & "statusBTO=""1""" & vbNewLine & Chr(37) & ">"
    Else
        strtext1 = "<" & Chr(37) & " Dim statusBTO" & vbNewLine & "statusBTO=""0""" & vbNewLine & Chr(37) & ">"    
    End IF
    call pcs_SaveUTF8(pcStrFolder & "\status.inc", pcStrFolder & "\status.inc", strtext1)

    strtext1 = ""
    If IsConfigPlus Then
        strtext1 = "<" & Chr(37) & " Dim statusCM" & vbNewLine & "statusCM=""1""" & vbNewLine & Chr(37) & ">"      
    Else
        strtext1 = "<" & Chr(37) & " Dim statusCM" & vbNewLine & "statusCM=""0""" & vbNewLine & Chr(37) & ">"
    End IF
    call pcs_SaveUTF8(pcStrFolder & "\statusCM.inc", pcStrFolder & "\statusCM.inc", strtext1)

End Function
%>