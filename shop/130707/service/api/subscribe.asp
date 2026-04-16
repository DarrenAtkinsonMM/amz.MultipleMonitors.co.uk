<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="ProductCart Market Subscription Service"
pageIcon="pcv4_icon_settings.png"
%>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="../../adminv.asp"--> 
<!--#include file="../../../includes/common.asp"-->
<!--#include file="../../../includes/common_checkout.asp"-->
<!--#include file="../../../includes/languagesCP.asp"-->
<% 
pcPageName="subscribe.asp"

'/////////////////////////////////////////////////////
'// Retrieve current database data
'/////////////////////////////////////////////////////
%>
<!--#include file="../../pcAdminRetrieveSettings.asp"-->
<%
pcv_baseURL = pcv_baseURL & "/accounts/user/"
pcv_strThisFeatureCode = request("code") '// "PCWSCDN"
pcv_strThisFeatureName = request("name") '// "ProductCart CDN"


'// START: Activate ProductCart WebServices
If request("event")="subscribe" Then


    '// 1) Get the Info (OK)
    query="SELECT pcPCWS_Uid, pcPCWS_AuthToken, pcPCWS_Username, pcPCWS_Password FROM pcWebServiceSettings;"
    Set rs=connTemp.execute(query)
    If Not rs.eof Then
        pcv_strUid = rs("pcPCWS_Uid")
        pcv_AuthToken = rs("pcPCWS_AuthToken")  
        pcv_strUsername = rs("pcPCWS_Username")  
        pcv_strPassword = enDeCrypt(rs("pcPCWS_Password"), scCrypPass)          
    End If
    Set rs=nothing



    '// 2) Gen Request (OK)
    Dim jsonService : Set jsonService = JSON.parse("{}")
    jsonService.Set "Type", "Feature"
    jsonService.Set "Value", pcv_strThisFeatureCode
    
    Dim jsonObj
    jsonObj = JSON.stringify(jsonService, null, 2)    
    jsonObj = "[" & jsonObj & "]"
    'response.Write(jsonObj)
    'response.End()


    '// 3) Send off the Service Activation (OK)
    cfuResult = pcf_PutRequest(jsonObj, pcv_baseURL & pcv_strUid & "/assignclaims", pcv_AuthToken)
    'response.Write(cfuResult)
    'response.End()


    '// 4) Parse Response (OK)
    pcv_strMessage = ""    
    If len(cfuResult)>0 Then    
        Dim Info : Set Info = JSON.parse(cfuResult)        
        If instr(cfuResult, "message")>0 Then
            pcv_strMessage = Info.message
        End If    
    End If



    '// 5) If Error Redirect with Message
    If pcv_strMessage="The request is invalid." Then 
        Dim Info2 : Set Info2 = Info.modelState
        For Each key In Info2.keys()
            pcv_strReason = Info2.get(key)
        Next
        msg = "The request is invalid. " & pcv_strReason
        'response.Redirect(pcPageName & "?msg=" & msg & "&s=2")
        call closeDb()
        response.Write(msg)
        response.End()
    End If
    
    If pcv_strMessage="An error has occurred." Then   
        msg = "Oops!  There was an error adding the feature. Please contact support."
        'response.Redirect(pcPageName & "?msg=" & msg & "&s=2")
        call closeDb()
        response.Write(msg)
        response.End()
    End If
    
    If len(pcv_strMessage)>0 Then   
        msg = "Oops!  There was an error adding the feature. Please contact support."
        'response.Redirect(pcPageName & "?msg=" & msg & "&s=2")
        call closeDb()
        response.Write(msg)
        response.End()
    End If



    '// 6) If Active...
    If pcv_strMessage="" Then


        '// 7) Verify the Claim by Requesting User Info
        pcv_boolIsClaimValid = pcf_VerifyClaim(pcv_baseURL, pcv_strUid, pcv_AuthToken)
        'response.Write(pcv_boolIsClaimValid)
        'response.End()


        '// 8) Request a new Token
        If pcv_boolIsClaimValid Then 
            pcv_strAuthToken = pcf_GetToken(pcv_strUsername, pcv_strPassword) 
            'response.Write(pcv_strAuthToken)
            'response.End()   
        Else
            msg = "Oops!  Looks like there was a minor issue. Please open a ticket."
            'response.Redirect(pcPageName & "?msg=" & msg & "&s=2")
            call closeDb()
            response.Write(msg)
            response.End()
        End If


        '// 9) Save Token
        call pcs_SaveToken(pcv_strAuthToken)

    End If

	msg = "success"
    call closeDb()
    response.Write(msg)
    response.End()
    
End If
'// END: Activate ProductCart WebServices
%>

<%
'// START: Unsubscribe
If request("unsubscribe")<>"" Then

    '// Unsubscribe
    pcv_boolIsUnsubscribed = True
    
    If pcv_boolIsUnsubscribed Then
    
        '// Turn Feature Off
        call pcs_UpdateFeatureStatusByCode(pcv_strThisFeatureCode, 0)
        
        '// Disable Service
        call pcs_UpdateFeature(pcv_strThisFeatureCode, 0)
        
        call pcs_GenGlobalWebServiceSettings()
    
        msg = "Unsubscribed successfully!"
        response.Redirect("pcws_Market.asp" & "?msg=" & msg & "&s=1")
    
    Else
    
        msg = "We couldn't unsubscribe you at the moment. Please call or try again later."
        response.Redirect(pcPageName & "?msg=" & msg & "&s=1")
    
    End If
  
End If
'// END: Unsubscribe
%>

<%
call closeDb()
%>