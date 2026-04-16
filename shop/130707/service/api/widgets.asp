<%@  language="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "service.asp"
' This page outputs a JSON representation of the shopping cart.
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../../../includes/common.asp"-->
<!--#include file="../../../includes/common_checkout.asp"-->
<% 
response.Clear()
Response.ContentType = "text/json"
Response.Charset = "UTF-8"

pcv_widgetID= request("wid")
If len(pcv_widgetID)>0 Then
    query = "DELETE FROM pcWidgets WHERE widget_ID = " & pcv_widgetID
    set rs2 = server.CreateObject("ADODB.RecordSet")
    set rs2 = connTemp.execute(query)
    set rs2 = Nothing
End If

pcv_hookID= request("hid")
If len(pcv_hookID)>0 Then
    query = "DELETE FROM pcHooks WHERE hook_ID = " & pcv_hookID
    set rs2 = server.CreateObject("ADODB.RecordSet")
    set rs2 = connTemp.execute(query)
    set rs2 = Nothing
End If

pcv_action = request("action")
If pcv_action = "createHook" Then
    pcv_strDesc = request("hook_Desc")
    pcv_strShortcode = request("hook_Shortcode")
    pcv_strType = "Execute"
    pcv_strUri = request("hook_Uri")
    pcv_strMethod = request("hook_Method")
    pcv_strLang = "ASP" '// request("hook_Lang")
    pcv_strEvent = request("hook_Event")  
      
    query = "INSERT INTO pcHooks (hook_Desc, hook_Shortcode, hook_Type, hook_Uri, hook_Method, hook_Lang, hook_Event) VALUES ('" & pcv_strDesc & "', '" & pcv_strShortcode & "', '" & pcv_strType & "', '" & pcv_strUri & "', '" & pcv_strMethod & "', '" & pcv_strLang & "', '" & pcv_strEvent & "')"
    set rs2 = server.CreateObject("ADODB.RecordSet")
    set rs2 = connTemp.execute(query)
    set rs2 = Nothing
End If

pcv_action = request("action")
If pcv_action = "createWidget" Then
    pcv_strDesc = request("widget_Desc")
    pcv_strShortcode = request("widget_Shortcode")
    pcv_strType = "Execute"
    pcv_strUri = request("widget_Uri")
    pcv_strMethod = request("widget_Method")
    pcv_strLang = "ASP" '// request("widget_Lang")  
     
    query = "INSERT INTO pcWidgets (widget_Desc, widget_Shortcode, widget_Type, widget_Uri, widget_Method, widget_Lang) VALUES ('" & pcv_strDesc & "', '" & pcv_strShortcode & "', '" & pcv_strType & "', '" & pcv_strUri & "', '" & pcv_strMethod & "', '" & pcv_strLang & "')"

    set rs2 = server.CreateObject("ADODB.RecordSet")
    set rs2 = connTemp.execute(query)
    set rs2 = Nothing
End If


dim jsonService : set jsonService = JSON.parse("{}")

query = "SELECT widget_ID, widget_Shortcode, widget_Desc, widget_Type, widget_Uri, widget_Method, widget_Lang FROM pcWidgets WHERE widget_Type <> 'Core' "
set rs = Server.CreateObject("ADODB.Recordset")  
rs.Open query, connTemp, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.Eof Then
    pcv_intTotalCount = rs.RecordCount
    Redim widgets(pcv_intTotalCount - 1)   
    pcv_intCounter = 0
    Do While Not rs.Eof    
        pcv_intID = rs("widget_ID")     
        pcv_strShortcode = rs("widget_Shortcode")
        pcv_strDesc = rs("widget_Desc")
        pcv_strType = rs("widget_Type")
        pcv_strUri = rs("widget_Uri")
        pcv_strMethod = rs("widget_Method")
        pcv_strLang = rs("widget_Lang")
        
        Dim widget : Set widget = JSON.parse("{}") 
        widget.set "ID", pcv_intID       
        widget.set "ShortCode", pcv_strShortcode
        widget.set "Desc", pcv_strDesc
        widget.set "Type", pcv_strType
        widget.set "Uri", pcv_strUri
        widget.set "Method", pcv_strMethod
        widget.set "Lang", pcv_strLang
        Set widgets(pcv_intCounter) = widget
        Set widget = Nothing
        
        pcv_intCounter = pcv_intCounter + 1
        rs.movenext
    Loop
    jsonService.Set "widgets", widgets
End If
Set rs = Nothing

query = "SELECT hook_ID, hook_Shortcode, hook_Desc, hook_Event, hook_Type, hook_Uri, hook_Event, hook_Method, hook_Lang FROM pcHooks WHERE hook_Type <> 'Core' "
set rs = Server.CreateObject("ADODB.Recordset")  
rs.Open query, connTemp, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.Eof Then
    pcv_intTotalCount = rs.RecordCount
    Redim hooks(pcv_intTotalCount - 1)   
    pcv_intCounter = 0
    Do While Not rs.Eof         
        pcv_intID = rs("hook_ID")
        pcv_strShortcode = rs("hook_Shortcode")
        pcv_strDesc = rs("hook_Desc")
        pcv_strType = rs("hook_Type")
        pcv_strUri = rs("hook_Uri")
        pcv_strEvent = rs("hook_Event")
        pcv_strMethod = rs("hook_Method")
        pcv_strLang = rs("hook_Lang")
        
        Dim hook : Set hook = JSON.parse("{}") 
        hook.set "ID", pcv_intID        
        hook.set "ShortCode", pcv_strShortcode
        hook.set "Desc", pcv_strDesc
        hook.set "Event", pcv_strEvent
        hook.set "Type", pcv_strType
        hook.set "Uri", pcv_strUri
        hook.set "Method", pcv_strMethod
        hook.set "Lang", pcv_strLang
        Set hooks(pcv_intCounter) = hook
        Set hook = Nothing
        
        pcv_intCounter = pcv_intCounter + 1
        rs.movenext
    Loop
    jsonService.Set "hooks", hooks
End If
Set rs = Nothing

Response.write( JSON.stringify(jsonService, null, 2) & vbNewline )

call closeDb()
%>