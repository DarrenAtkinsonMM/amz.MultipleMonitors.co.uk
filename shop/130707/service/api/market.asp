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
Response.ContentType = "text/html"
Response.Charset = "UTF-8"

query="SELECT pcPCWS_Uid, pcPCWS_AuthToken, pcPCWS_Username, pcPCWS_Password FROM pcWebServiceSettings;"
Set rs=connTemp.execute(query)
If Not rs.eof Then
    pcv_strUid = rs("pcPCWS_Uid")
    pcv_AuthToken = rs("pcPCWS_AuthToken")  
    pcv_strUsername = rs("pcPCWS_Username")  
    pcv_strPassword = enDeCrypt(rs("pcPCWS_Password"), scCrypPass)          
End If
Set rs=nothing

'response.Write(pcv_AuthToken)
'response.End()

pcv_strViewParams = getUserInput(Request("params"), 0)
pcv_strView = pcf_displayMarket(pcv_strViewParams, pcv_AuthToken)
response.Write(pcv_strView)
%>