<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=0%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<%
Dim AutoSendActionURL
AutoSendActionURL=""

nmTurnOn=scNM_IsEnabled
nmAuto=scNM_Auto
If IsNull(nmAuto) Or nmAuto="" Then
    nmAuto=0
End If
nmBText=scNM_ButtonText
If nmBText="" Then
    nmBText="Notify In-Stock"
End If

If (nmTurnOn="1") And (nmAuto="1") Then

    'query="SELECT idProduct FROM pcBIS_WaitList;"
    'Set rs=connTemp.execute(query)
    'intCount=-1
    'If Not rs.Eof Then
    '    pcArr=rs.getRows()
    '    intCount=ubound(pcArr,2)
    '    set rs = Nothing
    'End If
    'Set rs = Nothing
    
    'For i=0 to intCount
    '    tmpID=pcArr(0,i)
        %>
        <!--include file="nmSendToServer.asp"-->
        <%
    '    If (tmpSuccess>0) And (tmpErrors=0) Then
    '        call pcs_RmvWaitList(tmpID)
    '    End If
    'Next
    
End If   
   
response.Clear()
If len(AutoSendActionURL)=0 Then
	response.Write("0")
Else
	response.Write(AutoSendActionURL)
End If
response.End()
%>