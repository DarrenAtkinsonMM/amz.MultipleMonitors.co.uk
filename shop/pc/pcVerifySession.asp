<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'*******************************
' START Verify Session
'*******************************
Dim pcCartIndex
Private Sub pcs_VerifySession
	Dim pcv_strCatcher
	pcv_strCatcher = Session("pcCartIndex")
	If pcv_strCatcher=0 Then
		pcv_strCatcher=""
		Session("Cust_BuyGift")=""
		pcv_strCheckSession = getUserInput(Request("cs"),1)
		if (len(pcv_strCheckSession)>0) AND (session("pcSessionID") <> Session.SessionID) then
		 	response.redirect "msg.asp?message=212" '// enable cookies
		else
            pcv_strScriptName = Request.ServerVariables ("SCRIPT_NAME")
            If len(pcv_strScriptName)>0 Then
                If Not instr(pcv_strScriptName,"service.asp")>1 Then
                    response.redirect "msg.asp?message=1" '// cart empty
                End If
            Else
			    response.redirect "msg.asp?message=1" '// cart empty
            End IF
		end if
	End If
	pcCartArray=Session("pcCartSession")
	pcCartIndex=Session("pcCartIndex")
End Sub
'*******************************
' START Verify Session
'*******************************
%>
