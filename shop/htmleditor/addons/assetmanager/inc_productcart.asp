<!-- #include file="../../../includes/common.asp" -->
<!-- #include file="../../../includes/pcSanitizeUpload.asp" -->
<%
response.Buffer=true

'// Verifies if admin is logged, so as not send to login page
pcv_strAdminPrefix="1"

If (session("admin") = 0 OR session("admin") = 1 OR session("admin") = "") _
	OR _
	((instr(session("PmAdmin"),"*")=0 And instr(session("PmAdmin"),"19")=0)) _
	OR _
	(len(session("CUID"))=0) _
	OR _
	(session("admin." & pcf_getAdminToken()) <> Session.SessionID) Then
	
	If InStr(Request.ServerVariables("SCRIPT_NAME"), "upload.asp") = 0 Then
		Response.Write("You do not have proper rights to access this page.")
		Response.End()
	Else
		ShowErrorMessage "", "", "NOTALLOWED"
		Response.End()
	End If

End If
imageBaseFolder = ""

If PPD="1" Then
	imageFullPath = "true"
	If Len(scPcFolder) > 0 Then imageBaseFolder = imageBaseFolder & "/" & scPcFolder
Else
	imageFullPath = "false"
	imageBaseFolder = imageBaseFolder & "../../.."
End If

imageBaseFolder = imageBaseFolder & "/pc/catalog"

catalogURL = replace((scStoreURL&"/"&scPcFolder&"/pc/catalog/"),"//","/")
catalogURL = replace(catalogURL,"http:/","http://") 
catalogURL = replace(catalogURL,"https:/","https://") 

Private Function pcf_getAdminToken()
	pcv_strLocalAddress = Request.ServerVariables("LOCAL_ADDR") 
	pcv_strLocalSessionID = Session.SessionID
	pcv_strAdminToken = pcv_strLocalAddress & "." & pcv_strLocalSessionID
	pcf_getAdminToken = pcv_strAdminToken
End Function
%>