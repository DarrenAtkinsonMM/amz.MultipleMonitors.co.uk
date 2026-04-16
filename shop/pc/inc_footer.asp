

<!--#include file="../includes/common_javascripts.asp"-->


 
<%'// START: JS %>
<!--#include file="inc_footerJS.asp"-->
<%'// END: JS %>

<%
private const scIncFooter="1"

dim tempFooterURL
tempFooterURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
tempFooterURL=replace(tempFooterURL,"https:/","https://")
tempFooterURL=replace(tempFooterURL,"http:/","http://")
%>
<%
'// Restore Cart
if session("NeedToShowRSCMsg")="1" then
	session("NeedToShowRSCMsg")=""

	'Get Path Info
	pcv_filePath = Request.ServerVariables("PATH_INFO")
	do while instr(pcv_filePath,"/")>0
		pcv_filePath = mid(pcv_filePath,instr(pcv_filePath,"/")+1,len(pcv_filePath))
	loop

	pcv_Query = Request.ServerVariables("QUERY_STRING")
	If len(pcv_Query)>0 Then
		If instr(pcv_filePath,"404.asp")>0 AND instr(pcv_Query,"404;")>0 Then
			pcv_filePath = Right(pcv_Query,Len(pcv_Query)-4)
		Else
			pcv_filePath = pcv_filePath & "?" & pcv_Query
		end if
	End If

	session("SFClearCartURL")=pcv_filePath
	%>
    
    <%
    pcv_strButtons = ""
    pcv_strButtons = pcv_strButtons & "<a href=""" & tempFooterURL & "CustLOb.asp"" role=""button"" class=""btn btn-default"">" & dictLanguage.Item(Session("language")&"_opc_js_77") & "</a>"
    pcv_strButtons = pcv_strButtons & "<a href=""" & tempFooterURL & "viewcart.asp"" role=""button"" class=""btn btn-default"">" & dictLanguage.Item(Session("language")&"_opc_js_65") & "</a>"
    pcv_strButtons = pcv_strButtons & "<a role=""button"" class=""btn btn-default"" data-dismiss=""modal"">" & dictLanguage.Item(Session("language")&"_opc_js_78") & "</a>"
    %>

<% else
	session("SFClearCartURL")=""
end if
session("MobileURL")=""
session("idProductRedirect")=""
session("pcsTheme")=""
session("pcv_strCSFilters") = ""
session("pcv_strCSFieldQuery")
%>
<!--#include file="inc-GoogleAnalytics.asp"-->
<% if pcInterest=1 then %>
<script type="text/javascript" src="<%=pcf_getJSPath("//assets.pinterest.com/js","pinit.js")%>"></script>
<% end if %>

<% 
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing 

If scCartStack_IsEnabled = "1" Then
    If len(pcStrPageName)=0 Then
        pcStrPageName = "no_page"
    End If
	If lCase(pcStrPageName) = "viewcart.asp" OR lCase(pcStrPageName) = "onepagecheckout.asp" Then
		pcs_CartStackTracking
	ElseIf lCase(pcStrPageName) = "ordercomplete.asp" Then
		pcs_CartStackConfirmation
	End If
End If
%>
<!--#include file="inc-GoogleTrustedStore.asp"-->