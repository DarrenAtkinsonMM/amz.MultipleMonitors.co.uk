<!--#include file="../includes/common_init.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/themesettings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/securitysettings.asp" -->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/bto_language.asp"--> 
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/pcSeoLinks.asp"-->
<!--#include file="../includes/pcSBClassInc.asp"-->
<!--#include file="../includes/utilities/json2.asp"-->
<!--#include file="../includes/SearchConstants.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcMobileSettings.asp"-->
<!--#include file="../includes/coreMethods/content.asp"-->
<%
Function openDB()
    On Error Resume Next
    Set connTemp = server.createobject("adodb.connection")
    connTemp.Open scDSN  
    If err.number <> 0 Then
	    response.redirect "dbError.asp"
	    response.End()
    End If
End Function

Function closeDB()
    On Error Resume Next
    connTemp.close
    Set connTemp = nothing
End Function


'// START:  HIDE SIDE BAR  
pcv_CurrentPageName = Request.ServerVariables("SCRIPT_NAME")
    loc = instrRev(pcv_CurrentPageName,"/") 
    pcv_CurrentPageName = mid(pcv_CurrentPageName, loc+1, len(pcv_CurrentPageName) - loc) 
pcv_NoSideNavePageList = "|ONEPAGECHECKOUT.ASP|"
'// END:  HIDE SIDE BAR 
%>

<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

viewcartbtn = RSlayout("viewcartbtn")

'// Load validation resources
pcv_strRequiredIcon = rsIconObj("requiredicon")
pcv_strErrorIcon = rsIconObj("errorfieldicon")

conlayout.Close
%>