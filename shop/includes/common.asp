<%
'// SHARED FILES
%>
<!--#include file="opendb.asp"--> 
<!--#include file="common_init.asp"-->
<!--#include file="SQLFormat.txt"-->
<!--#include file="settings.asp"-->
<!--#include file="themesettings.asp"-->
<!--#include file="emailsettings.asp"-->
<!--#include file="taxsettings.asp"-->
<!--#include file="securitysettings.asp" -->
<!--#include file="storeconstants.asp"-->
<!--#include file="languages.asp"--> 
<!--#include file="bto_language.asp"--> 
<!--#include file="rewards_language.asp"-->
<!--#include file="productcartFolder.asp"-->
<!--#include file="productcartinc.asp"--> 
<!--#include file="dateinc.asp"-->
<!--#include file="currencyformatinc.asp"-->
<!--#include file="stringfunctions.asp"-->
<!--#include file="adovbs.inc"-->
<!--#include file="ErrorHandler.asp"-->
<!--#include file="rc4.asp" -->
<!--#include file="pcSeoLinks.asp"-->
<!--#include file="pcSBClassInc.asp"-->
<!--#include file="secureadminFolder.asp" -->
<!--#include file="settingsPCL.asp" -->
<!--#include file="settingsPCWS.asp" -->
<!--#include file="encrypt.asp"-->
<!--#include file="defenderSettings.asp" -->

<%
'// FEATURES
%>
<!--#include file="extendedMethods/apps.asp"-->
<!--#include file="extendedMethods/reCaptcha.asp"-->
<!--#include file="extendedMethods/Avalara.asp"-->

<%
'// FILE VERSIONS
%>
<!--#include file="../includes/status.inc"-->
<!--#include file="../includes/statusAPP.inc"-->
<!--#include file="../includes/statusCM.inc"-->
<!--#include file="../includes/statusM.inc"-->
<!--#include file="../includes/statusPCL.inc"-->
<!--#include file="../includes/ppdstatus.inc"-->

<%
'// UTILS
%>
<!--#include file="../pc/inc_mobiledetect.asp"-->
<!--#include file="utilities/json2.asp"-->

<%
'// GLOBAL "PRE-HEADER" 
call pcs_SetCodePage()
%>