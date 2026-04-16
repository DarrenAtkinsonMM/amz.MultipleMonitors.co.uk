<!--#include file="../../inc_theme_common.asp"--> 
<% If session("Mobile")="1" Then %>
<link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath("css","slidebars.min.css")%>" />
<link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath(session("pcsTheme") & "/css","theme_mobile.css")%>" />
<% Else %>
<link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath(session("pcsTheme") & "/css","theme.css")%>" />
<% End If %>