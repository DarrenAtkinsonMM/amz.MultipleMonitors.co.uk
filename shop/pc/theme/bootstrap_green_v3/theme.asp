<!--#include file="../../inc_theme_common.asp"--> 
<link href="https://netdna.bootstrapcdn.com/font-awesome/4.0.3/css/font-awesome.css" rel="stylesheet">
<link href='https://fonts.googleapis.com/css?family=Open+Sans:300italic,400italic,600italic,700italic,800italic,400,300,600,700,800' rel='stylesheet' type='text/css'>
<link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath(session("pcsTheme") & "/css","theme.min.css")%>" />
<% If session("Mobile")="1" Then %>
	<link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath(session("pcsTheme") & "/css","theme_mobile.css")%>" />
<% Else %>    
	<link rel="stylesheet" href="<%=pcf_getCSSPath(session("pcsTheme") & "/css","superfish.min.css")%>" media="screen">
<% End If %>