<!--#include file="../../shop/includes/pcMobileSettings.asp"-->
<!--#include file="include-metatags.asp"-->
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<% If pcv_PageName <> "" Then %>
<title itemprop="name"><%=pcv_PageName%></title>
<% End If %>
<% GenerateMetaTags() %>
<% Response.Buffer=True %> 
<%
Session("pcStrPageName") = pcStrPageName
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
	
viewcartbtn = RSlayout("viewcartbtn")

'// Load validation resources
pcv_strRequiredIcon = rsIconObj("requiredicon")
pcv_strErrorIcon = rsIconObj("errorfieldicon")
%>
<%
private const scIncHeader="1"

'////////////////////////////////////////////////////////////////////
'// START - Select Theme
'////////////////////////////////////////////////////////////////////
Dim pcv_theme
pcv_theme = scThemePath
If len(pcv_theme)=0 Then
	pcv_theme = scThemeFolder
End If

If len(session("userTheme")) > 0 Then
    pcv_theme = session("userTheme")
End If

'// Example: "sandbox.asp?to=" & Server.URLEncode("theme/basic_blue")
pcv_strThemeOverride = getUserInput(Request("theme"), 0)
If len(pcv_strThemeOverride)>0 Then
    pcv_theme = "/shop/pc/theme" & "/" & pcv_strThemeOverride
End If

If len(pcv_theme)=0 Then
    pcv_theme = "/shop/pc/theme/basic_blue"
End If
session("pcsTheme")=pcv_theme
'////////////////////////////////////////////////////////////////////
'// END - Select Theme
'////////////////////////////////////////////////////////////////////
%>
<%
'////////////////////////////////////////////////////////////////////
'// START - Check for WWW and redirect if absent
'////////////////////////////////////////////////////////////////////
Dim strDomain, strPath, strQueryString, strURL, strHttpsDomain, intDoRedirect, intRedirectType

'// Redirect to maintain consistent URL
call storeURLRedirect()
'////////////////////////////////////////////////////////////////////
'// END - Check for WWW and redirect if absent
'////////////////////////////////////////////////////////////////////
%>
<%'// START: CSS %>
<!--#include file="inc_headerCSS.asp"-->
<%'// END: CSS %>

<%'// START: SNIPPET PREREQUISITES %>
<!--#include file="smallRecentProductsCookie.asp"-->
<%'// END: SNIPPET PREREQUISITES %>


<%'// START: JAVASCRIPT %>
<!--#include file="inc_jquery.asp" -->
<!--#include file="inc_sb.asp"-->
<!--#include file="inc_AmazonHeader.asp" -->
<%'// END: JAVASCRIPT %>

<%call pcs_genReCaHeader()%>

<%
'// START: v4.5 Built-in Integration with Google Analytics
if trim(scGoogleAnalytics)<>"" and not IsNull(scGoogleAnalytics) then
if scGAType="1" then '//Google Universal Analytics%>
<script type=text/javascript>
	(function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
	(i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
	m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
	})(window,document,'script','//www.google-analytics.com/analytics.js','ga');
	
	ga('create', '<%=scGoogleAnalytics%>', 'auto');
	ga('send', 'pageview');
</script>
<% elseif scGAType="0" then %>
<script type=text/javascript>

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', '<%=scGoogleAnalytics%>']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
<!-- Google Tag Manager -->
<script>(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
j=d.createElement(s),dl=l!='dataLayer'?'&l='+l:'';j.async=true;j.src=
'https://www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
})(window,document,'script','dataLayer','GTM-KRN8HZZ');</script>
<!-- End Google Tag Manager -->	
	
<%end if
end if
'// END: v4.5 Built-in Integration with Google Analytics


'// START: Redirect if JavaScript is Disabled
If pcStrPageName<>"noscript.asp" Then
    If (pcStrPageName = "configurePrd.asp") Or (pcStrPageName = "Reconfigure.asp") Then
        %>
        <noscript>
            <meta http-equiv="refresh" content="0;URL=noscript.asp"/>
        </noscript>
        <%
    End If
End If
'// END: Redirect if JavaScript is Disabled
%>