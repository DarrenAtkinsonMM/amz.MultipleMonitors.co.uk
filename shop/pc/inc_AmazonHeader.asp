<!--#include file="inc_AmazonSettings.asp" -->
<%if pcAmazonTurnOn="1" then%>
<script type=text/javascript>
	<%if (Ucase(pcStrPageName)="ONEPAGECHECKOUT.ASP") OR (Ucase(pcStrPageName)="ORDERCOMPLETE.ASP") then%>
	var accessToken = "<%=session("Amz_access_token")%>";
	if (typeof accessToken === 'string' && accessToken.match(/^Atza/)) {
    document.cookie = "amazon_Login_accessToken=" + accessToken + ";secure";}
	<%end if%>
	window.onAmazonLoginReady = function() {
    amazon.Login.setClientId('<%=pcAMZClientID%>');
	<%if (Ucase(pcStrPageName)="ONEPAGECHECKOUT.ASP") OR (Ucase(pcStrPageName)="ORDERCOMPLETE.ASP") then%>
	amazon.Login.setUseCookie(true);
	<%end if%>
  };
</script>
<script type=text/javascript src='<%=pcAMZWidURL & pcAMZSellerID%>'></script>
<%end if%>