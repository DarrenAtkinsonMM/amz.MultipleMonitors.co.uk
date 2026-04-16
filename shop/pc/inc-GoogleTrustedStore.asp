<%
gtsPageLang="en_US"
gtsShopID=""
gtsShopAccID=""
gtsCountry="US"
gtsLang="en"
queryQ="SELECT pcGTS_TurnOn,pcGTS_AccNo,pcGTS_PageLang,pcGTS_ShopAccID,pcGTS_ShopCountry,pcGTS_ShopLang FROM pcGoogleTS;"
set rsQ=connTemp.execute(queryQ)
if not rsQ.eof then
	gtsTurnOn=rsQ("pcGTS_TurnOn")
	gtsAccNo=rsQ("pcGTS_AccNo")
	gtsPageLang=rsQ("pcGTS_PageLang")
	gtsShopAccID=rsQ("pcGTS_ShopAccID")
	gtsCountry=rsQ("pcGTS_ShopCountry")
	gtsLang=rsQ("pcGTS_ShopLang")
end if
set rsQ=nothing

if (gtsTurnOn="1") then%>
<!-- BEGIN: Google Trusted Stores -->
<script type="text/javascript">
var gts = gts || [];
gts.push(["id", "<%=gtsAccNo%>"]);
gts.push(["badge_position", "USER_DEFINED"]);
gts.push(["badge_container", "GTS_CONTAINER"]);
gts.push(["locale", "<%=gtsPageLang%>"]);
<%if gtsShopAccID<>"" then%>
gts.push(["google_base_subaccount_id", "<%=gtsShopAccID%>"]);
<%end if%>
gts.push(["google_base_country",  "<%=gtsCountry%>"]);
gts.push(["google_base_language",  "<%=gtsLang%>"]);
(function() {
var gts = document.createElement("script");
gts.type = "text/javascript";
gts.async = true;
gts.src = "https://www.googlecommerce.com/trustedstores/api/js";
var s = document.getElementsByTagName("script")[0];
s.parentNode.insertBefore(gts, s);
}) ();
</script>
<!-- END: Google Trusted Stores -->
<%end if%>

