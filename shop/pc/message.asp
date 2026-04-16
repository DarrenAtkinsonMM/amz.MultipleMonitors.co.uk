<!--#include file="../includes/common.asp"-->

<%

query = "SELECT TOP 1 bannerid, startdate, enddate, active, background, html FROM mod_bannermanagement WHERE GETDATE() BETWEEN startdate AND datediff(d,0, enddate)"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
  pcStrHPBannerText=rs("html")
  pcStrHPBannerColor=rs("background")
%>

<div class="message alert" style="background-color:<%=pcStrHPBannerColor%>">
  <%=pcStrHPBannerText%>
</div>

<%
end if
%>

