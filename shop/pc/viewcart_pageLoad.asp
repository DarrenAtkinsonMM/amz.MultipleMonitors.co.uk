<script type=text/javascript>

    // Remove ASP Code from JavaScript code    
    var pcCartIndex = <%=pcCartIndex%>;      
    var alert_recal = "<%=dictLanguage.Item(Session("language")&"_alert_recal")%>";
    var showcart_2 = "<%=dictLanguage.Item(Session("language")&"_showcart_2")%>";
    var alert_3 = "<%=dictLanguage.Item(Session("language")&"_alert_3")%>";
    var alert_4 = "<%=dictLanguage.Item(Session("language")&"_alert_4")%>";
    var alert_5 = "<%=dictLanguage.Item(Session("language")&"_alert_5")%>";
    var alert_8 = "<%=dictLanguage.Item(Session("language")&"_alert_8")%>";
    var alert_8b = "<%=dictLanguage.Item(Session("language")&"_alert_8b")%>";
    var alert_9 = "<%=dictLanguage.Item(Session("language")&"_alert_9")%>";

    var SaveCart_3 = "<%=dictLanguage.Item(Session("language")&"_SaveCart_3")%>";
    var SaveCart_4 = "<%=dictLanguage.Item(Session("language")&"_SaveCart_4")%>";
    var SaveCart_5 = "<%=dictLanguage.Item(Session("language")&"_SaveCart_5")%>";
    
    var RemainIssue="";
    var RemainIssue1="";

</script>

<%
If EditSB = 1 Then
    pcv_CheckoutButtonImage = RSlayout("pcLO_Update")
Else
    pcv_CheckoutButtonImage = RSlayout("checkout")
End If
%>