<%
private const scSB="1"

Dim pSubscriptionID, pcv_intBillingFrequency, pcv_strBillingPeriod, pcv_intBillingCycles, pSubStartImmed,pSubStartFromPurch, pSubStart, pcv_intTrialCycles,pcv_curTrialAmount,pSubStartDate, pSubType, pSubInstall, pcv_intIsTrial, pSubAddToMail, pSubReOccur, pcv_intBillingCyclesUntDate,pcv_intTrialCyclesUntDate
%>

<% If scSBStatus = "1" Then %>

<!--#include file="../includes/pcSBSettings.asp"-->
<!--#include file="../includes/pcSBBase64.asp"-->
<!--#include file="../includes/pcSBHelperInc.asp"-->
<script type="text/javascript" src="<%=pcf_getJSPath(gv_RootURL & "/Widget","widget.js")%>"></script>

<% End If %> 