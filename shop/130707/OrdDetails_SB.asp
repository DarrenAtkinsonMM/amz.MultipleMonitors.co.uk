
<% If len(pcv_strGUID)>0 Then %>
	<link type="text/css" rel="stylesheet" href="SubscriptionBridge/css/sb.css" />
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">SubscriptionBridge</th>
	</tr>	
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>	
	<tr>
		<td colspan="2">
			<a href="https://www.subscriptionbridge.com/MerchantCenter/" target="_blank">Manage Subscriptions</a>
		</td>
	</tr>
    <tr>
        <td colspan="2">
            Subscription ID: <strong><%=pcv_strGUID%></strong>
        </td>
    </tr>
    <tr>
        <td colspan="2">
            <%=pcv_strTerms%>
        </td>
    </tr>
<% End If %>

<%
Public Function GetParentOrderID(Guid)

    query="SELECT TOP 1 idOrder FROM SB_Orders WHERE SB_GUID='" & Guid & "' Order By idOrder ASC;"    
    Set rsSB2=Server.CreateObject("ADODB.Recordset")
    Set rsSB2=connTemp.execute(query)
    If NOT rsSB2.eof Then
		GetParentOrderID = rsSB2("idOrder")
	Else 
		GetParentOrderID = 0
    End If
	Set rsSB2 = nothing
	
End Function
%>
