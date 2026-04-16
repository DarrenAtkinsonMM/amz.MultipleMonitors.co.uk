<%
call openDb()
%>
<%

	paymentImageDir = session("pcsTheme") & "/images/payment"
	
	Set connTemp2=Server.CreateObject("ADODB.Connection")
	connTemp2.Open scDSN
	
	'// Get Accepted Payment types
	query = "SELECT pcAcceptedPayment_Name, pcAcceptedPayment_Image, pcAcceptedPayment_CustomImage, pcAcceptedPayment_Alt FROM pcAcceptedPayments WHERE pcAcceptedPayment_Active = 1 ORDER BY pcAcceptedPayment_Order, pcAcceptedPayment_Name"
	set rs=connTemp2.execute(query)
	if not rs.eof then
		pcAcceptedPayments = rs.GetRows()
		pcAcceptedPaymentsCnt = UBound(pcAcceptedPayments, 2) + 1
	else
		pcAcceptedPaymentsCnt = 0
	end if
	set rs = nothing
	
	set connTemp2 = nothing
%>

<% If pcAcceptedPaymentsCnt > 0 Then %>
  <ul id="pcAcceptedPayments">
  
  <% 
		For i = 0 To pcAcceptedPaymentsCnt - 1
			paymentName = pcAcceptedPayments(0, i)
			paymentImage = pcAcceptedPayments(1, i)
			paymentCustomImage = pcAcceptedPayments(2, i)
			paymentAlt = pcAcceptedPayments(3, i)
			
			If Len(paymentAlt) < 1 Then
				paymentAlt = paymentName
			End If
			
			If Len(paymentCustomImage) > 0 Then
				paymentImage = paymentCustomImage
				paymentDir = "catalog"
			Else
				paymentDir = paymentImageDir
			End If
	%>
  	<li class="pcAcceptedPayment<%= Replace(paymentName, " ", "") %>">
    	<img src="<%=pcf_getImagePath(paymentDir,paymentImage)%>" alt="<%= paymentAlt %>" title="<%= paymentAlt %>" />
    </li>
  <% Next %>
  </ul>
<% End If %>
<%
call closeDb()
%>