<%IF pcAmazonTurnOn="1" THEN
	SPath1=Request.ServerVariables("PATH_INFO")
	mycount1=0
	do while mycount1<1
		if mid(SPath1,len(SPath1),1)="/" then
			mycount1=mycount1+1
		end if
		if mycount1<1 then
			SPath1=mid(SPath1,1,len(SPath1)-1)
		end if
	loop
	SPathInfo="https://"
	SPathInfo=SPathInfo & Request.ServerVariables("HTTP_HOST") & SPath1%>
	<div id="pcAmazonButtons" class="pcAltCheckoutButtons">
		<div id="AmazonPayButton" class="pcAltCheckoutButton"></div>
		<script type=text/javascript>
			var authRequest;
			OffAmazonPayments.Button("AmazonPayButton", "<%=pcAMZSellerID%>", {
				type: "PwA",
				authorization: function() {
					loginOptions =
					{scope: "profile payments:widget payments:shipping_address", popup: false};
					authRequest = amazon.Login.authorize (loginOptions, "<%=SPathInfo%>pcPay_Amazon_Start.asp");
				},
				onError: function(error) {
					// your error handling code
					alert("<%=dictLanguage.Item(Session("language")&"_AmazonPay_6")%>");
					document.location="viewcart.asp";
				}
			});
		</script>
	</div>
<%END IF%>