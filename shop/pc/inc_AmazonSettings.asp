<%
pcAmazonTurnOn="0"
IF (Ucase(pcStrPageName)="GWAMAZONMWS.ASP") OR (Ucase(pcStrPageName)="ORDERCOMPLETE.ASP") OR (Ucase(pcStrPageName)="PCPAY_AMAZON_START.ASP") OR (Ucase(pcStrPageName)="VIEWCART.ASP") OR (Ucase(pcStrPageName)="ONEPAGECHECKOUT.ASP") OR (Ucase(pcStrPageName)="OPC_AMZUPDSHIPADDR.ASP") THEN
Set conAmz=Server.CreateObject("ADODB.Connection")
conAmz.Open scDSN
Set rsAmz = conAmz.Execute("SELECT idPayment FROM paytypes WHERE active=-1 AND gwCode=88;")
if not rsAmz.eof then
	Set rsAmz=nothing
	Set rsAmz = conAmz.Execute("SELECT gwAMZ_SellerID,gwAMZ_AccessKey,gwAMZ_SecretKey,gwAMZ_ClientID,gwAMZ_ClientSecret,gwAMZ_Mode,gwAMZ_TestMode FROM gwAmazon;")
	if not rsAmz.eof then
		pcAmazonTurnOn="1"
		x_S=rsAmz("gwAMZ_SellerID")
		x_S=enDeCrypt(x_S, scCrypPass)
		x_Login=rsAmz("gwAMZ_AccessKey")
		x_Login=enDeCrypt(x_Login, scCrypPass)
		x_Key=rsAmz("gwAMZ_SecretKey")
		x_Key=enDeCrypt(x_Key, scCrypPass)
		x_C=rsAmz("gwAMZ_ClientID")
		x_C=enDeCrypt(x_C, scCrypPass)
		x_CS=rsAmz("gwAMZ_ClientSecret")
		x_CS=enDeCrypt(x_CS, scCrypPass)
		x_Mode=rsAmz("gwAMZ_Mode")
		x_testmode=rsAmz("gwAMZ_TestMode")
	end if
	If x_testmode="1" then
		pcAMZEndPoint="https://mws.amazonservices.com/OffAmazonPayments_Sandbox/2013-01-01"
		pcAMZWidURL="https://static-na.payments-amazon.com/OffAmazonPayments/us/sandbox/js/Widgets.js?sellerId="
		pcAMZHost="mws.amazonservices.com"
		pcAMZUI="/OffAmazonPayments_Sandbox/2013-01-01"
		pcAMZAPI="https://api.sandbox.amazon.com/"
	Else
		pcAMZEndPoint="https://mws.amazonservices.com/OffAmazonPayments/2013-01-01"
		pcAMZWidURL="https://static-na.payments-amazon.com/OffAmazonPayments/us/js/Widgets.js?sellerId="
		pcAMZHost="mws.amazonservices.com"
		pcAMZUI="/OffAmazonPayments/2013-01-01"
		pcAMZAPI="https://api.amazon.com/"
	End if
	
	pcAMZSellerID=x_S
	pcAMZAccessKeyID=x_Login
	pcAMZSecretKey=x_Key
	pcAMZClientID=x_C
	pcAMZClientSecret=x_CS
end if
set rsAmz=nothing
Set conAmz=nothing
END IF%>