<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<%
query="SELECT pcPay_PxPay.pcPay_PxPay_PxPayUserId, pcPay_PxPay.pcPay_PxPay_PxPayTestUserId, pcPay_PxPay.pcPay_PxPay_PxPayKey, pcPay_PxPay.pcPay_PxPay_TxnType, pcPay_PxPay.pcPay_PxPay_TestMode, pcPay_PxPay.pcPay_PxPay_CurrencyInput FROM pcPay_PxPay WHERE (((pcPay_PxPay.pcPay_PxPay_ID)=1));"

'======================================================================================
'// End custom query
'======================================================================================

': Create recordset and execute query
set rs=server.CreateObject("ADODB.RecordSet") 'DELETE FOR HARD CODED VARS
set rs=connTemp.execute(query) 'DELETE FOR HARD CODED VARS

': Capture any errors
if err.number<>0 then 'DELETE FOR HARD CODED VARS
	call LogErrorToDatabase() 'DELETE FOR HARD CODED VARS
	set rs=nothing 'DELETE FOR HARD CODED VARS
	call closedb() 'DELETE FOR HARD CODED VARS
	response.redirect "techErr.asp?err="&pcStrCustRefID 'DELETE FOR HARD CODED VARS
end if 'DELETE FOR HARD CODED VARS

'======================================================================================
'// Set gateway specific variables - hard code is not using database to store gateway
'// information
'======================================================================================
pcv_PxPayUserId=rs("pcPay_PxPay_PxPayUserId")
pcv_PxPayTestUserId=rs("pcPay_PxPay_PxPayTestUserId")
pcv_PxPayKey=rs("pcPay_PxPay_PxPayKey")
pcv_TxnType=rs("pcPay_PxPay_TxnType")
pcv_CurrencyInput=rs("pcPay_PxPay_CurrencyInput")
pcv_TestMode=rs("pcPay_PxPay_TestMode")
if pcv_TestMode=1 then
	pcv_PxPayUserId=pcv_PxPayTestUserId
end if
'======================================================================================
'// End gateway specific variables
'======================================================================================

dim tempURL
If scSSL="" OR scSSL="0" Then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
Else
	tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/gwSubmit.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://")
End If

dim pcv_Result
pcv_Result = Request.QueryString ("result")

sXmlAction = sXmlAction & "<ProcessResponse>"
sXmlAction = sXmlAction & "<PxPayUserId>"& pcv_PxPayUserId &"</PxPayUserId>"
sXmlAction = sXmlAction & "<PxPayKey>"& pcv_PxPayKey &"</PxPayKey>"
sXmlAction = sXmlAction & "<Response>" & pcv_Result &"</Response>"
sXmlAction = sXmlAction & "</ProcessResponse>"	

Dim objXMLhttp 
Set objXMLhttp = server.Createobject("MSXML2.XMLHTTP") 

objXMLhttp.Open "POST", "https://www.paymentexpress.com/pxpay/pxaccess.aspx" ,False
objXMLhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objXMLhttp.send sXmlAction

Dim oXML, URI
Set oXML = Server.CreateObject("MSXML2.DomDocument")
oXML.loadXML(objXMLhttp.responseText)
strSuccess = oXML.selectSingleNode("//Success").text
strMerchantReference = oXML.selectSingleNode("//MerchantReference").text
strAuthCode = oXML.selectSingleNode("//AuthCode").text
strTxnId = oXML.selectSingleNode("//TxnId").text
strDpsTxnRef = oXML.selectSingleNode("//DpsTxnRef").text

if strSuccess="1" then
	response.write "Approved<hr>"
	session("GWOrderID")=strTxnId
	session("GWAuthCode")=strAuthCode
	session("GWTransId")=strDpsTxnRef
	'Redirect to complete order
	response.redirect "gwReturn.asp?s=true&gw=PxPay"
	
	response.write "Merchant Reference: "&strMerchantReference&"<BR>"
	response.write "AuthCode: "&strAuthCode&"<BR>"
	response.write "Trans ID: "&strTxnId&"<BR>"
	response.write "DPS Ref. Code: "&strDpsTxnRef&"<BR>"
end if

if strSuccess="0" then
	strDpsTxnRef = oXML.selectSingleNode("//ResponseText").text
	strTxnId = oXML.selectSingleNode("//TxnId").text
	strAmountSettlement = oXML.selectSingleNode("//AmountSettlement").text
	
    Session("message") = strDpsTxnRef
    Session("backbuttonURL") = tempURL & "?psslurl="  &session("redirectPage") & "&idCustomer=" & session("idCustomer") & "&idOrder=" & session("GWOrderID") & "&ordertotal=" & strAmountSettlement
    response.redirect "msgb.asp?back=1"

end if

Set objXMLhttp = nothing
%>
<!--#include file="footer_wrapper.asp"-->
