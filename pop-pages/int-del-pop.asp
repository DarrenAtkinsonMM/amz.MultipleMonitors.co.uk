<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#INCLUDE FILE="../shop/includes/settings.asp"-->
<!--#INCLUDE FILE="../shop/includes/storeconstants.asp"-->
<!--#INCLUDE FILE="../shop/includes/opendb.asp"-->
<!--#INCLUDE FILE="../shop/includes/adovbs.inc"-->
<!--#INCLUDE FILE="../shop/includes/stringfunctions.asp"-->
<!--#include FILE="../shop/includes/ErrorHandler.asp"--> 
<%
	if not validNum(request.querystring("prodid")) then
		daEditProdID = 333
	else
		daEditProdID = request.querystring("prodid")
	end if
	
	call opendb()
	
    'Get bundle details based on query string
	
	query="SELECT sku,weight FROM products WHERE idProduct = " & daEditProdID
    set rs=server.CreateObject("ADODB.RecordSet")
    set rs=conntemp.execute(query)
	
	strOutput = ""
	
	while not rs.eof
		intProdWeight = rs("weight")
		strSKU = rs("sku") 
		rs.movenext
	wend
		
	set rs = nothing
	
	'Check if bundle
	if not CInt(request.QueryString("sid")) > 0 then
		strType = "computer"
	else
		'We have a bundle
		strType = "bundle"
		
		'Work out number of screens based on stand
		Select Case CInt(request.QueryString("sid"))
			Case 326
				intMons = 2
				intStdWeight = 10
			Case 287
				intMons = 2
				intStdWeight = 10
			Case 324
				intMons = 3
				intStdWeight = 10
			Case 312
				intMons = 3
				intStdWeight = 10
			Case 313
				intMons = 4
				intStdWeight = 12
			Case 327
				intMons = 4
				intStdWeight = 12
			Case 325
				intMons = 4
				intStdWeight = 12
			Case 318
				intMons = 5
				intStdWeight = 12
			Case 314
				intMons = 6
				intStdWeight = 12
			Case 319
				intMons = 8
				intStdWeight = 16
		End Select
		
		'Work out weight of screens
		Select Case CInt(request.QueryString("mid"))
			Case 315
				intMonWeight = 4
			Case 316
				intMonWeight = 4
			Case 304
				intMonWeight = 4
			Case 321
				intMonWeight = 4
			Case 320
				intMonWeight = 4
			Case 317
				intMonWeight = 4
			Case 328
				intMonWeight = 8
			Case 329
				intMonWeight = 8
		End Select
		
		'Add stand, screen and PC weight
		intProdWeight = intProdWeight + intStdWeight + (intMonWeight * intMons)
		
	end if
	
	
	
	
	'Set computer / bundle lead time
	intLeadTime = 7
	
	if instr(strSKU, "MM-M") then
		strType = "monitor"
		intLeadTime = 5
	end if
		
	if instr(strSKU, "MM-S") then
		strType = "stand"
		intLeadTime = 5
	end if
	
	'Create Drop down list
	strOutput = strOutput & "<select id=""delCountry"" name=""delCountry"" style=""margin-top: 3px;"" onchange=""reCalc()""><option value=""0"" title=""0"">Select Destination</option>"	
	
	query="SELECT * FROM FlatShipTypeRules WHERE quantityFrom < " & intProdWeight & " AND quantityTo > " & intProdWeight & " ORDER BY idFlatshipType"
	set rs=server.CreateObject("ADODB.RecordSet")
    set rs=conntemp.execute(query)
	
	i = 0
	
	while not rs.eof
		i = i + 1
		Select Case rs("idFlatshipType")
			Case 4
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Northern Ireland</option>"
			Case 7
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Republic of Ireland</option>"
			Case 8
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Austria</option>"
			Case 9
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Belgium</option>"
			Case 10
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Denmark</option>"
			Case 11
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>France</option>"
			Case 12
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Germany</option>"
			Case 13
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Italy</option>"
			Case 14
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Luxembourg</option>"
			Case 15
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Monaco</option>"
			Case 16
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Netherlands</option>"
			Case 17
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Norway</option>"
			Case 18
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Portugal</option>"
			Case 19
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Spain</option>"
			Case 20
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Sweden</option>"
			Case 21
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Switzerland</option>"
			Case 24
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Malta</option>"
			Case 28
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Greece</option>"
			Case 29
				strOutput = strOutput & "<option value=""" & rs("idFlatshipType") & """ title=""" & rs("shippingPrice") & """>Finland</option>"
		End Select
		rs.movenext
	wend
	
	strOutput = strOutput & "</select>"

	set rs = nothing	
		
%>
<script src="/js/moment.min.js"></script>
<SCRIPT LANGUAGE="JavaScript">
var numDel = <%=intLeadTime%>;

function reCalc() {
var numSubTotal = 0;
var numDelTime = 0;
	
	var ddlop1 = document.getElementById('delCountry');
	if (ddlop1 != null)
	{
	for (var i = 0; i < ddlop1.options.length; i++) {
		if (ddlop1.options[i].selected == true) {
			numSubTotal = eval(ddlop1.options[i].title);
			numCountry = eval(ddlop1.options[i].value);
		}
	}
	}
	
	numSubTotal = numSubTotal.toFixed(2);
	
	//sort out dates based on delivery destination	
	switch(numCountry)
	{
	case 4:
		//Northern Ireland
		numDelTime = 2;
		break;
	case 7:
		//Republic of Ireland
		numDelTime = 2;
		break;
	case 8:
		//Austria
		numDelTime = 3;
		break;
	case 9:
		//Belgium
		numDelTime = 2;
		break;
	case 10:
		//Denmark
		numDelTime = 3;
		break;
	case 11:
		//France
		numDelTime = 2;
		break;
	case 12:
		//Germany
		numDelTime = 2;
		break;
	case 13:
		//Italy
		numDelTime = 4;
		break;
	case 14:
		//Luxembourg
		numDelTime = 2;
		break;
	case 15:
		//Monaco
		numDelTime = 2;
		break;
	case 16:
		//Netherlands
		numDelTime = 2;
		break;
	case 17:
		//Norway
		numDelTime = 6;
		break;
	case 18:
		//Portugal
		numDelTime = 5;
		break;
	case 19:
		//Spain
		numDelTime = 4;
		break;
	case 20:
		//Sweden
		numDelTime = 5;
		break;
	case 21:
		//Switzerland
		numDelTime = 3;
		break;
	case 28:
		//Greece
		numDelTime = 6;
		break;
	case 29:
		//Finland
		numDelTime = 5;
		break;
	}
		
	//add 2 extra days on to delivery time as we will definitely cross a weekend
	//numDel = numDel + 2;

	var dtDelivery = moment().add('d', numDel + numDelTime);
	
	//If date is a Saturday then add on 2 days to make it a Monday
	if (dtDelivery.day() == 6) {
		dtDelivery = dtDelivery.add('d', 2); }
		
	//If date is a Sunday then add on 1 day to make it a Monday
	if (dtDelivery.day() == 0) {
		dtDelivery = dtDelivery.add('d', 1); }
	
	//document.getElementById('delDate').innerHTML=moment().format("dddd Do MMMM");

	document.getElementById('delDate').innerHTML=dtDelivery.format("dddd Do MMMM");
	document.getElementById('delCost').innerHTML='&pound;'+numSubTotal;
	}

</SCRIPT>
<link href="/css/popcss.css" rel="stylesheet" type="text/css" />
<div id="pop-page">
<img src="/images/poppages/delivery-small.jpg" style="float:right; padding:3px;"/>
<p>We can deliver this <%=strType%> internationally, select your country from the list below to view cost and time estimates:</p>
<p class="lg"><span class="blue">Your Country:</span> <% response.write(strOutput) %></p>
<p class="lg"><span class="blue">Delivery Cost:</span> <strong><span id="delCost"></span></strong></p>
<p class="lg"><span class="blue">Estimated Delivery Date:</span> <strong><span id="delDate"></span></strong></p>
<h3>Delivery Notes</h3>
<ul>
<li>All deliveries are fully insured</li>
<li>Delivery dates are an estimate only, tracking codes will be emailed across once your order has been dispatched</li>
<li>EU orders for customers with a valid EU VAT registration can be sold without UK VAT applied, enter your VAT number at checkout to remove the charge</li>
<li>For orders billed and delivered to a non EU country no VAT charge is required, this will be automatically removed on checkout</li>
<li>We can only accept 3D Secured Debit / Credit Cards  or BACs / Wire Transfer payments for international orders</li>
<li>Please also see our dedicated <a href="/pages/international/" target="_parent">International Orders Terms &amp; Conditions page</a></li>
</ul>
</div>