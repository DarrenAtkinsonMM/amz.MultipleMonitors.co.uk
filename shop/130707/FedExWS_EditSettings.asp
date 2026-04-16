<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEx Web Services Shipping Configuration - Edit Settings" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/FedExWSconstants.asp"-->
<!--#include file="AdminHeader.asp"-->
<% 
query="SELECT ShipmentTypes.AccessLicense FROM ShipmentTypes WHERE (((ShipmentTypes.idShipment)=9));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
strAccessLicense=rs("AccessLicense")
if len(strAccessLicense)<1 then
	strAccessLicense="TEST"
end if
set rs=nothing

pcv_strFEDEXWS_DROPOFF_TYPE = FEDEXWS_DROPOFF_TYPE

if pcv_strFEDEXWS_DROPOFF_TYPE = "" then
	pcv_strFEDEXWS_DROPOFF_TYPE = "REGULARPICKUP"
end if

pcv_FEDEXWS_LISTRATE = FEDEXWS_LISTRATE
if pcv_FEDEXWS_LISTRATE = "" then
	pcv_FEDEXWS_LISTRATE = "0"
end if

pcv_FEDEXWS_SATURDAYDELIVERY = FEDEXWS_SATURDAYDELIVERY
if pcv_FEDEXWS_SATURDAYDELIVERY = "" then
	pcv_FEDEXWS_SATURDAYDELIVERY = "0"
end if

pcv_FEDEXWS_SATURDAYPICKUP = FEDEXWS_SATURDAYPICKUP
if pcv_FEDEXWS_SATURDAYPICKUP = "" then
	pcv_FEDEXWS_SATURDAYPICKUP = "0"
end if

if request.form("submit")<>"" then
	Session("ship_FEDEXWS_FEDEX_PACKAGE")=request.form("FEDEXWS_PACKAGE")
	Session("ship_FEDEXWS_DROPOFF_TYPE")=request.form("FEDEXWS_DROPOFF_TYPE")
	Session("ship_FEDEXWS_HEIGHT")=request.form("FEDEXWS_HEIGHT")
	Session("ship_FEDEXWS_WIDTH")=request.form("FEDEXWS_WIDTH")
	Session("ship_FEDEXWS_LENGTH")=request.form("FEDEXWS_LENGTH")
	Session("ship_FEDEXWS_DIM_UNIT")=request.form("FEDEXWS_DIM_UNIT")
	Session("ship_FEDEXWS_ADDDAY")=request.form("FEDEXWS_ADDDAY")
	Session("ship_FEDEXWS_LISTRATE")=request.form("FEDEXWS_LISTRATE")
	Session("ship_FEDEXWS_PREFERRED_CURRENCY")=request.form("FEDEXWS_PREFERRED_CURRENCY")
	Session("ship_FEDEXWS_ONERATE")=request.form("FEDEXWS_ONERATE")
	Session("ship_FEDEXWS_SATURDAYDELIVERY")=request.form("FEDEXWS_SATURDAYDELIVERY")
	Session("ship_FEDEXWS_SATURDAYPICKUP")=request.form("FEDEXWS_SATURDAYPICKUP")
	Session("ship_FEDEXWS_DYNAMICINSUREDVALUE")=request.form("DynamicInsuredValue")
	Session("ship_FEDEXWS_INSUREDVALUE")=request.form("InsuredValue")
	Session("ship_FEDEXWS_CURRENCY")=request.form("Currency")
	Session("ship_FEDEXWS_SMHUBID")=request.form("SMHubID")
	Session("ship_FEDEXWS_SMINDICIATYPE")=request.form("SMIndiciaType")
	response.redirect "../includes/PageCreateFedExWSConstants.asp?refer=viewShippingOptions.asp#FedExWS"
	response.end
else %>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

	<form name="form1" method="post" action="FedExWS_EditSettings.asp" class="pcForms">
		<table class="pcCPcontent">
				<tr>
				  <td colspan="2">
					<h2>Default Package Type</h2>
					Typically orders are shipped in different boxes depending on what customers purchased. Therefore, in most cases, you will select <em>Custom Packaging</em> here, and  specify the most common box size under <em>Default Package Size</em>. <a href="http://wiki.productcart.com/productcart/shipping-federal_express_ws#packaging_type" target="_blank">See the documentation for details</a>. </td>
				</tr>
				<tr>
					<td align="right">
						<input name="FEDEXWS_PACKAGE" type="radio" value="YOUR_PACKAGING" checked>
					</td>
					<td>Custom Packaging</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_10KG_BOX" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_10KG_BOX" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; 10kg Box</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_25KG_BOX" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_25KG_BOX" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; 25kg Box</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_BOX" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_BOX" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; Box</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_ENVELOPE" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_ENVELOPE" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; Envelope</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_SMALL_BOX" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_SMALL_BOX" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; Small Box</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_MEDIUM_BOX" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_MEDIUM_BOX" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; Medium Box</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_LARGE_BOX" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_LARGE_BOX" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; Large Box</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_EXTRA_LARGE_BOX" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_EXTRA_LARGE_BOX" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; Extra Large Box</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_PAK" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_PAK" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; Pak</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_TUBE" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_TUBE" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; Tube</td>
		  </tr>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
					<h2>Default Package Size</h2>
					If you selected <em>Custom Packaging</em> under <em>Default Package Type</em>, enter the most common package size below. This should refer to the size of the box used for the majority of your shipments. <a href="http://wiki.productcart.com/productcart/shipping-federal_express_ws#packaging_type" target="_blank">See the documentation for details.</a></td>
				</tr>
				<tr>
					<td align="right">Height: </td>
				  <td><input name="FEDEXWS_HEIGHT" type="text" id="FEDEXWS_HEIGHT" value="<%=FEDEXWS_HEIGHT%>" size="4" maxlength="4"></td>
		  </tr>
				<tr>
					<td align="right">Width: </td>
				  <td><input name="FEDEXWS_WIDTH" type="text" id="FEDEXWS_WIDTH" value="<%=FEDEXWS_WIDTH%>" size="4" maxlength="4"></td>
		  </tr>
				<tr>
					<td align="right">Length:</td>
					<td> <input name="FEDEXWS_LENGTH" type="text" id="FEDEXWS_LENGTH" value="<%=FEDEXWS_LENGTH%>" size="4" maxlength="4">
				  <span class="pcSmallText">This is the measurement of the longest side</span></td>
		  </tr>
				<tr>
					<td align="right">Measurement Unit:</td>
					<td>
					<% if FEDEXWS_DIM_UNIT="CM" then%>
					<input type="radio" name="FEDEXWS_DIM_UNIT" value="IN">
					Inches
					<input type="radio" name="FEDEXWS_DIM_UNIT" value="CM" checked>
					Centimeters
					<% else %>
					<input type="radio" name="FEDEXWS_DIM_UNIT" value="IN" checked>
					Inches
					<input type="radio" name="FEDEXWS_DIM_UNIT" value="CM">
					Centimeters
					<% end if %>
				  </td>
		  </tr>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<h2>FedEx Declared Value</h2>
					</td>
				</tr>
				<tr>
					<td colspan="2">To use the order total (total of products listed in the shopping cart) as the declared value amount for shipments, choose <strong>Use cart total</strong>. Otherwise you can set a flat rate as the declared value, and that amount will be used for every FedEx rate calculation in the storefront. The default is 100.00 if you have not set either dynamic or flat rates for the declared value. </td>
				</tr>
				<tr>
					<td align="right"><input name="DynamicInsuredValue" type="radio" value="1" <%if FEDEXWS_DYNAMICINSUREDVALUE="1" then%>checked<% end if %>></td>
					<td>Use cart total</td>
				</tr>
				<tr>
					<td align="right"><input type="radio" name="DynamicInsuredValue" value="0" <%if FEDEXWS_DYNAMICINSUREDVALUE="0" then%>checked<% end if %>></td>
					<td>Use flat rate</td>
				</tr>
				<tr>
					<td width="14%" nowrap>Flat Rate Value:</td>
					<td width="86%"><input type="text" name="InsuredValue" value="<%=FEDEXWS_INSUREDVALUE%>" size="15" maxlength="14" /></td>
				</tr>
        <tr>
					<td width="14%" nowrap>Rate Currency:</td>
          <td width="86%">
             <select name="Currency">
                <option value="ANG" <%if FEDEXWS_CURRENCY="ANG" then %>selected<% end if %>>Antilles Guilder</option>
                <option value="ARN" <%if FEDEXWS_CURRENCY="ARN" then %>selected<% end if %>>Argentinian Pesos</option>
                <option value="AWG" <%if FEDEXWS_CURRENCY="AWG" then %>selected<% end if %>>Arubin Florin</option>
                <option value="AUD" <%if FEDEXWS_CURRENCY="AUD" then %>selected<% end if %>>Australian Dollars</option>
                <option value="BSD" <%if FEDEXWS_CURRENCY="BSD" then %>selected<% end if %>>Bahamian Dollars</option>
                <option value="BHD" <%if FEDEXWS_CURRENCY="BHD" then %>selected<% end if %>>Bahraini Dinar</option>
                <option value="BBD" <%if FEDEXWS_CURRENCY="BBD" then %>selected<% end if %>>Barbados Dollars</option>
                <option value="BMD" <%if FEDEXWS_CURRENCY="BMD" then %>selected<% end if %>>Bermudian Dollars</option>
                <option value="BRL" <%if FEDEXWS_CURRENCY="BRL" then %>selected<% end if %>>Brazilian Real</option>
                <option value="BND" <%if FEDEXWS_CURRENCY="BND" then %>selected<% end if %>>Brunei Dollars</option>
                <option value="CAD" <%if FEDEXWS_CURRENCY="CAD" then %>selected<% end if %>>Canadian Dollars</option>
                <option value="CID" <%if FEDEXWS_CURRENCY="CID" then %>selected<% end if %>>Cayman Dollars</option>
                <option value="CHP" <%if FEDEXWS_CURRENCY="CHP" then %>selected<% end if %>>Chilean Pesos</option>
                <option value="CNY" <%if FEDEXWS_CURRENCY="CNY" then %>selected<% end if %>>Chinese Renminbi</option>
                <option value="COP" <%if FEDEXWS_CURRENCY="COP" then %>selected<% end if %>>Colombian Pesos</option>
                <option value="CRC" <%if FEDEXWS_CURRENCY="CRC" then %>selected<% end if %>>Costa Rican Colones</option>
                <option value="CZK" <%if FEDEXWS_CURRENCY="CZK" then %>selected<% end if %>>Czech Koruna</option>
                <option value="DKK" <%if FEDEXWS_CURRENCY="DKK" then %>selected<% end if %>>Danish Krone</option>
                <option value="RDD" <%if FEDEXWS_CURRENCY="RDD" then %>selected<% end if %>>Dominican R. Pesos</option>
                <option value="ECD" <%if FEDEXWS_CURRENCY="ECD" then %>selected<% end if %>>East Caribbean Dollars</option>
                <option value="EGP" <%if FEDEXWS_CURRENCY="EGP" then %>selected<% end if %>>Egyptian Pound</option>
                <option value="EEK" <%if FEDEXWS_CURRENCY="EEK" then %>selected<% end if %>>Estonian Kroon</option>
                <option value="EUR" <%if FEDEXWS_CURRENCY="EUR" then %>selected<% end if %>>Euro</option>
                <option value="HKD" <%if FEDEXWS_CURRENCY="HKD" then %>selected<% end if %>>Hong Kong Dollars</option>
                <option value="HUF" <%if FEDEXWS_CURRENCY="HUF" then %>selected<% end if %>>Hungarian Forint</option>
                <option value="INR" <%if FEDEXWS_CURRENCY="INR" then %>selected<% end if %>>Indian Rupees</option>
                <option value="JAD" <%if FEDEXWS_CURRENCY="JAD" then %>selected<% end if %>>Jamaican Dollars</option>
                <option value="JYE" <%if FEDEXWS_CURRENCY="JYE" then %>selected<% end if %>>Japanese Yen</option>
                <option value="KUD" <%if FEDEXWS_CURRENCY="KUD" then %>selected<% end if %>>Kuwaiti Dinar</option>
                <option value="LVL" <%if FEDEXWS_CURRENCY="LVL" then %>selected<% end if %>>Latvian Lat</option>
                <option value="LTL" <%if FEDEXWS_CURRENCY="LTL" then %>selected<% end if %>>Lithuanian Lita</option>
                <option value="MOP" <%if FEDEXWS_CURRENCY="MOP" then %>selected<% end if %>>Macau Pataca</option>
                <option value="MYR" <%if FEDEXWS_CURRENCY="MYR" then %>selected<% end if %>>Malaysian Ringgit</option>
                <option value="MXN" <%if FEDEXWS_CURRENCY="MXN" then %>selected<% end if %>>New Mexican Pesos</option>
                <option value="NZD" <%if FEDEXWS_CURRENCY="NZD" then %>selected<% end if %>>New Zealand Dollars</option>
                <option value="NOK" <%if FEDEXWS_CURRENCY="NOK" then %>selected<% end if %>>Norwegian Krone</option>
                <option value="PKR" <%if FEDEXWS_CURRENCY="PKR" then %>selected<% end if %>>Pakistan Rupee</option>
                <option value="PHP" <%if FEDEXWS_CURRENCY="PHP" then %>selected<% end if %>>Phillipine Pesos</option>
                <option value="PLN" <%if FEDEXWS_CURRENCY="PLN" then %>selected<% end if %>>Polish Zloty</option>
                <option value="UKL" <%if FEDEXWS_CURRENCY="UKL" then %>selected<% end if %>>Pounds Sterling (UK)</option>
                <option value="GTQ" <%if FEDEXWS_CURRENCY="GTQ" then %>selected<% end if %>>Quetzales</option>
                <option value="WST" <%if FEDEXWS_CURRENCY="WST" then %>selected<% end if %>>Samoa Currency</option>
                <option value="SAR" <%if FEDEXWS_CURRENCY="SAR" then %>selected<% end if %>>Saudi Arabian Riyal</option>
                <option value="SID" <%if FEDEXWS_CURRENCY="SID" then %>selected<% end if %>>Singapore Dollars</option>
                <option value="SBD" <%if FEDEXWS_CURRENCY="SBD" then %>selected<% end if %>>Solomon Islands Dollars</option>
                <option value="WON" <%if FEDEXWS_CURRENCY="WON" then %>selected<% end if %>>South-Korean Won</option>
                <option value="SEK" <%if FEDEXWS_CURRENCY="SEK" then %>selected<% end if %>>Swedish Krona</option>
                <option value="SFR" <%if FEDEXWS_CURRENCY="SFR" then %>selected<% end if %>>Swiss Francs</option>
                <option value="NTD" <%if FEDEXWS_CURRENCY="NTD" then %>selected<% end if %>>Taiwan Dollars</option>
                <option value="THB" <%if FEDEXWS_CURRENCY="THB" then %>selected<% end if %>>Thai Baht</option>
                <option value="TTD" <%if FEDEXWS_CURRENCY="TTD" then %>selected<% end if %>>Trinidad/Tobago Dollars</option>
                <option value="TRY" <%if FEDEXWS_CURRENCY="TRY" then %>selected<% end if %>>Turkey Lire</option>
                <option value="DHS" <%if FEDEXWS_CURRENCY="DHS" then %>selected<% end if %>>UAE Dirham</option>
                <option value="USD" <%if FEDEXWS_CURRENCY="USD" Or FEDEXWS_CURRENCY="" then %>selected<% end if %>>US Dollars</option>
                <option value="VEF" <%if FEDEXWS_CURRENCY="VEF" then %>selected<% end if %>>Venezuelan Bolivars Fuertes</option>
             </select>
          </td>
        </tr>
				<tr>
					<td>&nbsp;</td>
					 <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<h2>FedEx SmartPost<sup>&reg;</sup></h2>
						To use FedEx SmartPost you must have it enabled on your account. When you have FedEx SmartPost enabled you will be supplied a HUB ID that should be entered here. If you are not using it you can leave this field empty.<br>
					  <br>
						<font color="#FF0000">*To enable FedEx SmartPost for your FedEx account, please contact your FedEx account representative.</font>
					</td>
				</tr>
				<tr>
					<td width="14%"><p>Hub ID: </p></td>
					<td width="86%"><p>
					<select name="SMHubID">
						<option value="" selected>Select HUB ID</option>
						<option value="5015" <%if FEDEXWS_SMHUBID="5015" then%>selected<%end if%>>5015</option>Northborough, MA</option>
						<option value="5087" <%if FEDEXWS_SMHUBID="5087" then%>selected<%end if%>>5087</option>Edison, NJ</option>
						<option value="5150" <%if FEDEXWS_SMHUBID="5150" then%>selected<%end if%>>5150</option>Pittsburgh, PA</option>
						<option value="5185" <%if FEDEXWS_SMHUBID="5185" then%>selected<%end if%>>5185</option>Allentown, PA</option>
						<option value="5254" <%if FEDEXWS_SMHUBID="5254" then%>selected<%end if%>>5254</option>Martinsburg, WV</option>
						<option value="5281" <%if FEDEXWS_SMHUBID="5281" then%>selected<%end if%>>5281</option>Charlotte, NC</option>
						<option value="5303" <%if FEDEXWS_SMHUBID="5303" then%>selected<%end if%>>5303</option>Atlanta, GA</option>
						<option value="5327" <%if FEDEXWS_SMHUBID="5327" then%>selected<%end if%>>5327</option>Orlando, FL</option>
						<option value="5379" <%if FEDEXWS_SMHUBID="5379" then%>selected<%end if%>>5379</option>Memphis, TN</option>
						<option value="5431" <%if FEDEXWS_SMHUBID="5431" then%>selected<%end if%>>5431</option>Grove City, OH</option>
						<option value="5465" <%if FEDEXWS_SMHUBID="5465" then%>selected<%end if%>>5465</option>Indianapolis, IN</option>
						<option value="5481" <%if FEDEXWS_SMHUBID="5481" then%>selected<%end if%>>5481</option>Detroit, MI</option>
						<option value="5531" <%if FEDEXWS_SMHUBID="5531" then%>selected<%end if%>>5531</option>New Berlin, WI</option>
						<option value="5552" <%if FEDEXWS_SMHUBID="5552" then%>selected<%end if%>>5552</option>Minneapolis, MN</option>
						<option value="5631" <%if FEDEXWS_SMHUBID="5631" then%>selected<%end if%>>5631</option>St. Louis, MO</option>
						<option value="5648" <%if FEDEXWS_SMHUBID="5648" then%>selected<%end if%>>5648</option>Kansas, KS</option>
						<option value="5751" <%if FEDEXWS_SMHUBID="5751" then%>selected<%end if%>>5751</option>Dallas, TX</option>
						<option value="5771" <%if FEDEXWS_SMHUBID="5771" then%>selected<%end if%>>5771</option>Houston, TX</option>
						<option value="5802" <%if FEDEXWS_SMHUBID="5802" then%>selected<%end if%>>5802</option>Denver, CO</option>
						<option value="5843" <%if FEDEXWS_SMHUBID="5843" then%>selected<%end if%>>5843</option>Salt Lake City, UT</option>
						<option value="5854" <%if FEDEXWS_SMHUBID="5854" then%>selected<%end if%>>5854</option>Phoenix, AZ</option>
						<option value="5902" <%if FEDEXWS_SMHUBID="5902" then%>selected<%end if%>>5902</option>Los Angeles, CA</option>
						<option value="5929" <%if FEDEXWS_SMHUBID="5929" then%>selected<%end if%>>5929</option>Chino, CA</option>
						<option value="5958" <%if FEDEXWS_SMHUBID="5958" then%>selected<%end if%>>5958</option>Sacramento, CA</option>
						<option value="5983" <%if FEDEXWS_SMHUBID="5983" then%>selected<%end if%>>5983</option>Seattle, WA</option>
					</select>
					</p></td>
				</tr>
				<tr>
					<td><p>Indicia Type: </p></td>
          <td>
          	<p>
              <select name="SMIndiciaType">
                <option value="PARCEL_SELECT" selected>FedEx SmartPost Parcel Select (default)</option>
                <option value="PRESORTED_STANDARD" <% if FEDEXWS_SMINDICIATYPE = "PRESORTED_STANDARD" then response.write "selected" %>>FedEx SmartPost Parcel Select Lightweight</option>
                <option value="PRESORTED_BOUND_PRINTED_MATTER" <% if FEDEXWS_SMINDICIATYPE = "PRESORTED_BOUND_PRINTED_MATTER" then response.write "selected" %>>FedEx SmartPost&reg; Bound Printed Matter</option>
                <option value="MEDIA_MAIL" <% if FEDEXWS_SMINDICIATYPE = "MEDIA_MAIL" then response.write "selected" %>>FedEx SmartPost&reg; Media</option>
              </select>
            </p>
          </td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2"><h2>Default Pickup Method</h2></td>
				</tr>
				<tr>
					<td align="right"><input name="FEDEXWS_DROPOFF_TYPE" type="radio" value="REGULAR_PICKUP" <%if pcv_strFEDEXWS_DROPOFF_TYPE="REGULAR_PICKUP" then%>checked<%end if%>>                </td>
					<td>Regular Pick-up </td>
				</tr>
				<tr>
	<td align="right"><input name="FEDEXWS_DROPOFF_TYPE" type="radio" value="REQUEST_COURIER" <%if pcv_strFEDEXWS_DROPOFF_TYPE="REQUEST_COURIER" then%>checked<%end if%>>                </td>
				  <td>Request Courier  </td>
		  </tr>
				<tr>
	<td align="right"><input name="FEDEXWS_DROPOFF_TYPE" type="radio" value="DROP_BOX" <%if pcv_strFEDEXWS_DROPOFF_TYPE="DROP_BOX" then%>checked<%end if%>>                </td>
				  <td>Dropbox </td>
		  </tr>
				<tr>
	<td align="right"><input name="FEDEXWS_DROPOFF_TYPE" type="radio" value="BUSINESS_SERVICE_CENTER" <%if pcv_strFEDEXWS_DROPOFF_TYPE="BUSINESS_SERVICE_CENTER" then%>checked<%end if%>>                </td>
				  <td>Business Service Center  </td>
		  </tr>
				<tr>
	<td align="right"><input name="FEDEXWS_DROPOFF_TYPE" type="radio" value="STATION" <%if pcv_strFEDEXWS_DROPOFF_TYPE="STATION" then%>checked<%end if%>>                </td>
				  <td>Station </td>
		  </tr>
				<tr>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
					<h2>Rate Type</h2>
					Select which shipping rates you want to be shown in the storefront. NOTE: account and list rates will not always be different amounts. Such behavior does not indicate this feature isn't working.
					<a href="http://wiki.productcart.com/productcart/shipping-federal_express_ws#rate_type" target="_blank">See the documentation for details.</a></td>
				</tr>
				<tr>
					<td align="right" valign="top">
					<input name="FEDEXWS_LISTRATE" type="radio" value="0" <%if pcv_FEDEXWS_LISTRATE="0" then%>checked<%end if%> onclick="$pc('#preferredCurrency').hide();">
					</td>
					<td>Show Account Rates (Default)</td>
				</tr>
				<tr>
					<td align="right" valign="top">
					<input name="FEDEXWS_LISTRATE" type="radio" value="-1" <%if pcv_FEDEXWS_LISTRATE="-1" then%>checked<%end if%> onclick="$pc('#preferredCurrency').hide();">
					</td>
					<td>Show List Rates</td>
				</tr>
				<tr>
					<td align="right" valign="top">
					<input name="FEDEXWS_LISTRATE" type="radio" value="-2" <%if pcv_FEDEXWS_LISTRATE="-2" then%>checked<%end if%> onclick="$pc('#preferredCurrency').show();">
					</td>
					<td>
						Show Preferred Currency Rates
					</td>
				</tr>
				<tbody id="preferredCurrency" style="<% if pcv_FEDEXWS_LISTRATE<>"-2" then response.write "display: none" %>">
					<tr>
						<td></td>
						<td>
							Select Currency: &nbsp;
							<select name="FEDEXWS_PREFERRED_CURRENCY">
                <option value="ANG" <%if FEDEXWS_PREFERRED_CURRENCY="ANG" then %>selected<% end if %>>Antilles Guilder</option>
                <option value="ARN" <%if FEDEXWS_PREFERRED_CURRENCY="ARN" then %>selected<% end if %>>Argentinian Pesos</option>
                <option value="AWG" <%if FEDEXWS_PREFERRED_CURRENCY="AWG" then %>selected<% end if %>>Arubin Florin</option>
                <option value="AUD" <%if FEDEXWS_PREFERRED_CURRENCY="AUD" then %>selected<% end if %>>Australian Dollars</option>
                <option value="BSD" <%if FEDEXWS_PREFERRED_CURRENCY="BSD" then %>selected<% end if %>>Bahamian Dollars</option>
                <option value="BHD" <%if FEDEXWS_PREFERRED_CURRENCY="BHD" then %>selected<% end if %>>Bahraini Dinar</option>
                <option value="BBD" <%if FEDEXWS_PREFERRED_CURRENCY="BBD" then %>selected<% end if %>>Barbados Dollars</option>
                <option value="BMD" <%if FEDEXWS_PREFERRED_CURRENCY="BMD" then %>selected<% end if %>>Bermudian Dollars</option>
                <option value="BRL" <%if FEDEXWS_PREFERRED_CURRENCY="BRL" then %>selected<% end if %>>Brazilian Real</option>
                <option value="BND" <%if FEDEXWS_PREFERRED_CURRENCY="BND" then %>selected<% end if %>>Brunei Dollars</option>
                <option value="CAD" <%if FEDEXWS_PREFERRED_CURRENCY="CAD" then %>selected<% end if %>>Canadian Dollars</option>
                <option value="CID" <%if FEDEXWS_PREFERRED_CURRENCY="CID" then %>selected<% end if %>>Cayman Dollars</option>
                <option value="CHP" <%if FEDEXWS_PREFERRED_CURRENCY="CHP" then %>selected<% end if %>>Chilean Pesos</option>
                <option value="CNY" <%if FEDEXWS_PREFERRED_CURRENCY="CNY" then %>selected<% end if %>>Chinese Renminbi</option>
                <option value="COP" <%if FEDEXWS_PREFERRED_CURRENCY="COP" then %>selected<% end if %>>Colombian Pesos</option>
                <option value="CRC" <%if FEDEXWS_PREFERRED_CURRENCY="CRC" then %>selected<% end if %>>Costa Rican Colones</option>
                <option value="CZK" <%if FEDEXWS_PREFERRED_CURRENCY="CZK" then %>selected<% end if %>>Czech Koruna</option>
                <option value="DKK" <%if FEDEXWS_PREFERRED_CURRENCY="DKK" then %>selected<% end if %>>Danish Krone</option>
                <option value="RDD" <%if FEDEXWS_PREFERRED_CURRENCY="RDD" then %>selected<% end if %>>Dominican R. Pesos</option>
                <option value="ECD" <%if FEDEXWS_PREFERRED_CURRENCY="ECD" then %>selected<% end if %>>East Caribbean Dollars</option>
                <option value="EGP" <%if FEDEXWS_PREFERRED_CURRENCY="EGP" then %>selected<% end if %>>Egyptian Pound</option>
                <option value="EEK" <%if FEDEXWS_PREFERRED_CURRENCY="EEK" then %>selected<% end if %>>Estonian Kroon</option>
                <option value="EUR" <%if FEDEXWS_PREFERRED_CURRENCY="EUR" then %>selected<% end if %>>Euro</option>
                <option value="HKD" <%if FEDEXWS_PREFERRED_CURRENCY="HKD" then %>selected<% end if %>>Hong Kong Dollars</option>
                <option value="HUF" <%if FEDEXWS_PREFERRED_CURRENCY="HUF" then %>selected<% end if %>>Hungarian Forint</option>
                <option value="INR" <%if FEDEXWS_PREFERRED_CURRENCY="INR" then %>selected<% end if %>>Indian Rupees</option>
                <option value="JAD" <%if FEDEXWS_PREFERRED_CURRENCY="JAD" then %>selected<% end if %>>Jamaican Dollars</option>
                <option value="JYE" <%if FEDEXWS_PREFERRED_CURRENCY="JYE" then %>selected<% end if %>>Japanese Yen</option>
                <option value="KUD" <%if FEDEXWS_PREFERRED_CURRENCY="KUD" then %>selected<% end if %>>Kuwaiti Dinar</option>
                <option value="LVL" <%if FEDEXWS_PREFERRED_CURRENCY="LVL" then %>selected<% end if %>>Latvian Lat</option>
                <option value="LTL" <%if FEDEXWS_PREFERRED_CURRENCY="LTL" then %>selected<% end if %>>Lithuanian Lita</option>
                <option value="MOP" <%if FEDEXWS_PREFERRED_CURRENCY="MOP" then %>selected<% end if %>>Macau Pataca</option>
                <option value="MYR" <%if FEDEXWS_PREFERRED_CURRENCY="MYR" then %>selected<% end if %>>Malaysian Ringgit</option>
                <option value="MXN" <%if FEDEXWS_PREFERRED_CURRENCY="MXN" then %>selected<% end if %>>New Mexican Pesos</option>
                <option value="NZD" <%if FEDEXWS_PREFERRED_CURRENCY="NZD" then %>selected<% end if %>>New Zealand Dollars</option>
                <option value="NOK" <%if FEDEXWS_PREFERRED_CURRENCY="NOK" then %>selected<% end if %>>Norwegian Krone</option>
                <option value="PKR" <%if FEDEXWS_PREFERRED_CURRENCY="PKR" then %>selected<% end if %>>Pakistan Rupee</option>
                <option value="PHP" <%if FEDEXWS_PREFERRED_CURRENCY="PHP" then %>selected<% end if %>>Phillipine Pesos</option>
                <option value="PLN" <%if FEDEXWS_PREFERRED_CURRENCY="PLN" then %>selected<% end if %>>Polish Zloty</option>
                <option value="UKL" <%if FEDEXWS_PREFERRED_CURRENCY="UKL" then %>selected<% end if %>>Pounds Sterling (UK)</option>
                <option value="GTQ" <%if FEDEXWS_PREFERRED_CURRENCY="GTQ" then %>selected<% end if %>>Quetzales</option>
                <option value="WST" <%if FEDEXWS_PREFERRED_CURRENCY="WST" then %>selected<% end if %>>Samoa Currency</option>
                <option value="SAR" <%if FEDEXWS_PREFERRED_CURRENCY="SAR" then %>selected<% end if %>>Saudi Arabian Riyal</option>
                <option value="SID" <%if FEDEXWS_PREFERRED_CURRENCY="SID" then %>selected<% end if %>>Singapore Dollars</option>
                <option value="SBD" <%if FEDEXWS_PREFERRED_CURRENCY="SBD" then %>selected<% end if %>>Solomon Islands Dollars</option>
                <option value="WON" <%if FEDEXWS_PREFERRED_CURRENCY="WON" then %>selected<% end if %>>South-Korean Won</option>
                <option value="SEK" <%if FEDEXWS_PREFERRED_CURRENCY="SEK" then %>selected<% end if %>>Swedish Krona</option>
                <option value="SFR" <%if FEDEXWS_PREFERRED_CURRENCY="SFR" then %>selected<% end if %>>Swiss Francs</option>
                <option value="NTD" <%if FEDEXWS_PREFERRED_CURRENCY="NTD" then %>selected<% end if %>>Taiwan Dollars</option>
                <option value="THB" <%if FEDEXWS_PREFERRED_CURRENCY="THB" then %>selected<% end if %>>Thai Baht</option>
                <option value="TTD" <%if FEDEXWS_PREFERRED_CURRENCY="TTD" then %>selected<% end if %>>Trinidad/Tobago Dollars</option>
                <option value="TRY" <%if FEDEXWS_PREFERRED_CURRENCY="TRY" then %>selected<% end if %>>Turkey Lire</option>
                <option value="DHS" <%if FEDEXWS_PREFERRED_CURRENCY="DHS" then %>selected<% end if %>>UAE Dirham</option>
                <option value="USD" <%if FEDEXWS_PREFERRED_CURRENCY="USD" Or FEDEXWS_CURRENCY="" then %>selected<% end if %>>US Dollars</option>
                <option value="VEF" <%if FEDEXWS_PREFERRED_CURRENCY="VEF" then %>selected<% end if %>>Venezuelan Bolivars Fuertes</option>
							</select>
						</td>
					</tr>
				</tbody>

				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<h2>FedEx One Rate<sup>&reg;</sup></h2>
						FedEx One Rate<sup>&reg;</sup> is a flat rate shipping option now offered by FedEx that is not dependent on package weight or dimensions for all shipments under 50 lbs. <strong>Note:</strong> When this feature is "On" flat rate pricing is displayed only when available. 
					</td>
				</tr>
			
				<tr>
					<td align="right">
					<input name="FEDEXWS_ONERATE" type="radio" value="0" <%if FEDEXWS_ONERATE="0" then%>checked<%end if%>>
					</td>
					<td>Off</td>
				  </tr>
				<tr>
					<td align="right">
					<input name="FEDEXWS_ONERATE" type="radio" value="-1" <%if FEDEXWS_ONERATE="-1" then%>checked<%end if%>>
					</td>
					<td>On</td>
				  </tr>
				<tr>
				<tr>
					<td>&nbsp;</td>
					 <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<h2>Saturday Delivery</h2>
						<strong>Note:</strong> When this feature is "On" Saturday Delivery options are displayed only when available.
					</td>
				</tr>
				<tr>
					<td align="right">
					<input name="FEDEXWS_SATURDAYDELIVERY" type="radio" value="0" <%if pcv_FEDEXWS_SATURDAYDELIVERY="0" then%>checked<%end if%>>
					</td>
					<td>Off (recommended)</td>
				  </tr>
				<tr>
					<td align="right">
					<input name="FEDEXWS_SATURDAYDELIVERY" type="radio" value="-1" <%if pcv_FEDEXWS_SATURDAYDELIVERY="-1" then%>checked<%end if%>>
					</td>
					<td>On</td>
				  </tr>
				<tr>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<h2>Saturday Pickup</h2>
						<strong>Note:</strong> When this feature is "On" Saturday Pickup pricing is displayed  when available.
					If you never will ship on a Saturday, you will turn this feature off.</td>
				</tr>
				<tr>
					<td align="right">
					<input name="FEDEXWS_SATURDAYPICKUP" type="radio" value="0" <%if pcv_FEDEXWS_SATURDAYPICKUP="0" then%>checked<%end if%>>
					</td>
					<td>Off (recommended)</td>
				  </tr>
				<tr>
					<td align="right">
					<input name="FEDEXWS_SATURDAYPICKUP" type="radio" value="-1" <%if pcv_FEDEXWS_SATURDAYPICKUP="-1" then%>checked<%end if%>>
					</td>
					<td>On</td>
				  </tr>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<h2>Delay Shipment Setting</h2>
							<strong>Note:</strong> The system will add a lead time to the expected shipping date for the rate requests in the store front. Set this to the number of days in which you normally ship out packages after the date they are ordered. 
							<br /><br />
							For example: if you set this to 2 days, and an order was placed on a Monday, the ship date will then be set to Wednesday and will reflect on both the rates in the storefront and in the shipping center when generating shipping labels.<br>
					</td>
				</tr>
				<tr>
					<td align="right">Number of Days:
			  </td>
					<td><input name="FEDEXWS_ADDDAY" type="text" id="FEDEXWS_ADDDAY" value="<%=FEDEXWS_ADDDAY%>" size="2" maxlength="2">
				  </td>
		  </tr>
				<tr>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2" align="center">
					<hr>
					<input type="submit" name="Submit" value="Submit" class="btn btn-primary"></td>
				</tr>
	  </table>
</form>
<% end if %>

<!--#include file="AdminFooter.asp"-->