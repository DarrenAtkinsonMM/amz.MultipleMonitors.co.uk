<% Response.CacheControl = "no-cache" %>
<% Response.Expires = -1 %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="View/Edit Tax Settings" %>
<% Section="layout" %>
<%PmAdmin="1*6*"%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<script type=text/javascript>
function newWindow2(file,window) {
catWindow=open(file,window,'resizable=no,width=500,height=600,scrollbars=1');
if (catWindow.opener == null) catWindow.opener = self;
}
</script>
<form action="AdminTaxSettings.asp" class="pcForms">
<table class="pcCPcontent">	
	<tr>
		<td colspan="2">ProductCart can calculate taxes in three ways: using a tax file (database), using rates that you manually enter, or assuming that a Value Added Tax is included in the prices (VAT). In all cases, make sure to consult your local tax authority for information about the tax laws that you need to adhere to. Here is a summary of your current settings.</td>
	</tr>
	<tr> 
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<%
	IF ptaxfile=1 THEN ' Store is using a tax file
	 %>
		<tr> 
			<td colspan="2">
			<% if request.QueryString("nofile")="0" then %>
			<div class="pcCPmessageSuccess">You are currently using a tax data file.</div>
			<% elseif request.QueryString("nofile")="1" then %>
			<div class="pcCPmessage">The system was not able to locate the tax file that you specified in your 'tax' folder. Please check that you have uploaded the file and that you have typed in the file name correctly, including the file extension. <a href="#" onClick="window.open('taxuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')"><strong>Upload the file now</strong></a>. For more information about obtaining a <u>properly formatted</u> tax data file, <a href="http://www.productcart.com/support/updates/taxes.asp" target="_blank">click here</a>.</div>
			<% end if %>
            </td>
		</tr>
        <tr>
        	<td width="20%" nowrap>Tax file name:</td>
			<td><strong><%=ptaxfilename%></strong></td>
        </tr>
        <tr>
        	<td>Tax Wholesale Customers:</td>
            <td>
			<% If ptaxwholesale="1" then
				response.write "Yes"
			else
				response.write "No"
			end if %>
            </td>
		<tr>
        	<td nowrap valign="top">Fallback States Tax Rates</td>
            <td>
            	<table class="pcCPcontent">
                  <tr bgcolor="#FFFF99"> 
                    <td>State</td>
                    <td>Tax Rate</td>
                    <td><div align="center">Tax Shipping</div></td>
                    <td><div align="center">Tax Shipping and Handling Together</div></td>
                    <td>&nbsp;</td>
                  </tr>
                  <% stateArray=split(ptaxRateState,", ")
                                rateArray=split(ptaxRateDefault,", ")
                                if ptaxSNH<>"" then
                                    taxSNHArray=split(ptaxSNH,", ")
                                end if
                                if ubound(stateArray)=0 then %>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                    <td width="9%"><%=stateArray(0)%> <input type="hidden" name="taxRateState" value="<%=stateArray(0)%>"> 
                    </td>
                    <td><%=rateArray(0)%>%</td>
                    <% if ptaxSNH<>"" then
                                        select case taxSNHArray(0)
                                        case "YY"
                                            taxShippingAlone=""
                                            taxShippingAndHandlingTogether="Yes"
                                        case "YN"
                                            taxShippingAlone="Yes"
                                            taxShippingAndHandlingTogether=""
                                        case "NN"
                                            taxShippingAlone=""
                                            taxShippingAndHandlingTogether=""
                                        end select
                                    else
                                        taxShippingAlone=""
                                        taxShippingAndHandlingTogether=""
                                    end if %>
                    <td><div align="center"><%=taxShippingAlone%></div></td>
                    <td><div align="center"><%=taxShippingAndHandlingTogether%></div></td>
                    <td>&nbsp;</td>
                  </tr>
                  <% else
                                for i=0 to ubound(stateArray)-1 %>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                    <td><%=stateArray(i)%> <input type="hidden" name="taxRateState" value="<%=stateArray(i)%>"> 
                    </td>
                    <td><%=rateArray(i)%>%</td>
                    <%if ptaxSNH<>"" then
                                        select case taxSNHArray(i)
                                            case "YY"
                                            taxShippingAlone=""
                                            taxShippingAndHandlingTogether="Yes"
                                        case "YN"
                                            taxShippingAlone="Yes"
                                            taxShippingAndHandlingTogether=""
                                        case "NN"
                                            taxShippingAlone=""
                                            taxShippingAndHandlingTogether=""
                                        end select
                                    else
                                        taxShippingAlone=""
                                        taxShippingAndHandlingTogether=""
                                    end if %>
                    <td><div align="center"><%=taxShippingAlone%></div></td>
                    <td><div align="center"><%=taxShippingAndHandlingTogether%></div></td>
                    <td>&nbsp;</td>
                  </tr>
                  <% next
                    end if %>
                </table>
            </td>
        </tr>
			<tr> 
				<td class="pcCPspacer" colspan="2"></td>
			</tr>
        <tr>
        	<td colspan="2">
                <div style="margin-bottom: 20px;">
			        <input type="button" class="btn btn-default"  value="Edit Settings" onClick="location.href='AdminTaxSettings_file.asp'" class="btn btn-primary">
                    &nbsp;
                    <input type="button" class="btn btn-default"  onClick="location.href='manageTaxEpt.asp'" value="Set tax exemptions for US states">
                </div>
            </td>
		</tr>
	<% 
	END IF
	' End store using a tax file 
	%>
  
  <% if ptaxAvalara = 1 then %>
  <tr> 
    <td colspan="2">
	<%
        strURL = ptaxAvalaraURL & "/1.0/tax/47.627935,-122.51702/get?saleamount=10"
        
        Set srvAvalaraXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP" & scXML)
        srvAvalaraXmlHttp.open "GET", strURL, False
        srvAvalaraXmlHttp.SetRequestHeader "Content-Type", "text/xml"
        srvAvalaraXmlHttp.SetRequestHeader "Authorization", "Basic " & Base64_Encode(ptaxAvalaraAccount & ":" & ptaxAvalaraLicense)
        srvAvalaraXmlHttp.Send
        
        xmlResponse = srvAvalaraXmlHttp.responseText
        
        Set xmlDoc = server.CreateObject("Msxml2.DOMDocument")
        if xmlDoc.loadXML(xmlResponse) then
            Set result = xmlDoc.selectSingleNode("GeoTaxResult/ResultCode")
        end if
        
        if lcase(result.text) = "success" then
    %>
    	<div class="pcCPmessageSuccess">Avalara connection succeeded.</div>
    <% else %>
 		<div class="pcCPmessage">Avalara connection failed. Please check your information and try again.</div>
    <% end if %>
    	<div class="pcCPmessageSuccess">You are currently using Avalara tax system.</div>
  	</td>
  </tr>
  <tr>
    <td width="25%">Account Number :</td>
		<td><strong><%=ptaxAvalaraAccount%></strong></td>
  </tr>
  <tr>
    <td>License Key :</td>
    <td><strong><%=ptaxAvalaraLicense%></strong></td>
  </tr>
  <tr>
    <td>Company Code :</td>
    <td><strong><%=ptaxAvalaraCode%></strong></td>
  </tr>
  <tr>
    <td>Global Product Code :</td>
    <td><strong><%=ptaxAvalaraProductCode%></strong></td>
  </tr>
  <tr>
    <td>Global Shipping Code :</td>
    <td><strong><%=ptaxAvalaraShippingCode%></strong></td>
  </tr>
  <tr>
    <td>Global Handling Fee Code :</td>
    <td><strong><%=ptaxAvalaraHandlingCode%></strong></td>
  </tr>
  <tr>
    <td>WebService URL :</td>
    <td><strong><%=ptaxAvalaraURL%></strong></td>
  </tr>
  <tr> 
    <td class="pcCPspacer" colspan="2"></td>
  </tr>
  <tr>
    <td>Enable Avalara:</td>
      <td><% If ptaxAvalaraEnabled=1 then response.write "Yes" else response.write "No" end if %></td>
  </tr>
  <tr>
    <td>Enable Logging:</td>
      <td><% If ptaxAvalaraLog=1 then response.write "Yes" else response.write "No" end if %> &nbsp;-&nbsp; <span class="pcSmallText"><i>(includes/logs/avalara.log)</i></span></td>
  </tr>
  <tr>
    <td>Enable Address Validation:</td>
      <td><% If ptaxAvalaraAddressValidation=1 then response.write "Yes" else response.write "No" end if %></td>
  </tr>
  <tr>
    <td>Enable Document Committing:</td>
      <td><% If ptaxAvalaraCommit=1 then response.write "Yes" else response.write "No" end if %></td>
  </tr>
  <tr>
    <td>Tax Wholesale Customers?</td>
    <td>
    <%
		Select Case ptaxAvalaraReason
			case "A" reason = "Federal Government"
			case "B" reason = "State Government"
			case "C" reason = "Tribe / Status Indian / Indian Band"
			case "D" reason = "Foreign Diplomat"
			case "E" reason = "Charitable or Benevolent Organization"
			case "F" reason = "Religious or Education Organization"
			case "G" reason = "Resale"
			case "H" reason = "Commercial Agricultural Production"
			case "I" reason = "Industrial Production / Manufacturer"
			case "J" reason = "Direct Pay Permit"
			case "K" reason = "Direct Mail"
			case "L" reason = "Other"
			case "N" reason = "Local Government"
			case "M" reason = "Not Used"
		End Select
			
		If ptaxwholesale="1" then
			response.write "Yes &nbsp;-&nbsp; <strong>Reason</strong> : " & reason
		else
			response.write "No"
		end if
	%>
    </td>
  </tr>
  <tr> 
    <td class="pcCPspacer" colspan="2"></td>
  </tr>
  <tr>
    <td colspan="2">
      <div style="margin-bottom: 20px;">
				<input type="button" class="btn btn-default"  value="Edit Settings" onClick="location.href='AdminTaxSettings_avalara.asp'" class="btn btn-primary">
      </div>
    </td>
  </tr>
  <% end if %>
										
	<% if ptaxfile=0 AND ptaxsetup=1 then 
		if ptaxVAT="1" then
		' Store is using VAT		
		%>
			<tr> 
				<th colspan="2">VAT (Value Added Tax)</th>
			</tr>
			<tr> 
				<td class="pcCPspacer" colspan="2"></td>
			</tr>
			<tr> 
			<td colspan="2">
				<div class="pcCPmessageSuccess">You are currently setup to use the Value Added Tax (prices include taxes).</div>
            </td>
            </tr>
            <tr>
            	<td nowrap width="20%">Default VAT Rate:</td>
                <td><strong><%=ptaxVATrate%></strong></td>
            </tr>
            <tr>
            	<td nowrap>EU Member State:</td>
                <td>
				<%
				ttaxVATRate_State = "Not Selected - Use Default Rate"
				
				query="SELECT pcVATCountries.pcVATCountry_State From pcVATCountries WHERE pcVATCountries.pcVATCountry_Code = '"& ptaxVATRate_Code &"' Order By pcVATCountry_State ASC;"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
				if not rs.eof then
					ttaxVATRate_State=rs("pcVATCountry_State")
				end if
				set rs = nothing
				
				%>			
				<%=ttaxVATRate_State%>
                </td>
            </tr>
            <tr>
            	<td nowrap>Show VAT on product details page:</td>
                <td><% If ptaxdisplayVAT="1" then response.write "Yes" else response.write "No"	end if %></td>
            </tr>
            <tr>
            	<td nowrap>Include shipping charges:</td>
                <td><% If pTaxonCharges="1" then response.write "Yes" else response.write "No" end if %></td>
            </tr>
            <tr>
            	<td nowrap>Include handling fees:</td>
				<td><% If pTaxonFees="1" then response.write "Yes" else response.write "No" end if %></td>
            </tr>
            <tr>
            	<td nowrap>Tax Wholesale Customers:</td>
                <td><% If ptaxwholesale="1" then response.write "Yes" else response.write "No" end if %></td>
            </tr>
			<tr> 
				<td class="pcCPspacer" colspan="2"></td>
			</tr>
			<tr> 
				<td colspan="2"><input type="button" class="btn btn-default"  value="Edit Settings" onClick="location.href='AdminTaxSettings_VAT.asp'" class="btn btn-primary"></td>
			</tr>

		<% elseif ptaxAvalara=0 then
		' End store using VAT
		' The store is using manual tax calculation method: redirect to that page
			call closeDb()
response.redirect "viewTax.asp"
			response.end
		end if
	end if
	
	if ptaxsetup=0 then
		' Tax settings have not been configured yet: redirect to Tax Wizard
		call closeDb()
response.redirect "AdminTaxWizard.asp"
		Response.End()
	end if 
    %>
    <tr> 
        <td colspan="2">
            <hr>
            <div style="margin-top: 10px">
                Select 'Restart Tax Wizard' if you no longer wish to use this tax method, and would like to swith to an alternative tax calculation method.
            </div>
            <div style="margin-top: 10px">			
                <input type="button" class="btn btn-default"  onClick="location.href='AdminTaxWizard.asp'" value=" Restart Tax Wizard ">
                <input type="button" class="btn btn-default"  name="back" value=" Main Menu " onClick="location.href='menu.asp'">                
            </div>
        </td>
    </tr>
    <tr> 
        <td class="pcCPspacer" colspan="2"></td>
    </tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->
