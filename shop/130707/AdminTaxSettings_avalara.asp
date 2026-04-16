<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="View/Edit Tax Settings - Avalara" %>
<% Section="taxmenu" %>
<%PmAdmin="1*6*"%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script type=text/javascript>"&vbcrlf
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf

StrGenericJSError = dictLanguageCP.Item(Session("language")&"_cpCommon_403")

pcs_JavaTextField	"AvalaraAccount", true, StrGenericJSError, ""
pcs_JavaTextField	"AvalaraLicense", true, StrGenericJSError, ""
pcs_JavaTextField	"AvalaraCode", true, StrGenericJSError, ""

response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf
response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' End Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

%>
<form name="form1" method="post" action="../includes/PageCreateTaxSettings.asp" onSubmit="return Form1_Validator(this);" class="pcForms">
	<input type="hidden" name="taxAvalara" value="1">
	<input type="hidden" name="Page_Name" value="taxsettings.asp">
	<input type="hidden" name="refpage" value="AdminTaxSettings.asp">
    <table class="pcCPcontent">
    	<tr>
    		<td width="25%" align="right">Account Number :</td>
    		<td>
                <input type="text" name="AvalaraAccount" value="<%=ptaxAvalaraAccount%>" />
                <% pcs_RequiredImageTag "AvalaraAccount", true %>
            </td>
    	</tr>
    	<tr>
    		<td align="right">License Key :</td>
    		<td>
                <input type="text" name="AvalaraLicense" value="<%=ptaxAvalaraLicense%>" />
                <% pcs_RequiredImageTag "AvalaraLicense", true %>
            </td>
    	</tr>
    	<tr>
      		<td align="right">Company Code :</td>
      		<td>
                <input type="text" name="AvalaraCode" value="<%=ptaxAvalaraCode%>" />
                <% pcs_RequiredImageTag "AvalaraCode", true %>
            </td>
    	</tr>
        <tr>
      		<td align="right">Global Product Tax Code :</td>
      		<td><input type="text" name="AvalaraProductCode" value="<%=ptaxAvalaraProductCode%>" /></td>
    	</tr>
        <tr>
      		<td align="right">Global Shipping Tax Code :</td>
      		<td><input type="text" name="AvalaraShippingCode" value="<%=ptaxAvalaraShippingCode%>" /></td>
    	</tr>
        <tr>
      		<td align="right">Global Handling Fee Tax Code :</td>
      		<td><input type="text" name="AvalaraHandlingCode" value="<%=ptaxAvalaraHandlingCode%>" /></td>
    	</tr>
    	<tr>
      		<td colspan="2">Is your Avalara AvaTax account currently in development or production mode? Select the WebService URL to use for your Avalara account.</td>
    	</tr>
        <% 
        pcv_boolIsDevelopment = False
        If ptaxAvalaraURL = "https://development.avalara.net" Then 
            pcv_boolIsDevelopment = True
        End If
        %>
	    <tr>
      		<td>&nbsp;</td>
      		<td>
                <input type="radio" name="AvalaraURL" value="https://development.avalara.net" class="clearBorder" <% If pcv_boolIsDevelopment Then %>checked<% End If %> /> <strong>Development</strong> - https://development.avalara.net
      		</td>
    	</tr>
    	<tr>
      		<td>&nbsp;</td>
      		<td><input type="radio" name="AvalaraURL" value="https://avatax.avalara.net" class="clearBorder" <% If Not pcv_boolIsDevelopment Then %>checked<% End If %> /> <strong>Production</strong> - https://avatax.avalara.net</td>
    	</tr>
        <tr>
        	<td align="right">Enable Avalara :</td>
        	<td>
        		<input name="AvalaraEnabled" type="checkbox" value="1" <% If ptaxAvalaraEnabled=1 then%>checked<% end if %>>
        	</td>
    	</tr>
        <tr>
        	<td align="right">Enable Logging :</td>
        	<td>
        		<input name="AvalaraLog" type="checkbox" value="1" <% If ptaxAvalaraLog=1 then%>checked<% end if %>> <span class="pcSmallText"><i>(includes/logs/avalara.log)</i></span>
        	</td>
    	</tr>
        <tr>
        	<td align="right">Enable Address Validation :</td>
        	<td width="80%">
        		<input name="AvalaraAddressValidation" type="checkbox" value="1" <% If ptaxAvalaraAddressValidation=1 then%>checked<% end if %>>
        	</td>
    	</tr>
        <tr>
        	<td align="right">Enable Document Committing :</td>
        	<td>
        		<input name="AvalaraCommit" type="checkbox" value="1" <% If ptaxAvalaraCommit=1 then%>checked<% end if %>>
        	</td>
    	</tr>
	    <tr>
        	<td align="right">Tax Wholesale Customers?</td>
        	<td>
        		
        		<input type="radio" name="taxwholesale" value="1" <% If ptaxwholesale="1" then%>checked<% end if %> class="clearBorder" onClick="document.getElementById('showReason').style.display='none';"> Yes
                
                <input type="radio" name="taxwholesale" value="0" <% If ptaxwholesale<>"1" then%>checked<% end if %> class="clearBorder" onClick="document.getElementById('showReason').style.display='block';"> No
                
        	</td>
    	</tr>
        <tr>
        	<td>&nbsp;</td>
            <td>
            	<div id="showReason" <%if ptaxwholesale <> "0" then%>style="display:none;"<%end if%>>Reason : 
                    <select name="AvalaraReason">
                        <option value="A" <%if ptaxAvalaraReason = "A" then%>selected<%end if%>>Federal Government</option>
                        <option value="B" <%if ptaxAvalaraReason = "B" then%>selected<%end if%>>State Government</option>
                        <option value="C" <%if ptaxAvalaraReason = "C" then%>selected<%end if%>>Tribe / Status Indian / Indian Band</option>
                        <option value="D" <%if ptaxAvalaraReason = "D" then%>selected<%end if%>>Foreign Diplomat</option>
                        <option value="E" <%if ptaxAvalaraReason = "E" then%>selected<%end if%>>Charitable or Benevolent Organization</option>
                        <option value="F" <%if ptaxAvalaraReason = "F" then%>selected<%end if%>>Religious or Education Organization</option>
                        <option value="G" <%if ptaxAvalaraReason = "G" then%>selected<%end if%>>Resale</option>
                        <option value="H" <%if ptaxAvalaraReason = "H" then%>selected<%end if%>>Commercial Agricultural Production</option>
                        <option value="I" <%if ptaxAvalaraReason = "I" then%>selected<%end if%>>Industrial Production / Manufacturer</option>
                        <option value="J" <%if ptaxAvalaraReason = "J" then%>selected<%end if%>>Direct Pay Permit</option>
                        <option value="K" <%if ptaxAvalaraReason = "K" then%>selected<%end if%>>Direct Mail</option>
                        <option value="L" <%if ptaxAvalaraReason = "L" then%>selected<%end if%>>Other</option>
                        <option value="N" <%if ptaxAvalaraReason = "N" then%>selected<%end if%>>Local Government</option>
                        <option value="M" <%if ptaxAvalaraReason = "M" then%>selected<%end if%>>Not Used</option>
                    </select>
                </div>
            </td>
        </tr>
        <!--
        <tr> 
          	<td nowrap="nowrap">&nbsp;</td>
        </tr>
        <tr> 
          	<td colspan="2">Since you can enter more then one tax rule, ProductCart gives the ability to <strong>show different types of taxes separately</strong> to the customer. For example, this is useful for Canadian online stores (more on <a href="http://wiki.productcart.com/productcart/tax_manual#tax_by_zone_for_canadian_online_stores" target="_blank">tax calculation for Canada-based stores</a>).</td>
        </tr>
        <tr> 
          	<td nowrap="nowrap">Display taxes separately?</td>
          	<td>
            	<input type="radio" name="taxseparate" value="0" checked class="clearBorder"> No 
            	<input type="radio" name="taxseparate" value="1" <% If ptaxseparate="1" then%>checked<% end if %> class="clearBorder"> Yes
          	</td>
        </tr>
        -->
        <tr> 
            <td colspan="2">&nbsp;</td>
        </tr>
        <tr> 
          	<td colspan="2" align="center">
            	<input type="submit" name="Submit" value="Update & Continue" class="btn btn-primary">&nbsp;
            	<input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
          	</td>
        </tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->