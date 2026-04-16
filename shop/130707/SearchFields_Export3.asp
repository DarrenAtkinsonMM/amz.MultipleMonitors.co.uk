<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Seach Fields - Mapping Complete" %>
<% section = "products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
pcv_strExportType = request("export")
Select Case pcv_strExportType
	Case "f": pcv_strExportFile = "Froogle"
	Case "c": pcv_strExportFile = "Cashback"
End Select

validfields=request.form("validfields")

'/////////////////////////////////////////////////////////////////////
'// START: REMOVE MAPPINGS
'/////////////////////////////////////////////////////////////////////
query="DELETE from pcSearchFields_Mappings WHERE pcSearchFieldsFileID='"& pcv_strExportType &"' "
set rs=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)
'/////////////////////////////////////////////////////////////////////
'// END: REMOVE MAPPINGS
'/////////////////////////////////////////////////////////////////////


'/////////////////////////////////////////////////////////////////////
'// START: ADD MAPPINGS
'/////////////////////////////////////////////////////////////////////
For i=1 to validfields
	if trim(ucase(request("T" & i)))<>"0" then
		pcv_intSearchField = request("T" & i)
		pcv_intSearchFieldName = request("F" & i)
		query="INSERT INTO pcSearchFields_Mappings (idSearchField, pcSearchFieldsColumn, pcSearchFieldsFileID) VALUES ("& pcv_intSearchField &",N'"& pcv_intSearchFieldName &"','"& pcv_strExportType &"')"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
	end if
Next
'/////////////////////////////////////////////////////////////////////
'// END: ADD MAPPINGS
'/////////////////////////////////////////////////////////////////////
%>
<!--#include file="AdminHeader.asp"-->
<form method="post" action="SearchFields_Export3.asp?export=<%=pcv_strExportType%>" class="pcForms">
	<table class="pcCPcontent">
    <tr>
      <td class="pcCPspacer"></td>
    </tr>	
		<tr>
    <td valign="top">
      <div class="pcCPmessageSuccess">We have successfully mapped your custom search fields to export fields.</div>        	
     </td>
    </tr> 
    <tr>
      <td class="pcCPspacer"></td>
    </tr>	                 
    <tr>
        <td>
          <% if pcv_strExportFile = "Froogle" then %>
            <input type="button" class="btn btn-default" name=backstep value="<< Return to Google Shopping data feed" onClick="location='exportFroogle.asp';" class="btn btn-primary">&nbsp; 
          <% else %>
            <input type="button" class="btn btn-default" name=backstep value="<< Return to Bing Shopping data feed" onClick="location='exportCashback.asp';" class="btn btn-primary">&nbsp;
          <% end if %>
            <input type="button" class="btn btn-default" name=backstep value="Manage Search Fields " onClick="location='ManageSearchFields.asp';">          
      </td>
    </tr>
    </table>
</form>
<!--#include file="Adminfooter.asp"-->
