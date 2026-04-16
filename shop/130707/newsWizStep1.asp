<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Newsletter Wizard: Select Customers" %>
<% section="mngAcc" %>
<%PmAdmin=7%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="newsWizStep1a.asp" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2">
		<table width="100%">
		<tr>
			<td width="5%" align="right"><img border="0" src="images/step1a.gif"></td>
			<td width="95%"><b>Select Customers</b></td>
		</tr>
		<tr>
			<td align="right"><img border="0" src="images/step2.gif"></td>
			<td><font color="#A8A8A8">Verify customers</font></td>
		</tr>
		<tr>
			<td align="right"><img border="0" src="images/step3.gif"></td>
			<td><font color="#A8A8A8">Enter message</font></td>
		</tr>
		<tr>
			<td align="right"><img border="0" src="images/step4.gif"></td>
			<td><font color="#A8A8A8">Test message</font></td>
		</tr>
		<tr>
			<td align="right"><img border="0" src="images/step5.gif"></td>
			<td><font color="#A8A8A8">Send message</font></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Opted in</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">These are customers that chose to receive information from you. Make sure to comply with local regulations to avoid sending messages that might be considered SPAM and could trigger fines (<a href="http://wiki.productcart.com/productcart/customer-newsletters" target="_blank">More information</a>).</td>
</tr>
<tr>
	<td width="20%" align="right">Select one:</td>
	<td width="80%">
		<select name="SOptedIn">
			<option value="1">Opted in only</option>
			<option value="0">All customers</option>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Product purchased</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">You can filter based on who did or did not purchase a specific product, or any product within a specific product category.</td>
</tr>
<tr>
	<td width="20%" align="right" valign="middle"><input type="radio" name="purchaseType" value="0" checked class="clearBorder" onClick="document.getElementById('selectProduct').style.display='none'; document.getElementById('selectCategory').style.display='none';"></td>
	<td width="80%" valign="middle">All customers</td>
</tr>
<tr>
	<td width="20%" align="right" valign="middle"><input type="radio" name="purchaseType" value="4" class="clearBorder" onClick="document.getElementById('selectProduct').style.display='none'; document.getElementById('selectCategory').style.display='none';"></td>
	<td width="80%" valign="middle">Customers who have not yet purchased anything</td>
</tr>
<tr>
	<td width="20%" align="right" valign="middle"><input type="radio" name="purchaseType" value="3" class="clearBorder" onClick="document.getElementById('selectProduct').style.display='none'; document.getElementById('selectCategory').style.display='none';"></td>
	<td width="80%" valign="middle">Customers who have purchased something (<em>regardless of what they purchased</em>)</td>
</tr>
<tr>
	<td width="20%" align="right" valign="middle"><input type="radio" name="purchaseType" value="1" class="clearBorder" onClick="document.getElementById('selectProduct').style.display=''; document.getElementById('selectCategory').style.display='';"></td>
	<td width="80%" valign="middle">Customers who purchased...</td>
</tr>
<tr>
	<td width="20%" align="right" valign="middle"><input type="radio" name="purchaseType" value="2" class="clearBorder" onClick="document.getElementById('selectProduct').style.display=''; document.getElementById('selectCategory').style.display='';"></td>
	<td width="80%" valign="middle">Customers who did not purchase...</td>
</tr>
<tr id="selectProduct" style="display: none;">
	<td width="20%" align="right" valign="top">Select a product:</td>
	<td width="80%" valign="middle">Narrow your product search by selecting the category first. Then select the Product from the drop-down.
	<div id="CategoryList" style="margin-top: 4px; margin-bottom: 6px;"></div>
	<div id="ProductList"></div>
  </td>
</tr>
<tr id="selectCategory" style="display: none;">
	<td align="right">or a category: </td>
	<td>
		<select name="SIDCategory">
			<option value="0">Any</option>
			<%
			query="SELECT idcategory, idParentCategory, categorydesc FROM categories WHERE idcategory<>1 and iBTOHide<>1 ORDER BY categoryDesc ASC"
			set rstemp=conntemp.execute(query)
			if err.number <> 0 then
				set rstemp=nothing
				
				call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in retreiving categories from database: "&Err.Description) 
			end If
			if rstemp.eof then
				catnum="0"
				rstemp=nothing
			end If
			if catnum<>"0" then
				pcArr=rstemp.getRows()
				set rstemp=nothing
				intCount=ubound(pcArr,2)
				For i=0 to intCount
					idparentcategory=pcArr(1,i)
					if idparentcategory="1" then %>
					    <option value="<%response.write pcArr(0,i)%>"><%=pcArr(2,i)%></option>
				    <%else
					For j=0 to intCount
					if Clng(pcArr(0,j))=Clng(idparentcategory) then
					parentDesc=pcArr(2,j)%>
						<option value="<%response.write pcArr(0,i)%>"><%response.write pcArr(2,i)&" ["&parentDesc&"]"%></option>
					<%
					exit for
					end if 
					Next
					end if
				Next
			End if
			%>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Customer type</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">You can include all customers, only retail customers, or only wholesale customers.</td>
</tr>
<tr>
	<td width="20%" align="right">Select customer type:</td>
	<td width="80%">
		<select name="SCustType">
			<option value="0">Any</option>
			<option value="1">Only Retail Customers</option>
			<option value="2">Only Whoselale Customers</option>
			<% 'START CT ADD %>
					<% 'if there are PBP customer type categories - List them here 
					
					query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories;"
					SET rs=Server.CreateObject("ADODB.RecordSet")
					SET rs=conntemp.execute(query)
					if NOT rs.eof then 
						do until rs.eof 
							intIdcustomerCategory=rs("idcustomerCategory")
							strpcCC_Name=rs("pcCC_Name")
							%>
							<option value='CC_<%=intIdcustomerCategory%>'
							<%if Session("pcAdmincustomertype")="CC_"&intIdcustomerCategory then 
								response.write "selected"
							end if%>
							><%="Only " & strpcCC_Name%></option>
							<% rs.moveNext
						loop
					end if
					SET rs=nothing
					
					
					
			'END CT ADD %>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Date range</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">You can include only customers that have made a purchase within a certain date range.</td>
</tr>
<tr>
	<td align="right">Start date:</td>
	<td><input type="text" class="datepicker" name="SStartDate" size="20"> (mm/dd/yyyy)</td>
</tr>
<tr>
	<td align="right">End date:</td>
	<td><input type="text" class="datepicker" name="SEndDate" size="20"> (mm/dd/yyyy)</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td colspan="2" align="center">
		<input type="submit" name="submit" value="Continue" class="btn btn-primary">
		&nbsp;<input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
	</td>
</tr>
</table>
</form>
<script type=text/javascript>
	function LoadProductList(IDCat)
	{
		$pc("#ProductList").html("<img src=\"images/pc_AjaxLoader.gif\" border=0 align=\"texttop\"> Loading products...");
		$pc.ajax({
			type: "GET",
			data: "x_idCategory="+IDCat,
			url: "categories_productsxml.asp",
			timeout: 45000,
		}).done(function ( data ) {
		if(data=="NONE") {
			$pc("#ProductList").html("<b>No products</b>");
		}
		else
		{
			$pc("#ProductList").html("<select id=\"SIDproduct\" name=\"SIDproduct\">"+data+"</select>");
		}
		}).fail(function() { alert("error"); });
	}
	
	function LoadCategoryList()
	{
		$pc("#CategoryList").html("<img src=\"images/pc_AjaxLoader.gif\" border=0 align=\"texttop\"> Loading categories...");
		$pc.ajax({
			type: "GET",
			data: "CP=1&idRootCategory=<%=pcv_IdRootCategory%>",
			url: "pcRequestCategories.asp",
			timeout: 45000,
		}).done(function ( data ) {
		if(data=="NONE") {
			$pc("#CategoryList").html("<b>No categories</b>");
			$pc("#ProductList").html("<b>No products</b>");
		}
		else
		{
			$pc("#CategoryList").html("<select id=\"categorySelect\" name=\"categorySelect\" onchange=\"javascript:LoadProductList(this.value)\">"+data+"</select>");
			LoadProductList(document.getElementById("categorySelect").value);
		}
		}).fail(function() { alert("error"); });
	}
	$pc(document).ready(function()
	{
		LoadCategoryList();
	});
</script>
<!--#include file="AdminFooter.asp"-->
