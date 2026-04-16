<% pageTitle = "Reverse Import Wizard - Step 3: Choose the Fields to be Exported" %>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
if session("cp_revImport_prdlist")="" then
	call closeDb()
response.redirect "ReverseImport_step1.asp"
end if
%>
<!--#include file="AdminHeader.asp"-->
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<FORM name="checkboxform" method="post" action="ReverseImport_step3.asp" class="pcForms">
<table class="pcCPcontent">
	<tr><td align="right"><input type="checkbox" name="C58" value="1" checked class="clearBorder"></td><td>Product ID#</td></tr>
	<tr><td align="right"><input type="checkbox" name="C1" value="1" checked class="clearBorder"></td><td>SKU</td></tr>
	<tr><td align="right"><input type="checkbox" name="C2" value="1" checked class="clearBorder"></td><td>Name</td></tr>

	<%if session("cp_revImport_extype")="1" or session("cp_revImport_extype")="2" then%>
		<tr><td align="right"><input type="checkbox" name="C48" value="1" checked class="clearBorder"></td><td>Parent Product ID#</td></tr>
	<%end if%>

	<%if session("cp_revImport_extype")="0" or session("cp_revImport_extype")="2" then%>
		<tr><td align="right"><input type="checkbox" name="C3" value="1" class="clearBorder"></td><td>Description</td></tr>
		<tr><td align="right"><input type="checkbox" name="C4" value="1" class="clearBorder"></td><td>Short Description</td></tr>
		<tr><td align="right"><input type="checkbox" name="C5" value="1" checked class="clearBorder"></td><td>Product Type</td></tr>
		<tr><td align="right"><input type="checkbox" name="C47" value="1" checked class="clearBorder"></td><td>Apparel Product</td></tr>
    <%end if%>

	<tr><td align="right"><input type="checkbox" name="C6" value="1" checked class="clearBorder"></td><td>Online Price</td></tr>

	<%if session("cp_revImport_extype")="0" or session("cp_revImport_extype")="2" then%>
		<tr><td align="right"><input type="checkbox" name="C7" value="1" checked class="clearBorder"></td><td>List Price</td></tr>
    <%end if%>

	<tr><td align="right"><input type="checkbox" name="C8" value="1" checked class="clearBorder"></td><td>Wholesale Price</td></tr>

	<%if session("cp_revImport_extype")="1" or session("cp_revImport_extype")="2" then%>
		<tr><td align="right"><input type="checkbox" name="C6a" value="1" checked class="clearBorder"></td><td>Online Price Difference</td></tr>
		<tr><td align="right"><input type="checkbox" name="C6b" value="1" checked class="clearBorder"></td><td>Wholesale Price Difference</td></tr>
	<%end if%>
	
	<%query="SELECT idCustomerCategory,pcCC_Name FROM pcCustomerCategories ORDER BY idCustomerCategory ASC;"
	set rstemp=conntemp.execute(query)
	
	if not rstemp.eof then
		tmpArr=rstemp.getRows()
		intCount=ubound(tmpArr,2)
		For i=0 to intCount%>
			<tr><td align="right"><input type="checkbox" name="PCat<%=tmpArr(0,i)%>" value="1" checked class="clearBorder"></td><td>Pricing Category: <%=tmpArr(1,i)%></td></tr>
		<%Next
	end if
	set rstemp=nothing
	%>

	<tr><td align="right"><input type="checkbox" name="C9" value="1" checked class="clearBorder"></td><td>Weight</td></tr>
	<tr><td align="right"><input type="checkbox" name="C10" value="1" checked class="clearBorder"></td><td>Stock</td></tr>
   
	<%if session("cp_revImport_extype")="0" or session("cp_revImport_extype")="2" then%>
		<tr><td align="right"><input type="checkbox" name="C11" value="1" class="clearBorder"></td><td>Categories Information</td></tr>
		<tr><td align="right"><input type="checkbox" name="C12" value="1" class="clearBorder"></td><td>Brand Information</td></tr>
    	<tr><td align="right"><input type="checkbox" name="C13" value="1" checked class="clearBorder"></td><td>Thumbnail Image</td></tr>
    <%end if%>

	<tr><td align="right"><input type="checkbox" name="C14" value="1" checked class="clearBorder"></td><td>General Image</td></tr>
	<tr><td align="right"><input type="checkbox" name="C15" value="1" checked class="clearBorder"></td><td>Detail view Image</td></tr>
    <tr><td align="right"><input type="checkbox" name="C68" value="1" checked class="clearBorder"></td><td>Alt Tag Text</td></tr>
	<tr><td align="right"><input type="checkbox" name="C16" value="1" checked class="clearBorder"></td><td>Active</td></tr>
 
	<%if session("cp_revImport_extype")="0" or session("cp_revImport_extype")="2" then%>
		<tr><td align="right"><input type="checkbox" name="C17" value="1" checked class="clearBorder"></td><td>Show savings</td></tr>
		<tr><td align="right"><input type="checkbox" name="C18" value="1" checked class="clearBorder"></td><td>Special</td></tr>
		<tr><td align="right"><input type="checkbox" name="C46" value="1" checked class="clearBorder"></td><td>Featured</td></tr>
		<tr><td align="right"><input type="checkbox" name="C19" value="1" class="clearBorder"></td><td>Product Options Information</td></tr>
		<tr><td align="right"><input type="checkbox" name="C20" value="1" class="clearBorder"></td><td>Reward Points</td></tr>
		<tr><td align="right"><input type="checkbox" name="C21" value="1" checked class="clearBorder"></td><td>Non-taxable</td></tr>
		<tr><td align="right"><input type="checkbox" name="C22" value="1" checked class="clearBorder"></td><td>No shipping charge</td></tr>
		<tr><td align="right"><input type="checkbox" name="C23" value="1" checked class="clearBorder"></td><td>Not for sale</td></tr>
		<tr><td align="right"><input type="checkbox" name="C24" value="1" class="clearBorder"></td><td>Not for sale copy</td></tr>
    <%end if%>

	<tr><td align="right"><input type="checkbox" name="C25" value="1" checked class="clearBorder"></td><td>Disregard stock</td></tr>
  
	<%if session("cp_revImport_extype")="0" or session("cp_revImport_extype")="2" then%>
		<tr><td align="right"><input type="checkbox" name="C26" value="1" checked class="clearBorder"></td><td>Display No Shipping Text</td></tr>
		<tr><td align="right"><input type="checkbox" name="C27" value="1" checked class="clearBorder"></td><td>Minimum Quantity customers can buy</td></tr>
		<tr><td align="right"><input type="checkbox" name="C28" value="1" checked class="clearBorder"></td><td>Force purchase of multiples of minimum</td></tr>
		<tr><td align="right"><input type="checkbox" name="C29" value="1" class="clearBorder"></td><td>Oversized Product Details</td></tr>
		<tr><td align="right"><input type="checkbox" name="C30" value="1" class="clearBorder"></td><td>Product Cost</td></tr>
    <%end if%>

	<tr><td align="right"><input type="checkbox" name="C31" value="1" class="clearBorder"></td><td>Back-Order</td></tr>
	<tr><td align="right"><input type="checkbox" name="C32" value="1" class="clearBorder"></td><td>Ship within N Days</td></tr>
   
	<%if session("cp_revImport_extype")="0" or session("cp_revImport_extype")="2" then%>
	
		<tr><td align="right"><input type="checkbox" name="C33" value="1" checked class="clearBorder"></td><td>Low inventory notification</td></tr>
		<tr><td align="right"><input type="checkbox" name="C34" value="1" checked class="clearBorder"></td><td>Reorder Level</td></tr>
		<tr><td align="right"><input type="checkbox" name="C35" value="1" class="clearBorder"></td><td>Is Drop-shipped</td></tr>
		<tr><td align="right"><input type="checkbox" name="C36" value="1" class="clearBorder"></td><td>Supplier ID</td></tr>
		<tr><td align="right"><input type="checkbox" name="C37" value="1" class="clearBorder"></td><td>Drop-Shipper ID</td></tr>
		
		<tr><td align="right"><input type="checkbox" name="C39" value="1" class="clearBorder"></td><td>Meta Tags Information</td></tr>
		
		<tr><td align="right"><input type="checkbox" name="C40" value="1" class="clearBorder"></td><td>Downloadable Products Information</td></tr>
		<tr><td align="right"><input type="checkbox" name="C41" value="1" class="clearBorder"></td><td>Gift Certificates Information</td></tr>
		<%if scBTO=1 then%>
		<tr><td align="right"><input type="checkbox" name="C42" value="1" class="clearBorder"></td><td>Hide Configurator Prices</td></tr>
		<tr><td align="right"><input type="checkbox" name="C43" value="1" class="clearBorder"></td><td>Hide Default Configuration</td></tr>
		<tr><td align="right"><input type="checkbox" name="C44" value="1" class="clearBorder"></td><td>Disallow Purchasing</td></tr>
		<tr><td align="right"><input type="checkbox" name="C45" value="1" class="clearBorder"></td><td>Skip Product Details Page</td></tr>
		<%end if%>

    <%end if%>

	<%if session("cp_revImport_extype")="0" or session("cp_revImport_extype")="2" then%>
		<tr><td align="right"><input type="checkbox" name="C49" value="1" class="clearBorder"></td><td>Units to make 1 lb</td></tr>
		<tr><td align="right"><input type="checkbox" name="C50" value="1" class="clearBorder"></td><td>First Unit Surcharge</td></tr>
		<tr><td align="right"><input type="checkbox" name="C51" value="1" class="clearBorder"></td><td>Additional Unit(s) Surcharge</td></tr>
		<tr><td align="right"><input type="checkbox" name="C52" value="1" class="clearBorder"></td><td>Product Notes</td></tr>
		<tr><td align="right"><input type="checkbox" name="C53" value="1" class="clearBorder"></td><td>Enable Image Magnifier</td></tr>
        <tr><td align="right"><input type="checkbox" name="C67" value="1" class="clearBorder"></td><td>Hide Additional Images</td></tr>
		<tr><td align="right"><input type="checkbox" name="C54" value="1" class="clearBorder"></td><td>Page Display Layout</td></tr>
		<tr><td align="right"><input type="checkbox" name="C59" value="1" class="clearBorder"></td><td>Custom Product Page Layout</td></tr>
		<tr><td align="right"><input type="checkbox" name="C55" value="1" class="clearBorder"></td><td>Hide SKU on the product details page</td></tr>
		<tr><td align="right"><input type="checkbox" name="CSearchFields" value="1" class="clearBorder"></td><td>Product Search Fields</td></tr>
		<tr><td align="right"><input type="checkbox" name="C56" value="1" class="clearBorder"></td><td>Google Product Category</td></tr>
    <%end if%>

	<tr><td align="right"><input type="checkbox" name="C57" value="1" class="clearBorder"></td><td>Google Shopping - Product Attributes</td></tr>

	<%if session("cp_revImport_extype")="0" or session("cp_revImport_extype")="2" then%>
		<tr><td align="right"><input type="checkbox" name="C60" value="1" class="clearBorder"></td><td>Show out of stock items</td></tr>
		<tr><td align="right"><input type="checkbox" name="C61" value="1" class="clearBorder"></td><td>Out of stock items message</td></tr>
		<tr><td align="right"><input type="checkbox" name="C62" value="1" class="clearBorder"></td><td>Display type</td></tr>
		<tr><td align="right"><input type="checkbox" name="C63" value="1" class="clearBorder"></td><td>Size chart text link</td></tr>
		<tr><td align="right"><input type="checkbox" name="C64" value="1" class="clearBorder"></td><td>Size chart description</td></tr>
		<tr><td align="right"><input type="checkbox" name="C65" value="1" class="clearBorder"></td><td>Size chart image file</td></tr>
		<tr><td align="right"><input type="checkbox" name="C66" value="1" class="clearBorder"></td><td>Size chart image URL</td></tr>
	<%end if%>

	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
			<script type=text/javascript>
				function checkAll() {
					var theForm, z = 0;
					theForm = document.checkboxform;
					 for(z=0; z<theForm.length;z++){
					  if(theForm[z].type == 'checkbox'){
					  theForm[z].checked = true;
					  }
					}
				}
				 
				function uncheckAll() {
					var theForm, z = 0;
					theForm = document.checkboxform;
					 for(z=0; z<theForm.length;z++){
					  if(theForm[z].type == 'checkbox'){
					  theForm[z].checked = false;
					  }
					}
				}
				
				function testCheckBox()
				{
					var theForm, z = 0;
					theForm = document.checkboxform;
					 for(z=0; z<theForm.length;z++){
					  if((theForm[z].type == 'checkbox') && (theForm[z].checked == true)) {
					  return(true);
					  }
					}
				
					return(false);
				}
			</script>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2">
        
        	<input type="submit" name="submit" value=" Export products " class="btn btn-primary" onClick="javascript: if (testCheckBox()) { pcf_Open_Import(); return(confirm('You are about to export the selected product fields. Are you sure you want to complete this action?')); } else { return(false); }">
        
			<%
            '// Loading Window
            '	>> Call Method with OpenHS();
            response.Write(pcf_ModalWindow("This could take several minutes. Do not close this page.", "Import", 300))
            %>
        </td>
	</tr>
</table>
</FORM>
<!--#include file="AdminFooter.asp"-->