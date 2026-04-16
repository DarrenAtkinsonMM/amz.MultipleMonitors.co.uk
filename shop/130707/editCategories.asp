<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%> 
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
Dim il_intListID, il_categoryDesc, il_strMode, pcStrPageName

pcStrPageName="editCategories.asp"

if request("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=Request("iPageCurrent")
end If
if request("nav")="1" then
	section="services"
else
	section="products"
end if

il_strMode=""

Dim reqstr, reqidproduct, reqidcategory
reqstr=Request.QueryString("reqstr")
reqidproduct=Request.QueryString("reqidproduct")
reqidcategory=Request.QueryString("reqidcategory")
If reqstr="" then
	reqstr=Request.Form("reqstr")
	reqidproduct=Request.Form("reqidproduct")
	reqidcategory=Request.Form("reqidcategory")
End If
	
If isNumeric(Request.QueryString("lid")) AND Request.QueryString("lid") <> "" Then
	il_intListID=Request.QueryString("lid")
ElseIf isNumeric(Request.QueryString("lid")) AND Request.Form("lid") <> "" Then
	il_intListID=Request.Form("lid")
Else
	call closeDb()
	response.redirect "../manageCategories.asp"
End If

If Request.QueryString("mode") <> "" Then
	il_strMode=Request.QueryString("mode")
ElseIf Request.Form("mode") <> "" Then
	il_strMode=Request.Form("mode")
Else
	il_strMode="view"
End If

Set rs=Server.CreateObject("ADODB.Recordset")

strSQL="SELECT categoryDesc, idCategory, pcCats_ProductOrder, (SELECT COUNT(*) FROM categories_products WHERE POrder > 0 AND idCategory = " & il_intListID & ") AS numOrdered FROM categories WHERE idCategory=" & il_intListID & " ORDER BY categoryDesc"
rs.Open strSQL, conntemp, adOpenStatic, adLockReadOnly

If err.number <> 0 Then
	TrapError Err.Description
Else
	il_categoryDesc=rs("categoryDesc")
  	il_productOrder=rs("pcCats_ProductOrder")
  	il_numOrdered=CInt(rs("numOrdered"))
End If
	
rs.Close

If il_numOrdered > 0 Then
  SOrder = "POrder"
Else
  If il_productOrder <> "" Then
    POrder = il_productOrder
  Else
    POrder = PCOrd
  End If

  Select Case POrder
  Case "0":
    SOrder = "sku"
  Case "1": 
    SOrder = "description"
  Case "2":
    SOrder = "price desc"
  Case "3":
    SOrder = "price asc"
  Case Else:
    SOrder = "sku"
  End Select
End If

pageTitle="Products currently assigned to: " & il_categoryDesc
%>
<!--#include file="AdminHeader.asp"-->
<%src_checkPrdType=request("cptype")
if src_checkPrdType="" then
	src_checkPrdType="0"
end if%>

	<div class="cpOtherLinks"><a href="JavaScript:;" onClick="document.getElementById('FindProducts').style.display=''; document.getElementById('pcHideInactive').style.display='none';">Add new products</a> to &quot;<%=il_categoryDesc%>&quot;&nbsp;|&nbsp;<a href="updPrdPrices.asp?idcategory=<%=il_intListID%>">Update Product Prices</a>&nbsp;|&nbsp;<a href="../pc/viewcategories.asp?idcategory=<%=il_intListID%>" target="_blank">View in the storefront</a>
	</div>
    
    <table id="FindProducts" class="pcCPcontent" style="display:none;">
        <tr>
            <td>
            <div id="FindProductsClose" style="float: right; padding: 18px 80px 0 0;"><a href="JavaScript:;" onClick="document.getElementById('FindProducts').style.display='none'; document.getElementById('pcHideInactive').style.display='';"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Close this panel"></a></div>
            <%
                src_ShowPrdTypeBtns=1
                src_FormTitle1="Find Products"
                src_FormTitle2="Add Products to " & il_categoryDesc
                src_FormTips1="Use the following filters to look for products in your store."
                src_FormTips2="Select the products that you would like to add to the '" & il_categoryDesc & "' category."
                src_IncNormal=0
                src_IncBTO=0
                src_IncItem=0
                src_DisplayType=1
                src_ShowLinks=0
                src_FromPage="editCategories.asp?iPageCurrent=1&Sorder=" & request("SOrder") & "&Sort=" & request("Sort") & "&nav=" & request("nav") & "&lid=" & il_intListID & "&mode=view"
                src_ToPage="actionCategories.asp?Sorder=" & request("SOrder") & "&Sort=" & request("Sort") & "&nav=" & request("nav") & "&lid=" & il_intListID & "&mode=edit&reqstr=" & reqstr & "&reqidproduct=" & reqidproduct & "&reqidcategory=" & reqidcategory
                src_Button1=" Search "
                src_Button2=" Add products to " & il_categoryDesc
                src_Button3=" Back "
                src_PageSize=15
                UseSpecial=1
                session("srcprd_from")=""
                session("srcprd_where")=" AND (products.idProduct NOT IN (SELECT idProduct FROM categories_products WHERE idCategory=" & il_intListID & ")) "
            %>
                <!--#include file="inc_srcPrds.asp"-->
            </td>
        </tr>
    </table>
	
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
	
    <div>
    	<h5 class="help-block">
            <span class="glyphicon glyphicon-info-sign" aria-hidden="true"></span>
            To modify product order, drag &amp; drop to the desired location, then click the "Update Product Order" button.
    	</h5>
    </div>
     
	<form method="POST" action="actionCategories.asp" name="myForm" class="pcForms">
		<input type="hidden" name="mode" value="<%=il_strMode%>"><input type="hidden" name="lid" value="<%= il_intListID %>">
		<input type="hidden" name="reqstr" value="<%=reqstr%>">
		<input type="hidden" name="reqidproduct" value="<%=reqidproduct%>">
		<input type="hidden" name="reqidcategory" value="<%=reqidcategory%>">
		<input type="hidden" name="nav" value="<%=request("nav")%>"> 
		<input type="hidden" name="SOrder" value="<%=SOrder%>">
		<input type="hidden" name="Sort" value="<%=Sort%>">
		<input type="hidden" name="frmHideInactiveProducts" value="<%=Request("frmHideInactiveProducts")%>" >
        
    	<div id="pcHideInactive" style="float: right; margin-top: 5px; margin-right: 5px;">
        <input class="clearBorder" onclick="javascript:reloadFormToShowHideInactiveProducts();" type="checkbox" <%if Request("frmHideInactiveProducts") = "true" then%>checked<%end if %> name="chkHideInActiveProducts" />&nbsp;Hide inactive products
        </div>
		

    <script type=text/javascript>
    </script>
    
		<div class="pcCPsortableTableHeader">
      <div class="pcCPsortableTableIndex">#</div>
			<div class="pcCPcatProductCheck">&nbsp;</div>
			<div class="pcCPcatProductSKU">SKU</div>
			<div class="pcCPcatProductName">Product Name</div>
		</div>

			<%
			strSQL="SELECT A.idProduct AS idProduct,smallImageUrl,description,active,A.configOnly,A.serviceSpec,A.sku,AL.POrder FROM products A, categories_products AL, categories L WHERE A.removed=0"
			if ( Request("frmHideInactiveProducts") = "true" ) then
			    strSQL=strSQL&" AND A.Active = -1 AND A.idProduct=AL.idProduct AND AL.idCategory=L.idCategory AND L.idCategory=" & il_intListID & " ORDER BY " & Sorder & " " & Sort
			else
			    strSQL=strSQL&" AND A.idProduct=AL.idProduct AND AL.idCategory=L.idCategory AND L.idCategory=" & il_intListID & " ORDER BY " & Sorder & " " & Sort
			end if			  
            
            Dim iCnt
            iCnt=0
              
			set rs=Server.CreateObject("ADODB.Recordset")
            rs.Open strSQL, conntemp
			If rs.EOF OR rs.BOF Then 
            %>
		        <div class="pcCPmessage">No Products Found</div>
			<% Else %>
		        <ul class="pcCPsortable pcCPsortableTable">
                <%
				Do While NOT rs.EOF
                    pIdProduct = rs("idProduct")
                    psmallImageUrl = rs("smallImageUrl")
                    pdescription = rs("description")
                    porder = rs("POrder")

                    If Len(psmallImageUrl) < 1 Then
                        psmallImageUrl = "no_image.gif"
                    End If

			        iCnt=iCnt+1
					%>
					<li class="cpItemlist"> 
                        <div class="pcCPsortableTableIndex">
                            <span class="pcCPsortableIndex"><%= iCnt %></span>
                                <input type="hidden" class="pcCPsortableOrder" name="POrder" value="<%= porder %>"/>
                                <input type="hidden" name="listidproduct" value="<%= pIdProduct %>">
                        </div>
						<div class="pcCPcatProductCheck" >
                            <input type="checkbox" name="Address" value="<%= pIdProduct %>" class="clearBorder">
						</div>
						<div class="pcCPcatProductSKU">
                            <%=rs("sku")%>
						</div>
						<div class="pcCPcatProductName">
                            <% If Len(psmallImageUrl) > 0 Then %>
                                <img src="../pc/catalog/<%= psmallImageUrl %>" />
                            <% End If %>
							<%if (rs("configOnly")=0) and (rs("serviceSpec")<>0) then%>
								<a href="FindProductType.asp?id=<%= pIdProduct %>"><%= pdescription %></a>&nbsp;
								<span style="color: #FF0000">(configurable product)</span>
							<%elseif (rs("configOnly")<>0) and (rs("serviceSpec")=0) then%>
								<a href="FindProductType.asp?id=<%= pIdProduct %>"><%= pdescription %></a>&nbsp;
								<span style="color: #FF0000">(configurable-only item)</span>
							<%else%>
								<a href="FindProductType.asp?id=<%= pIdProduct %>"><%= pdescription %></a>
							<%end if%>
							<% if rs("active")<>-1 then %>
								&nbsp;<span style="color: #FF0000">(Inactive)</span>
							<% end if %>
						</div>
					</li>
						<% count=count + 1
						rs.MoveNext
					Loop
					%>
		</ul>

		<div class="pcCPsortableTableFooter">
      <div class="pcCPsortableTableIndex">&nbsp;</div>
		  <div class="pcCPcatProductCheck"> 
			  <input type="checkbox" value="ON" onClick="javascript: checkTheBoxes();" class="clearBorder">
		  </div>
		  <div class="pcCPcatProductSKU">Select All</div>
    </div>

    <br />

	<% End If ' end of rs.eof%>

	<table class="pcCPcontent">
    <% If iCnt>0 Then %>
    <tr>
      <td colspan="4">
        <input type="submit" value="Remove Checked" class="btn btn-primary" onclick="javascript: if (!(checkedBoxes())) { alert('Please select at least one product to remove from this category.'); return (false) }">
							&nbsp;<input type="submit" name="UpdateOrder" value="Update Product Order" class="btn btn-primary">
                            &nbsp;<input type="submit" name="ResetOrder" value="Reset Product Order" class="btn btn-primary" onClick="JavaScript: if (confirm('PLEASE NOTE: you are about to reset to &quot;0&quot; the Order value for all products in this category. If you do so, products will be ordered based on the general product sorting criteria set under &quot;Store Settings > Display Settings&quot;. Would you like to continue?'));">
                            &nbsp;<input type="submit" name="CopyTo" value="Copy Selected to..." onclick="javascript: if (!(checkedBoxes())) { alert('Please select at least one product to copy to another category.'); return (false) }">
                            &nbsp;<input type="submit" name="MoveTo" value="Move Selected to...." onclick="javascript: if (!(checkedBoxes())) { alert('Please select at least one product to move to another category.'); return (false) }">
      </td>
    </tr>
    <% End If %>
    
	<% if reqstr<>"" then %>
	<tr>
		<td colspan="4">
			<% if il_intListID <> 1 then %>
			<input type="button" class="btn btn-default"  value="Edit Category" onClick="document.location.href='modCata.asp?idcategory=<%=il_intListID%>'" name="button">&nbsp;
			<% end if %>
			<input type="button" class="btn btn-default"  value="Manage Categories" onClick="document.location.href='<%=reqstr%>&nav=<%=request("nav")%>&idproduct=<%=reqidproduct%>&idcategory=<%=reqidcategory%>';" name="button">
		</td>
	</tr>
	<% else %>
	<tr>
		<td colspan="4">
			<% if il_intListID <> 1 then %>
			<input type="button" class="btn btn-default"  value="Edit Category" onClick="document.location.href='modCata.asp?idcategory=<%=il_intListID%>'">&nbsp;
			<% end if %>
			<input type="button" class="btn btn-default"  value="Manage Categories" onClick="document.location.href='manageCategories.asp?nav=<%=request("nav")%>';">
		</td>
	</tr>
	<% end if %>
    
	</table>
</form>
<script type=text/javascript>
function checkedBoxes()
{
	with (document.myForm) {
		for(i=0; i < elements.length -1; i++) {
			if ( elements[i].name== "Address" ) {
				if (elements[i].checked==true) return(true);
			}
		}
	}
	return(false);
}
</script>
<!--#include file="Adminfooter.asp" -->
