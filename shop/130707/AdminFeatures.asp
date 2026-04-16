<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Featured Products" %>
<% section="specials" %>
<%PmAdmin=3%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim pid

'****************************
'* Action Items
'****************************

	'// UPDATE list of featured products
		sMode=Request("Submit")
		If sMode <> "" Then

			if (request("prdlist")<>"") and (request("prdlist")<>",") then
			
				prdlist=split(request("prdlist"),",")
			   
				For i=lbound(prdlist) to ubound(prdlist)
					id=prdlist(i)
					If (id<>"0") and (id<>"") Then

						query="UPDATE products SET showInHome=-1 WHERE idproduct="&id
						Set rstemp2 = Server.CreateObject("ADODB.Recordset")
						Set rstemp2 = connTemp.execute(query)
						Set rstemp2 = Nothing
						call pcs_hookProductModified(id, "")

					End If
				Next

				call closeDb()
				response.redirect "AdminFeatures.asp?s=1&msg="&msg
			end if

			'// UPDATE featured products order
			PCount=request("PCount")
			if (pCount<>"") and (PCount<>"0") then
				
				For i=1 to Cint(PCount)
					idproduct=request("IDFP" & i)
					OrdInHome=request("FPOrd" & i)
                    If len(idproduct)>0 And len(OrdInHome)>0 Then
					    
                        query="UPDATE products SET pcprod_OrdInHome=" & OrdInHome & " WHERE idproduct=" & idproduct
					    set rstemp=Server.CreateObject("ADODB.Recordset")
						set rstemp=connTemp.execute(query)
                        set rstemp=nothing
						call pcs_hookProductModified(idproduct, "")
                    End If
				Next

				call closeDb()
				response.redirect "AdminFeatures.asp?s=1&msg="&Server.URLEncode("Product Orders were updated successfully!")
			end if

		End If '// If sMode <> "" Then

'****************************
'* END Action Items
'****************************

	' Paging and sorting
	
	if request("iPageCurrent")="" then
    iPageCurrent=1 
	else
		iPageCurrent=Request("iPageCurrent")
	end If

	' gets group assignments
	query="SELECT pcprod_OrdInHome,idproduct,sku,description,smallimageurl FROM products WHERE active=-1 AND showInHome=-1 ORDER BY pcprod_OrdInHome, description"

	Set rstemp=Server.CreateObject("ADODB.Recordset")   
	rstemp.CacheSize=100
	rstemp.PageSize=100

	rstemp.Open query, connTemp, adOpenStatic, adLockReadOnly
	dontshow="0"
	If rstemp.eof Then 
		dontshow="1"
	end if
	
	' Find out if all products have been set as Featured Product
	query="SELECT idProduct FROM products WHERE active=-1 AND configOnly=0 AND removed=0 AND showInHome=0"
	Set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if rs.EOF then
		pcAllFeatured = 1
	end if
	set rs = nothing
%>
	<table class="pcCPcontent">
		<tr>
			<td>
			<p>Featured products are shown on the &quot;<a href="http://wiki.productcart.com/productcart/marketing-featured_products" target="_blank">home page</a>&quot; and on the &quot;featured product&quot; page, in the order specified below.&nbsp;<a href="http://wiki.productcart.com/productcart/marketing-featured_products" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information about this feature." width="16" height="16" border="0"></a></p>
            <div class="cpOtherLinks"><% if pcAllFeatured <> 1 then %><a href="JavaScript:;" onClick="javascript:document.getElementById('FindProducts').style.display=''">Add New</a>&nbsp;|&nbsp;<% end if %>View the store's <a href="../pc/home.asp" target="_blank">home page</a>&nbsp;|&nbsp;View the store's <a href="../pc/showfeatured.asp" target="_blank">featured products page</a>&nbsp;|&nbsp;<a href="manageHomePage.asp">Manage the home page</a></div>
            <table id="FindProducts" class="pcCPcontent" style="display:none;">
                <tr>
                    <td>
                    <%
                        src_FormTitle1="Find Products"
                        src_FormTitle2="Add New Featured Products"
                        src_FormTips1="Use the following filters to look for products in your store."
                        src_FormTips2="Select the products that you would like to add to your list of featured products."
                        src_IncNormal=1
                        src_IncBTO=1
                        src_IncItem=0
                        src_Featured=2
                        src_DisplayType=1
                        src_ShowLinks=0
                        src_FromPage="AdminFeatures.asp"
                        src_ToPage="AdminFeatures.asp?submit=yes"
                        src_Button1=" Search "
                        src_Button2=" Add as Featured Product "
                        src_Button3=" Back "
                        src_PageSize=15
                    %>
                        <!--#include file="inc_srcPrds.asp"-->
                    </td>
                </tr>
            </table>

			<% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	 
	 <tr>
	 	<td>
	
			<form action="AdminFeatures.asp" method="post" name="form1" class="pcForms">
				<!--<table class="pcCPcontent">-->

					<div class="pcCPsortableTableHeader">
						<div class="pcCPsortableTableIndex">#</div>
						<div class="pcCPfeaturedProductSKU">SKU</div>
						<div class="pcCPfeaturedProductName">Product Name</div>
						<div class="pcCPfeaturedProductActions"></div>
					</div>

					<% If rstemp.eof Then %>
						<div class="pcCPmessage">No Featured Items Found</div>
					<% Else %>
						<ul class="pcCPsortable pcCPsortableTable">
							<%
								rstemp.MoveFirst
								' get the max number of pages
								Dim iPageCount
								iPageCount=rstemp.PageCount
								If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
								If iPageCurrent < 1 Then iPageCurrent=1
														
								' set the absolute page
								rstemp.AbsolutePage=iPageCurrent
								Dim Count
								hCnt=0
								Count=0
								Do While NOT rstemp.EOF
									Count=Count+1
								
									pIdProduct	= rstemp("idproduct")
									pSKU				= rstemp("sku")
									pDesc				= rstemp("description")
									pSmallImage	= rstemp("smallImageUrl")
									%>
									<li class="cpItemlist">
										<div class="pcCPsortableTableIndex">
											<span class="pcCPsortableIndex"><%=Count%></span>
											<input type="hidden" class="pcCPsortableOrder" name="FPOrd<%=Count%>" value="<%=FPOrd%>">
											<input type="hidden" name="IDFP<%=Count%>" value="<%=pIdProduct%>">
										</div>
										<div class="pcCPfeaturedProductSKU">
											<a href="FindProductType.asp?id=<%=pIdProduct%>" target="_blank"><%=pSKU%></a>
										</div>
										<div class="pcCPfeaturedProductName">
											<% If Len(pSmallImage) > 0 Then %>
												<img src="../pc/catalog/<%= pSmallImage %>" />
											<% End If %>
											<a href="FindProductType.asp?id=<%=pIdProduct%>" target="_blank"><%=pDesc%></a>
										</div>
										<div class="pcCPfeaturedProductActions">
											<a href="javascript:if (confirm('You are about to remove this item as a featured item. Are you sure you want to complete this action?')) location='delFeaturesb.asp?idproduct=<%= pIdProduct %>'" title="Remove Featured Product"><img src="images/pcIconDelete.jpg"></a>
										</div>
									</li>
									<%
								rstemp.MoveNext
								Loop
								%>			
						</ul>
						<input type=hidden name="PCount" value="<%=Count%>">	
					<% 
						End If
						set rstemp = nothing
					%>

					<%if dontshow="0" then%>
					<table class="pcCPcontent">
						<tr>
							<td colspan="4" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="4" align="center">
								<input type="Submit" name="Submit" value="Update Order" class="btn btn-primary">
								&nbsp;<input type="button" class="btn btn-default"  onClick="location.href = 'manageHomePage.asp'" value="Manage Home Page">
							</td>
						</tr>
					</table>
					<%
					end if
					%>
			</form>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->
