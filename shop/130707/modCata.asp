<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Modify Category" %>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" -->

<% dim pIdCategory

pIdCategory=request.Querystring("idCategory")
if NOT validNum(pIdCategory) then
 call closeDb()
response.redirect "msg.asp?message=3"
end if

pcv_strCSFId = pIdCategory

Dim top, parent
top=request.QueryString("top")

if NOT validNum(top) then	top = 1

parent=request.QueryString("parent")
if NOT validNum(parent) then	parent = 1

' get data of the category to modify
query="SELECT idParentCategory, categoryDesc, [image], largeimage, iBTOhide, SDesc, LDesc, HideDesc, pcCats_RetailHide, pcCats_SubCategoryView, pcCats_CategoryColumns, pcCats_CategoryRows, pcCats_PageStyle, pcCats_ProductOrder, pcCats_ProductColumns, pcCats_ProductRows, pcCats_FeaturedCategory, pcCats_FeaturedCategoryImage, pcCats_DisplayLayout, pcCats_MetaTitle, pcCats_MetaDesc, pcCats_MetaKeywords, pcCats_NotImg, pcCats_AvalaraTaxCode FROM categories WHERE idCategory=" &pIdCategory&";"

set rs=Server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	
 	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("There was an error retrieving data for this category. The issue might be that you clicked on an outdated link. Go to 'Products > Manage Categories > View and Edit Categories' and update the category navigation used in the Control Panel. You will see text at the top that says 'If the list below does not seem current, please click here (updates XML cache).' Click on that link to update category information used to navigate categories on that page.") 
end If

pIdParentCategory=rs("idParentCategory")
pCategoryDesc=replace(rs("categoryDesc"), """", "&quot;")
pCategoryDesc=replace(pCategoryDesc, "&amp;", "&")
pImage=rs("image")
plargeImage=rs("largeimage")
piBTOhide=rs("iBTOhide")
SDesc=rs("SDesc")
LDesc=rs("LDesc")
HideDesc=rs("HideDesc")
pcv_intRetailHide=rs("pccats_RetailHide")
pIntSubCategoryView=rs("pcCats_SubCategoryView")
pIntCategoryColumns=rs("pcCats_CategoryColumns")
pIntCategoryRows=rs("pcCats_CategoryRows")
pStrPageStyle=rs("pcCats_PageStyle")
pStrProductOrder=rs("pcCats_ProductOrder")
pIntProductColumns=rs("pcCats_ProductColumns")
pIntProductRows=rs("pcCats_ProductRows")
pIntFeaturedCategory=rs("pcCats_FeaturedCategory")
pIntFeaturedCategoryImage=rs("pcCats_FeaturedCategoryImage")
pStrCatDisplayLayout=rs("pcCats_DisplayLayout")
	if IsNull(pStrCatDisplayLayout) then pStrCatDisplayLayout=""
pStrCatMetaTitle=rs("pcCats_MetaTitle")
pStrCatMetaDesc=rs("pcCats_MetaDesc")
pStrCatMetaKeywords=rs("pcCats_MetaKeywords")
NotImg=rs("pcCats_NotImg")
pStrCatAvalaraTaxCode=rs("pcCats_AvalaraTaxCode")

if NOT validNum(pIntSubCategoryView) then pIntSubCategoryView=3
if NOT validNum(pIntCategoryColumns) then pIntCategoryColumns=0
if NOT validNum(pIntCategoryRows) then pIntCategoryRows=0
if NOT validNum(pIntProductColumns) then pIntProductColumns=0
if NOT validNum(pIntProductRows) then pIntProductRows=0
if NOT validNum(pIntFeaturedCategory) then pIntFeaturedCategory=0
if NOT validNum(pIntFeaturedCategoryImage) then pIntFeaturedCategoryImage=0
if NOT validNum(HideDesc) then HideDesc=0
if NOT validNum(pcv_intRetailHide) then pcv_intRetailHide=0
	
set rs=nothing

query = "SELECT COUNT(*) AS numOrdered FROM categories_products WHERE POrder > 0 AND idCategory = " & pIdCategory & ";"
set rs=Server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number <> 0 then
	set rs=nothing
	
 	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("There was an error retrieving data for this category. The issue might be that you clicked on an outdated link. Go to 'Products > Manage Categories > View and Edit Categories' and update the category navigation used in the Control Panel. You will see text at the top that says 'If the list below does not seem current, please click here (updates XML cache).' Click on that link to update category information used to navigate categories on that page.") 
end If

pIntNumOrdered=rs("numOrdered")
If Not validNum(pIntNumOrdered) Then pIntNumOrdered=0

Set rs=nothing



'// Add category name to Page Title
pageTitle=pageTitle & ": " & pCategoryDesc
%>
  
<!--#include file="AdminHeader.asp"-->

<!-- #include file="../htmleditor/editor.asp" -->

<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>

<script type=text/javascript>
	function newWindow(file,window) {
			msgWindow=open(file,window,'resizable=no,width=400,height=500');
			if (msgWindow.opener == null) msgWindow.opener = self;
	}

	function Form1_Validator(theForm)
	{
		// InnovaStudio HTML Editor Workaround for this keyword
		theForm = document.hForm;

		if (theForm.categoryDesc.value == "")
			{
				 alert("Please enter a name for this category.");
					theForm.categoryDesc.focus();
					return (false);
			}
	return (true);
	}


	function OpenHS() {
		if (document.hForm.runSubCats.checked==true) {
			pcf_Open_Modify10();
		}
	}
</script>
<%
	if request.QueryString("message")="OK1" then
		msg="Category updated successfully! If you are using &quot;Static Navigation&quot; in your storefront, remember to update the navigation files using the <a href=genCatNavigation.asp target=_blank>Generate Navigation</a> feature."
		msgType=1
    end if
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form action="modCatb.asp" method="post" name="hForm" onSubmit="return Form1_Validator(this)" class="pcForms">	

	<% '// TEXT NAVIGATION - Start %>
        <div class="cpOtherLinks" style="margin: 0 12px 0 7px;">
            <%
			Dim pcIntHidden
			pcIntHidden=0
			if (piBTOhide<>"") and (piBTOhide="1") then pcIntHidden=1
			if pcIntHidden=1 then
			%>
            Hidden in the storefront
            <%else%>
            <a href="../pc/viewcategories.asp?idcategory=<%=pIdCategory%>" target="_blank">View in the storefront</a></font>
            <%end if%>
            &nbsp;|&nbsp;
            <a href="editCategories.asp?nav=&lid=<%=pIdCategory%>">View/Add Products</a>
            &nbsp;|&nbsp;
            <a href="updPrdPrices.asp?idcategory=<%=pIdCategory%>">Update Product Prices</a>
            &nbsp;|&nbsp;            
            <a href="viewCat.asp?parent=<%=pIdCategory%>&hidden=<%=pcIntHidden%>">Subcategories</a>
            <!-- Feature not ready yet
            &nbsp;|&nbsp;            
            <a href="ModPromotionCat.asp?idcategory=<%=pIdCategory%>&iMode=start">Promotion</a>
            -->
            &nbsp;|&nbsp;
            <a href="ManageSearchFields.asp">Search Fields</a>
            &nbsp;|&nbsp;
            <a href="AddDupCat.asp?idcategory=<%=pIdCategory%>&top=<%=request("top")%>&parent=<%=request("parent")%>">Clone</a>
            &nbsp;|&nbsp;
            <a href="manageCategories.asp">Manage All</a>
        </div>
	<% '// TEXT NAVIGATION - End %>

	
		<%
		'// TABBED PANELS - MAIN DIV START
		%>
	  <div id="TabbedPanels1" class="tabbable-left">
		
		<%
		'// TABBED PANELS - START NAVIGATION
		%>
		<div class="col-xs-3">
            <ul class="nav nav-tabs tabs-left">
				<li class="active"><a href="#tabs-1" data-toggle="tab">Name, Parent &amp; Images</a></li>
				<li><a href="#tabs-2" data-toggle="tab">Descriptions</a></li>
				<li><a href="#tabs-3" data-toggle="tab">Display Settings</a></li>				
				<li><a href="#tabs-4" data-toggle="tab">Other Settings</a></li>
				<li><a href="#tabs-5" data-toggle="tab">Meta Tags</a></li>
                <li><a href="#tabs-6" data-toggle="tab">Search Fields</a></li>
				<li>
					<div style="margin-top:10px; margin-bottom:10px; text-align: center">
					<input type="hidden" name="idcategory" value='<%=pIdCategory%>'>
					<input type="hidden" name="top" value='<%=top%>'>
					<input type="hidden" name="parent" value='<%=parent%>'>
					<input type="submit" name="modify" value="Save" onClick="OpenHS();" class="btn btn-primary">
					<div style="padding: 6px 0 15px 0; text-align:left; color: #000000; font-family: Arial, sans-serif; font-weight: normal; font-size: 12px;">
					<input type="checkbox" name="runSubCats" value="1" class="clearBorder" style="vertical-align:middle"> Apply to sub-categories
					</div>
					<div style="text-align: center;">
					<input type="submit" class="btn btn-default"  name="delete" value="Delete" onClick="return confirm('Are you sure you want to delete this category? If the category contains any products, you will first be prompted to remove them from it. If no products have been assigned, the category will be immediately deleted.')">
					&nbsp;
					<input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">
					</div>
				</li>
			</ul>
		</div>
		<%
		'// TABBED PANELS - END NAVIGATION
		
		'// TABBED PANELS - START PANELS
		%>
        <div class="col-xs-9">
            <div class="tab-content">
		
			<%
			'// =========================================
			'// FIRST PANEL - START - Name, Descriptions, Images
			'// =========================================
			%>
				<div id="tabs-1" class="tab-pane active">
				
					<table class="pcCPcontent">				
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Category Name, Images and Parent Category</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td width="20%" align="right">Category Name:</td>
							<td width="80%"><input name="categoryDesc" type="text" value="<%=pCategoryDesc%>" size="40" tabindex="101"></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>       
							<td align="right">Small Image:</td>
							<td>
								<input type="text" name="image" value="<%=pImage%>" size="40" tabindex="102"><a href="#" onClick="chgWin('../pc/imageDir.asp?ffid=image&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=439"></a>
							</td>
						</tr>
						<tr> 
							<td align="right">Large Image:</td>
							<td> 
								<input type="text" name="largeimage" value="<%=plargeImage%>" size="40" tabindex="103"><a href="#" onClick="chgWin('../pc/imageDir.asp?ffid=largeimage&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=439"></a> 
							</td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td><p> 
								<!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->
								<%If HaveImgUplResizeObjs=1 then%>
									To upload and resize an image <a href="#" onClick="window.open('uploadresize/catResizea.asp','popup','toolbar=no,status=no,location=no,menubar=no,height=350,width=400,scrollbars=no'); return false;">click here</a>.
								<% Else %>
									To upload an image <a href="#" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">click here</a>.
								<% End If %>
							</td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td>
								<input type="checkbox" name="NotImg" value="1" <%if NotImg="1" then%>checked<%end if%> class="clearBorder"> Hide category images, except when using "Thumbnails only" under display settings.  
							</td>
						</tr>
						<tr>
							<td colspan="2"><hr></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td align="right">Parent Category:</td>
							<td>
								<%
								cat_DropDownName="idParentCategory"
								cat_DropDownTabIndex="104"
								cat_Type="0"
								cat_DropDownSize="1"
								cat_MultiSelect="0"
								cat_ExcBTOHide="0"
								cat_StoreFront="0"
								cat_ShowParent="1"
								cat_DefaultItem=""
								cat_SelectedItems="" & pIdParentCategory & ","
								cat_ExcItems="" & pIdCategory & ","
								cat_ExcSubs="0"
								cat_ExcCircle=pIdCategory
								%>
								
								<%call pcs_CatList()%>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>

					</table>
					
				</div>
			<%
			'// =========================================
			'// FIRST PANEL - END
			'// =========================================
			
			'// =========================================
			'// SECOND PANEL - START - Descriptions
			'// =========================================
			%>
				<div id="tabs-2" class="tab-pane">

					<table class="pcCPcontent">	
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Category Descriptions&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=440"></a></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr valign="top">
							<td width="30%">Short Description
								<div class="small">Shown on pages that display categories</div>
							</td>
							<td width="70%">
								<textarea name="SDesc" id="SDesc" cols="60" rows="6" tabindex="201" maxlength="255"><%=SDesc%></textarea>
								<div class="pcSmallText">Maximum Length: 255 Characters</div>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr valign="top">
							<td>Long Description
							<div class="small">Only shown on the page that displays products within the category</div>
							<td>
							  <textarea class="htmleditor" name="LDesc" id="LDesc" cols="60" rows="6" tabindex="202"><%=LDesc%></textarea>				
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td align="right">
							<input type="checkbox" name="HideDesc" value="1" <%if HideDesc="1" then%>checked<%end if%> class="clearBorder" tabindex="203">
							</td>
							<td>Do not show category descriptions</td>
						</tr>
						
					</table>
					
				</div>
			<%
			'// =========================================
			'// SECOND PANEL - END
			'// =========================================
			
			'// =========================================
			'// THIRD PANEL - START - Display settings
			'// =========================================
			%>
				<div id="tabs-3" class="tab-pane">

					<table class="pcCPcontent">	
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Display Settings</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td colspan="2">
							Choose a display option for how <a href="manageCategories.asp?nav=&top=1&parent=<%=pIdCategory%>" target="_blank">subcategories</a> are displayed (if no option is selected, the default <a href="AdminSettings.asp?tab=3">store-wide setting</a> is used):
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td align="right" width="30%">Display Subcategories:</td>
							<td width="70%">
								<select name="intSubCategoryView" tabindex="301">
									<option value="3"<% if pIntSubCategoryView="3" then %> selected<%end if%>>Default</option>
									<option value="2"<% if pIntSubCategoryView="2" then %> selected<%end if%>>In a drop-down</option>
									<option value="0"<% if pIntSubCategoryView="0" then %> selected<%end if%>>In a list, with images</option>
									<option value="1"<% if pIntSubCategoryView="1" then %> selected<%end if%>>In a list, without images</option>
									<option value="4"<% if pIntSubCategoryView="4" then %> selected<%end if%>>Thumbnails only</option>
								</select>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=427"></a>
							</td>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">The following settings apply when categories are not displayed in a drop-down <br />
(if empty or 0, the default <a href="AdminSettings.asp?tab=3">store-wide setting</a> is used):</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td align="right">Categories per row:</td>
							<td align="left"><input type="text" name="intCategoryColumns" value="<%=pIntCategoryColumns%>" tabindex="302">
							</td>
						</tr>
						<tr> 
							<td align="right">Rows per page:</td>
							<td align="left"> 
							<input type="text" name="intCategoryRows" value="<%=pIntCategoryRows%>" tabindex="303">
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2"><hr></td>
						</tr>	
						
						<%
						' Get list of subcategories	
						
						
						' get data of the category to modify
						query="SELECT idCategory, categoryDesc FROM categories WHERE idParentCategory=" &pIdCategory&";"
						SET rs=Server.CreateObject("ADODB.RecordSet")
						SET rs=conntemp.execute(query)
						
						if err.number <> 0 then
							SET rs=nothing
							
							call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error retreiving subcategories from database") 
						end If
						
						if not rs.EOF then
						%>
						<tr>
							<td colspan="2">Select a sub-category to be displayed as the &quot;featured category&quot;:&nbsp;
								<select name="intFeaturedCategory" size="1" tabindex="304">
									<option value="0" selected="selected">No featured category</option>
									<%
									do while not rs.EOF
										pIdCategory=rs("idCategory")
										pCategoryDesc=replace(rs("categoryDesc"), """", "&quot;")
										pCategoryDesc=replace(pCategoryDesc, "&amp;", "&")
									%>
									<option value="<%=pIdCategory%>" <% if pIntFeaturedCategory = pIdCategory then%>Selected<%end if%>><%=pCategoryDesc%></option>
									<%
									rs.movenext
									loop
									%>
								</select>
							</td>
						</tr>
						<tr>
							<td colspan="2">When displaying the featured subcategory, use the <input type="radio" value="0" name="intFeaturedCategoryImage" <%if pIntFeaturedCategoryImage=0 then%>checked<%end if%> class="clearBorder" tabindex="305">small image <input type="radio" value="1" name="intFeaturedCategoryImage" <%if pIntFeaturedCategoryImage=1 then%>checked<%end if%> class="clearBorder" tabindex="306">large Image</td>
						</tr>
						<tr>
							<td colspan="2"><hr></td>
						</tr>
						<%
						end if
						set rs=nothing
						
						%>
						<tr> 
							<td colspan="2">
							Choose an option for how products are displayed on the category page (if no option is selected, the default <a href="AdminSettings.asp?tab=4">store-wide setting</a> is used):
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td align="right">Display products:</td>
							<td>
								<select name="strPageStyle" tabindex="307">
									<option value="">Default</option>
									<option value="h" <% if pStrPageStyle="h" then %>selected<%end if%>>Horizontally</option>
									<option value="p" <% if pStrPageStyle="p" then %>selected<%end if%>>Vertically</option>
									<option value="l" <% if pStrPageStyle="l" then %>selected<%end if%>>In a list</option>
									<option value="m" <% if pStrPageStyle="m" then %>selected<%end if%>>In a list (multiple Add to Cart)</option>
								</select>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=429"></a>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">The following settings regulate how many products are shown (if empty or 0, the default <a href="AdminSettings.asp?tab=3">store-wide setting</a> is used):</td>
						</tr>
						<tr> 
							<td align="right">Product per row:</td>
							<td align="left"><input type="text" name="intProductColumns" value="<%=pIntProductColumns%>" tabindex="309">
							</td>
						</tr>
						<tr> 
							<td align="right">Rows per page:</td>
							<td align="left"> 
							<input type="text" name="intProductRows" value="<%=pIntProductRows%>" tabindex="310">
							</td>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<th colspan="2">Product Details Page Display Options &nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=424"></a></th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">Choose a display option for the product details page. It will apply to all products within this category. This option <strong>overrides</strong> the <a href="AdminSettings.asp?tab=4">corresponding storewide setting</a>. This is a setting that can also be defined at the product level when adding/editing products.</td>
						</tr>
						<tr> 
							<td colspan="2">  
								<% If ucase(trim(pStrCatDisplayLayout))="C" then  %>
								 <input type="radio" name="CatDisplayLayout" value="C" checked class="clearBorder" tabindex="311">
								<% Else %>
								 <input type="radio" name="CatDisplayLayout" value="C" class="clearBorder" tabindex="311">
								<% End If %>
							 Two-column layout, with product image(s) on the <u>right</u></td>
						</tr>
						<tr> 
							<td colspan="2">  
							<% If ucase(trim(pStrCatDisplayLayout))="L" then  %>
							 <input type="radio" name="CatDisplayLayout" value="L" checked class="clearBorder" tabindex="312">
							<% Else %>
							 <input type="radio" name="CatDisplayLayout" value="L" class="clearBorder" tabindex="312">
							<% End If %>
							Two-column layout, with product image(s) on the <u>left</u></td>
						</tr>
						<tr> 
							<td colspan="2">  
							<% If ucase(trim(pStrCatDisplayLayout))="O" then  %>
								<input type="radio" name="CatDisplayLayout" value="O" checked class="clearBorder" tabindex="313">
							<% Else %>
								<input type="radio" name="CatDisplayLayout" value="O" class="clearBorder" tabindex="313">
							<% End If %>
							One-column layout</td>
						</tr>
						<tr> 
							<td colspan="2">  
							<% If trim(pStrCatDisplayLayout)="" then  %>
								<input type="radio" name="CatDisplayLayout" value="D" checked class="clearBorder" tabindex="314">
							<% Else %>
								<input type="radio" name="CatDisplayLayout" value="D" class="clearBorder" tabindex="314">
							<% End If %>
							Use store's default value</td>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
            <tr>
					    <th colspan="2">Product sorting method within this category:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=432"></a></th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
            <tr>
              <td colspan="2">
                Specify a default ordering for products in this category only. This option <strong>overrides</strong> the <a href="AdminSettings.asp?tab=4">corresponding storewide setting</a>. To specify custom product ordering wtihin this category click <a href="editCategories.asp?nav=&lid=<%=pIdCategory%>" target="_blank">here</a>:
              </td>
            </tr>
            <% 
              optionDisabled = ""
              If pIntNumOrdered > 0 Then 
                optionDisabled = "disabled"
              End If
            %>
						<tr>
							<td colspan="2">
							  <input type="radio" name="strProductOrder" value="" <% If pStrProductOrder="" Then %>checked<% End If %> <%= optionDisabled %> class="clearBorder">
							  Use store's default value
							</td>
						</tr>
						<tr>
							<td colspan="2">
							  <input type="radio" name="strProductOrder" value="0" <% If pStrProductOrder="0" Then %>checked<% End If %> <%= optionDisabled %> class="clearBorder">
							  <%=dictLanguageCP.Item(Session("language")&"_cpSettings_58")%>
							</td>
						</tr>
						<tr>
							<td colspan="2">
							  <input type="radio" name="strProductOrder" value="1" <% If pStrProductOrder="1" Then %>checked<% End If %> <%= optionDisabled %> class="clearBorder">
							  <%=dictLanguageCP.Item(Session("language")&"_cpSettings_59")%>
							</td>
						</tr>
						<tr>
							<td colspan="2">
							  <input type="radio" name="strProductOrder" value="2" <% If pStrProductOrder="2" Then %>checked<% End If %> <%= optionDisabled %> class="clearBorder">
							  <%=dictLanguageCP.Item(Session("language")&"_cpSettings_60")%>&nbsp;
							</td>
						</tr>
						<tr>
							<td colspan="2">
							  <input type="radio" name="strProductOrder" value="3" <% If pStrProductOrder="3" Then %>checked<% End If %> <%= optionDisabled %> class="clearBorder">
							  <%=dictLanguageCP.Item(Session("language")&"_cpSettings_61")%>
							</td>
						</tr>
            <% If pIntNumOrdered > 0 Then %>
						  <tr>
							  <td colspan="2">
							    <input type="radio" name="strProductOrder" value="<%= pStrProductOrder %>" checked class="clearBorder">
							    Custom Product Ordering (click <a href="editCategories.asp?nav=&lid=<%=pIdCategory%>" target="_blank">here</a> to reset).
							  </td>
						  </tr>
            <% End If %>

					</table>
					
				</div>
			<%
			'// =========================================
			'// THIRD PANEL - END
			'// =========================================
			
			'// =========================================
			'// FOURTH PANEL - START - Other settings
			'// =========================================
			%>
				<div id="tabs-4" class="tab-pane">

					<table class="pcCPcontent">	
					
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Other Settings</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2">Restrict the visibility of this category (the products that it contains are also hidden):</td>
						</tr>
						<tr> 
							<td align="right">
								<input type="checkbox" name="iBTOhide" value="1" <%If piBTOhide="1" then%> checked<% end if%> class="clearBorder" tabindex="401">
							</td>
							<td>Hide this category in the storefront</td>
						</tr>
						<tr> 
							<td align="right"><input type="checkbox" name="RetailHide" value="1" <%If pcv_intRetailHide="1" then%> checked<% end if%> class="clearBorder" tabindex="402"></td>
							<td>Hide this category in the storefront from retail customers (wholesale customers can see it)</td>
						</tr>
						<% If ptaxAvalara = 1 Then %>
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <th colspan="2">Avalara Settings</th>
                        </tr>
                        <tr> 
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                        	<td colspan="2">Tax Code : <input name="AvalaraTaxCode" type="text" value="<%=pStrCatAvalaraTaxCode%>"></td>
                        </tr>
                        <% End If %>
					</table>
					
				</div>
			<%
			'// =========================================
			'// FOURTH PANEL - END
			'// =========================================
			
			'// =========================================
			'// FIFTH PANEL - START - Meta Tags
			'// =========================================
			%>
				<div id="tabs-5" class="tab-pane">

					<table class="pcCPcontent">	

						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>  
						<tr>
							<th colspan="2">Category Meta Tags</th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">Enter Meta Tags specific to this category.&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=204"></a></td>
						</tr>
						<tr>
							<td align="right" valign="top">Title</td>
							<td><textarea name="CatMetaTitle" cols="50" rows="2" tabindex="501"><%=pStrCatMetaTitle%></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Description</td>
							<td><textarea name="CatMetaDesc" cols="50" rows="6" tabindex="502"><%=pStrCatMetaDesc%></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Keywords</td>
							<td><textarea name="CatMetaKeywords" cols="50" rows="4" tabindex="503"><%=pStrCatMetaKeywords%></textarea>
						</tr>
					
					</table>
					
				</div>
			<%
			'// =========================================
			'// FIFTH PANEL - END
			'// =========================================



			'// =========================================
			'// SIXTH PANEL - END
			'// =========================================
			'// START - Custom fields
			'if pcv_ProductType<>"item" then
			%>
				<div id="tabs-6" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>  
						<tr>
							<th colspan="2">Custom Search Fields</th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">This tab will allow the store manager to view, add, and edit custom search fields associated with this category.</td>
						</tr>
						<tr>
							<td colspan="2">
								<%
								
								
								query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName FROM pcSearchFields INNER JOIN pcSearchFields_Categories ON pcSearchFields.idSearchField=pcSearchFields_Categories.idSearchData WHERE pcSearchFields_Categories.idCategory=" & pcv_strCSFId &";"
								set rs=connTemp.execute(query)
								tmpJSStr=""
								tmpJSStr=tmpJSStr & "var SFID=new Array();" & vbcrlf
								tmpJSStr=tmpJSStr & "var SFNAME=new Array();" & vbcrlf
								intCount=-1
								IF not rs.eof THEN
									pcArr=rs.getRows()
									set rs=nothing
									intCount=ubound(pcArr,2)
									For i=0 to intCount
										tmpJSStr=tmpJSStr & "SFID[" & i & "]=" & pcArr(0,i) & ";" & vbcrlf
										tmpJSStr=tmpJSStr & "SFNAME[" & i & "]='" & replace(pcArr(1,i),"'","\'") & "';" & vbcrlf
									Next
								END IF
								set rs=nothing
								tmpJSStr=tmpJSStr & "var SFCount=" & intCount & ";" & vbcrlf%>
								<script type=text/javascript>
									<%=tmpJSStr%>
									function CreateTable()
									{
										var tmp1="";
										var tmp2="";
										var i=0;
										var found=0;
										tmp1='<table class="pcCPcontent"><tr><td></td><td nowrap><strong>Text to display</strong></td></tr>';
										for (var i=0;i<=SFCount;i++)
										{
											found=1;
											tmp1=tmp1 + '<tr><td align="right"><a href="javascript:ClearSF(SFID['+i+']);"><img src="../pc/images/minus.jpg" alt="Remove" border="0"></a></td><td width="275" nowrap>'+SFNAME[i]+'</td><td width="100%">&nbsp;</td></tr>';
											if (tmp2=="") tmp2=tmp2 + "||";
											tmp2=tmp2 + "^^^" + SFID[i] + "^^^||"
										}
										tmp1=tmp1+'</table>';
										if (found==0) tmp1="<br><b>No search fields are assigned to this category</b><br><br>";
										document.getElementById("stable").innerHTML=tmp1;
										document.hForm.SFData.value=tmp2;
									}
									function ClearSF(tmpSFID)
									{
										var i=0;
										for (var i=0;i<=SFCount;i++)
										{
											if (SFID[i]==tmpSFID)
											{
												removedArr = SFID.splice(i,1);
												removedArr = SFNAME.splice(i,1);
												SFCount--;
												break;
											}
										}
										CreateTable();
									}
					
									function AddSF(tmpSFID,tmpSFName)
									{
										if (tmpSFID!="")
										{
											var i=0;
											var found=0;
											for (var i=0;i<=SFCount;i++)
											{
												if (SFID[i]==tmpSFID)
												{
													found=1;
													break;
												}
											}
											if (found==0)
											{
												SFCount++;
												SFID[SFCount]=tmpSFID;
												SFNAME[SFCount]=tmpSFName;
											}
											CreateTable();
										}
									}
								</script>
								<span id="stable" name="stable"></span>
								<input type="hidden" name="SFData" value="">
								<%query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields WHERE pcSearchFieldCPShow=1 ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
								set rs=Server.CreateObject("ADODB.Recordset")
								set rs=conntemp.execute(query)
								if not rs.eof then
									pcv_tempFunc=""
									pcv_tempFunc=pcv_tempFunc & "<script type=text/javascript>" & vbcrlf
									pcv_tempFunc=pcv_tempFunc & "function CheckList(cvalue) {" & vbcrlf
									pcv_tempFunc=pcv_tempFunc & "if (cvalue==0) {" & vbcrlf
									pcv_tempFunc=pcv_tempFunc & "var SelectA = document.hForm.SearchValues;" & vbcrlf
									pcv_tempFunc=pcv_tempFunc & "SelectA.options.length = 0; }" & vbcrlf
					
									pcv_tempList=""
									pcv_tempList=pcv_tempList & "<select name=""customfield"" onchange=""CheckList(document.hForm.customfield.value);"">" & vbcrlf
					
									pcArray=rs.getRows()
									intCount=ubound(pcArray,2)
									set rs=nothing
					
									For i=0 to intCount
										pcv_tempList=pcv_tempList & "<option value=""" & pcArray(0,i) & """>" & replace(pcArray(1,i),"""","&quot;") & "</option>" & vbcrlf
									Next
			
									pcv_tempList=pcv_tempList & "</select>" & vbcrlf
									pcv_tempFunc=pcv_tempFunc & "}" & vbcrlf
									pcv_tempFunc=pcv_tempFunc & "</script>" & vbcrlf
									%>
									<br><br>
									<hr>
									<table class="pcCPcontent" style="width:auto;">
										<tr>
											<td colspan="2"><a name="2"></a><b>Add new search field values to this category</b></td>
										</tr>
										<tr>
											<td width="20%" nowrap><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_92")%></td>
											<td width="80%">
											<%=pcv_tempList%>
											<%=pcv_tempFunc%>
											<script type=text/javascript>
												CheckList(document.hForm.customfield.value);
											</script>
											&nbsp;<a href="javascript:AddSF(document.hForm.customfield.value,document.hForm.customfield.options[document.hForm.customfield.selectedIndex].text);"><img src="../pc/images/plus.jpg" alt="Add" border="0"></a>
											</td>
										</tr>
										</table>
								
								
								
								
								<%else
									query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
									set rs=Server.CreateObject("ADODB.Recordset")
									set rs=conntemp.execute(query)
									if not rs.eof then%>
										<a href="ManageSearchFields.asp">Click here</a> to manage custom search fields.</a>
									<%else%>
										<a href="ManageSearchFields.asp">Click here</a> to add new custom search field.</a>
									<%end if
									set rs=nothing%>
								<%end if%>
								<script type=text/javascript>CreateTable();</script>
							</td>
						</tr>
					</table>
					<%  %>
				</div>
				
			<%
			'end if
			'// END - Custom fields
			'// =========================================
			'// SIXTH PANEL - END
			'// =========================================
			%>
            
			
			</div>
			
		<%
		'// TABBED PANELS - MAIN DIV END
		%>
        
        </div>  
    </div>
    <div style="clear: both;">&nbsp;</div>
  <script type=text/javascript>
		$pc(function() {
			$pc( "#TabbedPanels1" ).tab('show')
		});
  </script>

</form>

<!--#include file="AdminFooter.asp"-->
