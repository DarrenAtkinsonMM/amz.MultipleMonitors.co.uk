<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" -->
<%
dim pcIntBrandID, pcvBrandsDescription, pcvBrandsSDescription, pcIntBrandsActive, pcIntSubBrandsView, pcvProductsView, pcIntBrandsParent, pcvBrandsMetaTitle, pcvBrandsMetaDesc, pcvBrandsMetaKeywords, pcvBrandsBrandLogoLg

	pcIntBrandID=request("idbrand")
	if not validNum(pcIntBrandID) then 
        call closeDb()
        response.redirect "techErr.asp?error="& Server.Urlencode("Not a valid brand ID.") 
	end if
'// Load data from Existing Brand - START

	
	query="SELECT BrandName, BrandLogo, pcBrands_Description, pcBrands_SDescription, pcBrands_SubBrandsView, pcBrands_ProductsView, pcBrands_Active, pcBrands_Parent, pcBrands_MetaTitle, pcBrands_MetaDesc, pcBrands_MetaKeywords, pcBrands_BrandLogoLg FROM Brands WHERE idbrand=" & pcIntBrandID
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error loading data Brands table with brand ID " & pcIntBrandID) 
	end if

	BrandName=pcf_PrintCharacters(rs("BrandName"))
	BrandLogo=rs("BrandLogo")
	pcvBrandsDescription=pcf_PrintCharacters(rs("pcBrands_Description"))
	pcvBrandsSDescription=pcf_PrintCharacters(rs("pcBrands_SDescription"))
	pcIntSubBrandsView=rs("pcBrands_SubBrandsView")
	pcvProductsView=rs("pcBrands_ProductsView")
	pcIntBrandsActive=rs("pcBrands_Active")
	pcIntBrandsParent=rs("pcBrands_Parent")
	pcvBrandsMetaTitle=rs("pcBrands_MetaTitle")
	pcvBrandsMetaDesc=rs("pcBrands_MetaDesc")
	pcvBrandsMetaKeywords=rs("pcBrands_MetaKeywords")
	pcvBrandsBrandLogoLg=rs("pcBrands_BrandLogoLg")

	set rs=nothing
	
	
	if not validNum(pcIntSubBrandsView) then pcIntSubBrandsView=0
	if not validNum(pcIntBrandsActive) then pcIntBrandsActive=1
	if not validNum(pcIntBrandsParent) then pcIntBrandsParent=0

'// Load data from Existing Brand - END

'// Update Existing Brand - START
if request("action")="update" then
	BrandName=pcf_SanitizeApostrophe(request.form("BrandName"))
	BrandLogo=request.form("image")
	pcvBrandsBrandLogoLg=request.form("largeimage")
	pcvBrandsDescription=pcf_SanitizeApostrophe(request.form("pcBrandsDescription"))
	pcvBrandsSDescription=pcf_SanitizeApostrophe(request.form("pcBrandsSDescription"))
	pcIntSubBrandsView=request.form("intSubBrandsView")
	pcvProductsView=request.form("pcProductsView")
	pcIntBrandsActive=request.form("pcBrandsActive")
	pcIntBrandsParent=request.form("pcBrandsParent")
	pcvBrandsMetaTitle=getUserInput(request.form("pcBrandsMetaTitle"),0)
	pcvBrandsMetaDesc=getUserInput(request.form("pcBrandsMetaDesc"),0)
	pcvBrandsMetaKeywords=getUserInput(request.form("pcBrandsMetaKeywords"),0)
	
	if not validNum(pcIntSubBrandsView) then pcIntSubBrandsView=0
	if not validNum(pcIntBrandsActive) then pcIntBrandsActive=1
	if not validNum(pcIntBrandsParent) then pcIntBrandsParent=0
	
	
	query="UPDATE Brands SET BrandName=N'" & BrandName & "', BrandLogo='" & BrandLogo & "', pcBrands_Description=N'" & pcvBrandsDescription & "', pcBrands_SDescription=N'" & pcvBrandsSDescription& "', pcBrands_SubBrandsView=" & pcIntSubBrandsView & ", pcBrands_ProductsView='" & pcvProductsView& "', pcBrands_Active=" & pcIntBrandsActive & ", pcBrands_Parent=" & pcIntBrandsParent & ", pcBrands_MetaTitle=N'" & pcvBrandsMetaTitle & "', pcBrands_MetaDesc=N'" & pcvBrandsMetaDesc & "', pcBrands_MetaKeywords=N'" & pcvBrandsMetaKeywords & "', pcBrands_BrandLogoLg='" & pcvBrandsBrandLogoLg & "' WHERE idbrand=" & pcIntBrandID
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		Session("message") = "Error updating brand ID " & pcIntBrandID
    response.redirect "msgb.asp"
	else
		set rs=nothing
		
		call closeDb()
        response.redirect "BrandsEdit.asp?s=1&idbrand=" & pcIntBrandID &"&msg="&Server.URLEncode("Brand updated successfully.")
	end if
end if
'// Update Existing Brand - END

'// Show Add New Brand Page
pageTitle="Edit Brand: " & BrandName %>
<!--#include file="AdminHeader.asp"-->

<!-- #include file="../htmleditor/editor.asp" -->

<!--#include file="pcv4_showMessage.asp"-->

<script type=text/javascript>
	function newWindow(file,window) {
			msgWindow=open(file,window,'resizable=no,width=400,height=500');
			if (msgWindow.opener == null) msgWindow.opener = self;
	}

	function Form1_Validator(theForm)
	{
		// InnovaStudio HTML Editor Workaround for this keyword
		theForm = document.hForm;

		if (theForm.BrandName.value == "")
			{
				 alert("Please enter a name for the brand.");
					theForm.BrandName.focus();
					return (false);
			}
	return (true);
	}
</script> 

	<form action="BrandsEdit.asp?action=update" method="post" name="hForm" onSubmit="return Form1_Validator(this)" class="pcForms">
		<div class="cpOtherLinks" style="margin: 0 12px 0 7px;">
			<a href="../pc/viewBrands.asp?idBrand=<%=pcIntBrandID%>" target="_blank">View in the storefront</a>
			&nbsp;|&nbsp;
			<a href="BrandsProducts.asp?idBrand=<%=pcIntBrandID%>" target="_blank">View/Add Products</a>
			&nbsp;|&nbsp;
			<a href="BrandsManage.asp?parent=<%=pcIntBrandID%>" target="_blank">Sub-Brands</a>
      &nbsp;|&nbsp;
      <a href="BrandsManage.asp">Manage All</a>
		</div>

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
				<li><a href="#tabs-3" data-toggle="tab">Display &amp; Other Settings</a></li>				
				<li><a href="#tabs-4" data-toggle="tab">Meta Tags</a></li>
				<li>
					<div style="margin-top:10px; margin-bottom:10px; text-align: center">
                	<input type="hidden" name="idbrand" value="<%=pcIntBrandID%>">
					<input name="Submit" type="submit" value="Update" class="btn btn-primary"><br />
                    <div style="margin-top: 5px"><input type="button" class="btn btn-default"  value="Manage Brands" onClick="document.location.href='BrandsManage.asp';"></div>
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
							<th colspan="2">Brand Name, Images &amp; Parent (if any)</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td width="20%" align="right">Brand Name:</td>
							<td width="80%"><input name="BrandName" id="brandName" type="text" value="<%=BrandName%>" size="40" tabindex="101"></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td align="right">Small Brand Logo:</td>
							<td>
								<input type="text" name="image" value="<%=BrandLogo%>" size="40" tabindex="102"><a href="#" onClick="chgWin('../pc/imageDir.asp?ffid=image&fid=hForm','window2')"><img src="images/search.gif" alt="Locate previously uploaded images" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=478"></a>
							</td>
						</tr>
						<tr> 
							<td align="right">Large Brand Logo:</td>
							<td> 
				        		<input type="text" name="largeimage" value="<%=pcvBrandsBrandLogoLg%>" size="40" tabindex="103"><a href="#" onClick="chgWin('../pc/imageDir.asp?ffid=largeimage&fid=hForm','window2')"><img src="images/search.gif" alt="Locate previously uploaded images" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=478"></a>
							</td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td>
								<!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->
								<%If HaveImgUplResizeObjs=1 then%>
								<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_8")%>&nbsp;<a href="#" onClick="window.open('uploadresize/catResizea.asp','popup','toolbar=no,status=no,location=no,menubar=no,height=350,width=400,scrollbars=no'); return false;">click here</a>.
								<% Else %>
									<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_9")%>&nbsp;<a href="#" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">click here</a>.
								<% End If %>
							</td>
						</tr>
						<tr>
							<td colspan="2"><hr></td>
						</tr> 
					<tr> 
						<td align="right" valign="top" nowrap>Parent Brand:</td>
						<td>
                        	<%
							
								Dim pcBrandsParentExist
								query="SELECT idbrand, BrandName, (SELECT b.BrandName FROM Brands b WHERE b.idBrand = pBrand.pcBrands_Parent) AS ParentBrandName FROM Brands pBrand WHERE idBrand <> " & pcIntBrandID & " ORDER BY BrandName ASC"
								set rs=Server.CreateObject("ADODB.Recordset")
								set rs=connTemp.execute(query)
								if rs.EOF then
									pcBrandsParentExist=0
								else
									pcBrandsParentExist=1
									pcBrandsArr=rs.getRows()
								end if
								set rs=nothing
							
							if pcBrandsParentExist=0 then
							%>
                                No brands available.
                                <br />
                                First add a brand, then you can use it as a &quot;Parent&quot; of another brand.
                            <%
							else
							%>
                            	<select name="pcBrandsParent" tabindex="104">
                                	<option value="0">None</option>
                            <%
                                intCount=ubound(pcBrandsArr,2)
                                For m=0 to intCount %>
									<option value="<%=pcBrandsArr(0,m)%>"<% if pcBrandsArr(0,m)=pcIntBrandsParent then %>selected<% end if %>><%=pcBrandsArr(1,m)%> <% If Len(pcBrandsArr(2,m)) > 0 Then Response.Write " [" & pcBrandsArr(2,m) & "]" %></option>
                            <%
                                Next
                            %>
								</select>
                            <%
							end if
							%>
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
							<th colspan="2">Descriptions:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=479"></a></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr valign="top">
							<td>
								Short Description:
								<div class="small">Shown on pages that display brands</div>
							</td>			
							<td>
								<textarea name="pcBrandsSDescription" id="pcBrandsSDescription" cols="50" rows="6" tabindex="201" maxlength="255"><%=pcvBrandsSDescription%></textarea>
								<div class="pcSmallText">Maximum Length: 255 Characters</div>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr valign="top">
							<td>
								Long Description:
								<div class="small">Only shown on the page that displays products within the brand</div>
							</td>
							<td>
							  <textarea class="htmleditor" name="pcBrandsDescription" id="pcBrandsDescription" cols="50" rows="6" tabindex="202"><%=pcvBrandsDescription%></textarea>
							</td>
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
                            <th colspan="2">Display &amp; Other Settings</th>
                        </tr>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                            <td width="20%" valign="top" nowrap>Display Sub-brands:</td>
                            <td>															
																<select name="intSubBrandsView" tabindex="301">
																	<option value="3"<% if pcIntSubBrandsView="3" then %> selected<%end if%>>Default (like categories)</option>
																	<option value="0"<% if pcIntSubBrandsView="0" then %> selected<%end if%>>In a list, with images</option>
																	<option value="1"<% if pcIntSubBrandsView="1" then %> selected<%end if%>>In a list, without images</option>
																	<option value="2"<% if pcIntSubBrandsView="2" then %> selected<%end if%>>In a drop-down</option>
																	<option value="4"<% if pcIntSubBrandsView="4" then %> selected<%end if%>>Thumbnails only</option>
																</select>

                                &nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=427"></a>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top">Display Products:</td>
                            <td>
                                <select name="pcProductsView" tabindex="302">
                                    <option value=""<% if pcvProductsView="" or isNull(pcvProductsView) then %> selected<% end if %>><%=dictLanguageCP.Item(Session("language")&"_cpCommon_315")%></option>
                                    <option value="h"<% if pcvProductsView="h" then %> selected<% end if %>><%=dictLanguageCP.Item(Session("language")&"_cpCommon_517")%></option>
                                    <option value="p"<% if pcvProductsView="p" then %> selected<% end if %>><%=dictLanguageCP.Item(Session("language")&"_cpCommon_518")%></option>
                                    <option value="l"<% if pcvProductsView="l" then %> selected<% end if %>><%=dictLanguageCP.Item(Session("language")&"_cpCommon_519")%></option>
                                    <option value="m"<% if pcvProductsView="m" then %> selected<% end if %>><%=dictLanguageCP.Item(Session("language")&"_cpCommon_520")%></option>
                                </select>
                                &nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=427"></a>
                            </td>
                        </tr>
                        <tr>
                            <td>Active:</td>
                            <td><input type="radio" name="pcBrandsActive" value="1" class="clearBorder" <% if pcIntBrandsActive="1" then %>checked="checked" <% end if %>tabindex="303"> Yes <input type="radio" name="pcBrandsActive" value="0" class="clearBorder" <% if pcIntBrandsActive="0" then %>checked="checked" <% end if %>tabindex="303"> No</td>
                        </tr>
					</table>
					
				</div>
			<%
			'// =========================================
			'// THIRD PANEL - END
			'// =========================================

			'// =========================================
			'// FOURTH PANEL - START - Meta Tags
			'// =========================================
			%>
				<div id="tabs-4" class="tab-pane">

					<table class="pcCPcontent">	

						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>  
						<tr>
							<th colspan="2">Brand Meta Tags</th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">Enter Meta Tags specific to this brand.&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=204"></a></td>
						</tr>
						<tr>
							<td align="right" valign="top">Title</td>
							<td><textarea name="pcBrandsMetaTitle" cols="50" rows="2" tabindex="401"><%=pcvBrandsMetaTitle%></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Description</td>
							<td><textarea name="pcBrandsMetaDesc" cols="50" rows="6" tabindex="402"><%=pcvBrandsMetaDesc%></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Keywords</td>
							<td><textarea name="pcBrandsMetaKeywords" cols="50" rows="4" tabindex="403"><%=pcvBrandsMetaKeywords%></textarea>
						</tr>
					
					</table>
					
				</div>
			<%
			'// =========================================
			'// FOURTH PANEL - END
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
