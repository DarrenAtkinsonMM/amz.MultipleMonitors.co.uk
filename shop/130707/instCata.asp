<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" -->
<%
dim f
%>
<% pageTitle=dictLanguageCP.Item(Session("language")&"_cpInstCat_0") %>
<% section="products" %>
<!--#include file="AdminHeader.asp"-->

<!-- #include file="../htmleditor/editor.asp" -->

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

</script>

<form action="instCatb.asp" method="post" name="hForm" onSubmit="return Form1_Validator(this)" class="pcForms">
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
				<li>
					<div style="height:40px; margin-top:10px; text-align: center">
					<input type="hidden" name="reqstr" value="<%=request.QueryString("reqstr")%>">
					<input name="Submit" type="submit" value="Add" class="btn btn-primary">
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
							<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_1")%></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td width="20%" align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_157")%>:</td>
							<td width="80%"><input name="categoryDesc" type="text" value="" size="40" tabindex="101"></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_2")%>:</td>
							<td>
								<input type="text" name="image" value="" size="40" tabindex="102"><a href="#" onClick="chgWin('../pc/imageDir.asp?ffid=image&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=439"></a>
							</td>
						</tr>
						<tr> 
							<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_3")%>:</td>
							<td> 
				        <input type="text" name="largeimage" value="<%=plargeImage%>" size="40" tabindex="103"><a href="#" onClick="chgWin('../pc/imageDir.asp?ffid=largeimage&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=439"></a>
							</td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td><p> 
								<!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->
								<%If HaveImgUplResizeObjs=1 then%>
								<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_8")%>&nbsp;<a href="#" onClick="window.open('uploadresize/catResizea.asp','popup','toolbar=no,status=no,location=no,menubar=no,height=350,width=400,scrollbars=no'); return false;">click here</a>.
								<% Else %>
									<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_9")%>&nbsp;<a href="#" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">click here</a>.
								<% End If %>
							</td>
						</tr>
						<tr> 
							<td>&nbsp;</td>
							<td>
								<input type="checkbox" name="NotImg" value="1" class="clearBorder"> Hide category images, except when using "Thumbnails only" under display settings.  
							</td>
						</tr>
						<tr>
							<td colspan="2"><hr></td>
						</tr> 
					<tr> 
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_158")%>:</td>
						<td>
							<%
							cat_DropDownName="idParentCategory"
							cat_Type="0"
							cat_DropDownSize="1"
							cat_MultiSelect="0"
							cat_ExcBTOHide="0"
							cat_StoreFront="0"
							cat_ShowParent="1"
							cat_DefaultItem=""
							cat_SelectedItems="1,"
							cat_ExcItems=""
						
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
							<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_5")%>:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=440"></a></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr valign="top">
							<td><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_7")%>:
								<div class="small"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_8")%></div>
							</td>			
							<td>
								<textarea name="SDesc" id="SDesc" cols="50" rows="6" tabindex="201" maxlength="255"></textarea>
								<div class="pcSmallText">Maximum Length: 255 Characters</div>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr valign="top">
							<td><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_9")%>:
							<div class="small"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_10")%></div>
							</td>
							<td>
							  <textarea class="htmleditor" name="LDesc" id="LDesc" cols="50" rows="6" tabindex="202"></textarea>			
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td align="right">
							<input type="checkbox" name="HideDesc" value="1" class="clearBorder" tabindex="203">
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
						<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_11")%></th>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr> 
						<td colspan="2">
						<%=dictLanguageCP.Item(Session("language")&"_cpInstCat_12")%><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_13")%><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_14")%>
			
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td align="right" width="30%">Display Subcategories:</td>
						<td width="70%">
							<select name="intSubCategoryView" tabindex="301">
								<option value="3"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_315")%></option>
								<option value="2"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_507")%></option>
								<option value="0"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_506")%></option>
								<option value="1"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_505")%></option>
								<option value="4">Thumbnails only</option>
							</select>
							&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=427"></a>
						</td>
					</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
					<tr>
						<td colspan="2">The following settings apply when categories are not displayed in a drop-down (if empty or 0, the default <a href="AdminSettings.asp?tab=3">store-wide setting</a> is used):</td>
					</tr>
					<tr> 
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_508")%>:</td>
						<td align="left"><input type="text" name="intCategoryColumns" value="<%=intCategoryColumns%>" tabindex="302">
						</td>
					</tr>
					<tr> 
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_509")%>:</td>
						<td align="left"> 
						<input type="text" name="intCategoryRows" value="<%=intCategoryRows%>" tabindex="302">
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td colspan="2"><hr></td>
					</tr>
					<tr> 
						<td colspan="2">
						<%=dictLanguageCP.Item(Session("language")&"_cpInstCat_17")%><a href="editCategories.asp?nav=&lid=<%=pIdCategory%>" target="_blank"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_18")%></a><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_19")%>:
						</td>
					</tr>
					<tr>
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_15")%>:</td>
						<td>
							<select name="strPageStyle" tabindex="303">
								<option value=""><%=dictLanguageCP.Item(Session("language")&"_cpCommon_315")%></option>
								<option value="h"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_517")%></option>
								<option value="p"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_518")%></option>
								<option value="l"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_519")%></option>
								<option value="m"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_520")%></option>
							</select>
							&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=429"></a>
						</td>
					</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
					<tr>
						<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_16")%>:</td>
					</tr>
					<tr> 
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_514")%>:</td>
						<td align="left"><input type="text" name="intProductColumns" value="" tabindex="304">
						</td>
					</tr>
					<tr> 
						<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_509")%>:</td>
						<td align="left"> 
						<input type="text" name="intProductRows" value="" tabindex="304">
						</td>
					</tr>
					<tr>
						<td class="pcCPspacer" colspan="2" style="height: 25px;"></td>
					</tr>  
					<tr>
						<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstCat_23")%>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=424"></a></th>
					</tr>
					<tr>
						<td class="pcCPspacer" colspan="2"></td>
					</tr>
					<tr>
						<td colspan="2">Choose a display option for the product details page. It will apply to all products within this category. This option <strong>overrides</strong> the <a href="AdminSettings.asp?tab=3">corresponding storewide setting</a>. This is a setting that can also be defined at the product level when adding/editing products.</td>
					</tr>
					<tr> 
						<td colspan="2">  
						 <input type="radio" name="CatDisplayLayout" value="C" class="clearBorder" tabindex="305"> <%=dictLanguageCP.Item(Session("language")&"_cpCommon_502")%></td>
					</tr>
					<tr> 
						<td colspan="2">  
						 <input type="radio" name="CatDisplayLayout" value="L" class="clearBorder" tabindex="306"> <%=dictLanguageCP.Item(Session("language")&"_cpCommon_503")%></td>
					</tr>
					<tr> 
						<td colspan="2">  
						<input type="radio" name="CatDisplayLayout" value="O" class="clearBorder" tabindex="307"> <%=dictLanguageCP.Item(Session("language")&"_cpCommon_504")%></td>
					</tr>
                    <tr> 
                        <td colspan="2">  
                        <input type="radio" name="CatDisplayLayout" value="D" checked class="clearBorder" tabindex="312"> Use store's default value</td>
                    </tr><tr>
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
								<input type="checkbox" name="iBTOhide" value="1" class="clearBorder" tabindex="401">
							</td>
							<td>Hide this category in the storefront</td>
						</tr>
						<tr> 
							<td align="right"><input type="checkbox" name="RetailHide" value="1" class="clearBorder" tabindex="402"></td>
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
                        	<td colspan="2">Tax Code : <input name="AvalaraTaxCode" type="text" value=""></td>
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
							<td><textarea name="CatMetaTitle" cols="50" rows="2" tabindex="501"></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Description</td>
							<td><textarea name="CatMetaDesc" cols="50" rows="6" tabindex="502"></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Keywords</td>
							<td><textarea name="CatMetaKeywords" cols="50" rows="4" tabindex="503"></textarea>
						</tr>
					
					</table>
					
				</div>
			<%
			'// =========================================
			'// FIFTH PANEL - END
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
