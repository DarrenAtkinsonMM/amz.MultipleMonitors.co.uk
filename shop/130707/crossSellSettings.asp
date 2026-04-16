<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="products" %>
<%PmAdmin="2*3*"%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
idmain=request.QueryString("idmain")
	if not isNumeric(idmain) or trim(idmain)="" then
		idmain=1
	end if
%>
<% 
	if idmain=1 then
		pageTitle="Cross Selling - General Settings"
	else
		pageTitle="Cross Selling - Product-specific Settings"
	end if
	pageIcon="pcv4_icon_process.gif"
%>
<!--#include file="AdminHeader.asp"-->

<%
' Remove product specific settings
if request("RemoveSettings")<>"" then
	
	idmain=request.Form("idmain") 
	query="DELETE FROM crossSelldata WHERE id="&idmain&";"
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	set rstemp=nothing
	
	call closeDb()
response.redirect "crossSellSettings.asp?idmain="&idmain
	response.end 
end if
' End Remove

' Update Settings
if request("SubmitSettings")<>"" then
	
	idmain=request.Form("idmain")
	csw_status=request.form("csw_status") 
	cs_status=request.Form("cs_status")
	cs_showprod=request.Form("cs_showprod")
	cs_ProductViewCnt=request.Form("cs_ProductViewCnt")
	if cs_ProductViewCnt="" then
		cs_ProductViewCnt=0
	end if
	cs_showcart=request.Form("cs_showcart")
	cs_CartViewCnt=request.Form("cs_CartViewCnt")
	if cs_CartViewCnt="" then
		cs_CartViewCnt=0
	end if
	cs_showimage=request.Form("cs_showimage")
    cs_ImageHeight=request.Form("cs_ImageHeight")
	if not isNumeric(cs_ImageHeight) or trim(cs_ImageHeight)="" then
		cs_ImageHeight=0
	end if
    cs_ImageWidth=request.Form("cs_ImageWidth")
	if not isNumeric(cs_ImageWidth) or trim(cs_ImageWidth)="" then
		cs_ImageWidth=0
	end if
	crossSellText=getUserInput(request.Form("crossSellText"),250)
	cs_showNFS=request.Form("cs_showNFS")
	
	query="UPDATE crossSelldata SET cs_status="&cs_status&", cs_showprod="&cs_showprod&", cs_showcart="&cs_showcart&", cs_showimage="&cs_showimage&", crossSellText=N'"&crossSellText&"',cs_ProductViewCnt="&cs_ProductViewCnt&",cs_CartViewCnt="&cs_CartViewCnt&",cs_ImageHeight="&cs_ImageHeight&",cs_ImageWidth="&cs_ImageWidth&",cs_showNFS="&cs_showNFS&", csw_status="&csw_status&" WHERE id="&idmain&";"
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	set rstemp=nothing

	query="UPDATE crossSelldata SET csw_status="&csw_status&";"
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	set rstemp=nothing	
	call closeDb()
response.redirect "crossSellSettings.asp?idmain="&idmain
	response.end 
end if
' End Update Settings

' Show Settings

query="SELECT cs_status,cs_showprod,cs_showcart,cs_showimage,crossSellText,cs_ProductViewCnt,cs_CartViewCnt,cs_ImageHeight,cs_ImageWidth,cs_showNFS, csw_status FROM crossSelldata WHERE id="&idmain&";"
set rstemp=Server.CreateObject("ADODB.Recordset") 
set rstemp=conntemp.execute(query)
if rstemp.eof then
	set rstemp=nothing
	' There are no product-specific settings -> Add them
	' (1) Load default settings
	query="SELECT cs_status,cs_showprod,cs_showcart,cs_showimage,crossSellText,cs_ProductViewCnt,cs_CartViewCnt,cs_ImageHeight,cs_ImageWidth,cs_showNFS,csw_status FROM crossSelldata WHERE id=1;"
	set rstemp=Server.CreateObject("ADODB.Recordset") 
	set rstemp=conntemp.execute(query)
	cs_status=rstemp("cs_status")
	cs_showprod=rstemp("cs_showprod")
	cs_showcart=rstemp("cs_showcart")
	cs_showimage=rstemp("cs_showimage")
	crossSellText=rstemp("crossSellText")
	cs_ProductViewCnt=rstemp("cs_ProductViewCnt")
	cs_CartViewCnt=rstemp("cs_CartViewCnt")
	cs_ImageHeight=rstemp("cs_ImageHeight")
	cs_ImageWidth=rstemp("cs_ImageWidth")
	cs_showNFS=rstemp("cs_showNFS")
	csw_status=rstemp("csw_status")
	' (2) Create a new record for this product and populate with default settings
	query="INSERT INTO crossSelldata (id, cs_status, cs_showprod, cs_showcart, cs_showimage, crossSellText, cs_ProductViewCnt, cs_CartViewCnt, cs_ImageHeight, cs_ImageWidth,cs_showNFS, csw_status) values ("&idmain&","&cs_status&","&cs_showprod&","&cs_showcart&","&cs_showimage&",N'"&crossSellText&"',"&cs_ProductViewCnt&","&cs_CartViewCnt&","&cs_ImageHeight&","&cs_ImageWidth&","&cs_showNFS&","&csw_status&");"
	set rstemp=conntemp.execute(query)
	set rstemp=nothing
	
else
	' There are product-specific settings -> load them now
	cs_status=rstemp("cs_status")
	cs_showprod=rstemp("cs_showprod")
	cs_showcart=rstemp("cs_showcart")
	cs_showimage=rstemp("cs_showimage")
	crossSellText=rstemp("crossSellText")
	cs_ProductViewCnt=rstemp("cs_ProductViewCnt")
	cs_CartViewCnt=rstemp("cs_CartViewCnt")
	cs_ImageHeight=rstemp("cs_ImageHeight")
	cs_ImageWidth=rstemp("cs_ImageWidth")
	cs_showNFS=rstemp("cs_showNFS")
	csw_status=rstemp("csw_status")
end if
set rstemp=nothing

%>
<form name="form1" method="post" action="crossSellSettings.asp" class="pcForms">
<input name="idmain" type="hidden" value="<%=idmain%>">
	<table class="pcCPcontent">
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr><td><h2>Cross Selling Quick Widget</h2>
		If you do not have time te set up comprehensive cross selling then enable this widget.  This setting determines if the cross selling widget is enabled on your product pages.  If turned on, Products from the same
		category as the current product and products from the same brand will be displayed.
		</td></tr>
		<tr> 
			<td>
				Turn cross selling widgets:&nbsp;
				<input name="csw_status" type="radio" value="-1" <% if csw_status="-1" then %>checked<%end if%> class="clearBorder">On 
				<input name="csw_status" type="radio" value="0" <% if csw_status="0" or isnull(csw_status) then %>checked<%end if%> class="clearBorder">Off
			</td>
		</tr>
		
		<tr>
			<td>
			<% if idmain="1" then%>
				<h2>You are editing general cross selling settings for your store</h2>
				These settings apply to all products <u>except</u> for those products for which you have setup product-specific cross selling settings. To do so, first create a cross selling relationship for that product, then click on the &quot;Edit Product-specific Settings&quot; button on the page where you can view and edit the relationship.
			<% else
					
					query="SELECT description, sku FROM products WHERE idproduct="&idmain&";"
					set rs=Server.CreateObject("ADODB.Recordset") 
					set rs=conntemp.execute(query)
					productName=rs("description")
					productSku=rs("sku")
					set rs = nothing
					
			%>
			 	<h2>You are editing cross selling settings that apply only to <%=productName%> (<%=productSku%>)</h2>
				These settings will overwrite the <a href="crossSellSettings.asp?idmain=1">general cross selling settings</a> used by your store on the product details page (pc/viewprd.asp). They do not affect the View Shopping Cart page (pc/viewcart.asp). <a href="../pc/viewPrd.asp?idproduct=<%=idmain%>&adminPreview=1" target="_blank">View</a> this product in the storefront.
			<% end if %>
			</td>
		</tr>
		<tr> 
			<td>
				Turn cross selling:&nbsp;
				<input name="cs_status" type="radio" value="-1" <% if cs_status="-1" then %>checked<%end if%> class="clearBorder">On 
				<input name="cs_status" type="radio" value="0" <% if cs_status="0" then %>checked<%end if%> class="clearBorder">Off
			</td>
		</tr>
		<tr> 
			<td><hr></td>
		</tr>		

		<tr> 
			<td>Use the settings below to specify <u>where</u> cross-selling relationships should be shown.</td>
		</tr>
		<tr> 
			<td>
				Show related products on product details page:&nbsp;
				<input name="cs_showprod" type="radio" value="-1" <% if cs_showprod="-1" then %>checked<%end if%> class="clearBorder">Yes 
				<input name="cs_showprod" type="radio" value="0" <% if cs_showprod="0" then %>checked<%end if%> class="clearBorder">No
			</td>
		</tr>
		<tr>
			<td style="padding-left: 10px;">
				Number of cross selling products to show on product details page:
				<input name="cs_ProductViewCnt" type="text" value="<%=cs_ProductViewCnt%>" size="4">
			</td>
		</tr>
        <tr> 
			<td>
				Display NFS (Not for Sale) products in cross selling results:&nbsp;
				<input name="cs_showNFS" type="radio" value="-1" <% if cs_showNFS="-1" then %>checked<%end if%> class="clearBorder">Yes 
				<input name="cs_showNFS" type="radio" value="0" <% if cs_showNFS="0" then %>checked<%end if%> class="clearBorder">No
			</td>
		</tr>
		<% 
		' Start - Product specific settings
		' Hide cross selling settings for viewCart.asp when the admin is setting product-specific settings
		if idmain="1" then 
		%>
		<tr> 
			<td>
				Show related products when adding to cart:&nbsp;
				<input name="cs_showcart" type="radio" value="-1" <% if cs_showcart="-1" then%>checked<%end if%> class="clearBorder">Yes 
				<input name="cs_showcart" type="radio" value="0" <% if cs_showcart="0" then%>checked<%end if%> class="clearBorder">No
			</td>
		</tr>
		<tr>
			<td style="padding-left: 10px;">
				Number of results to show on the &quot;view cart&quot; page:&nbsp;
				<input name="cs_CartViewCnt" type="text" value="<%=cs_CartViewCnt%>" size="4">
			</td>
		</tr>
		<tr>
			<td><hr></td>
		</tr>
		<%
		else
		%>
		<tr>
			<td>
				<input name="cs_showcart" type="hidden" value="0">
				<input name="cs_CartViewCnt" type="hidden" value="0">
				<hr>
			</td>
		</tr>
		<%		
		end if
		' End - Product specific settings
		%>
		<tr>
			<td>Use the settings below to specify <u>how</u> they should be presented on the page.</td>
		</tr>
		<tr> 
			<td>
				Show product thumbnails? If no, text links appear.&nbsp;&nbsp;
				<input name="cs_showimage" type="radio" value="-1" <% if cs_showimage="-1" then %>checked<%end if%> class="clearBorder">Yes 
				<input name="cs_showimage" type="radio" value="0" <% if cs_showimage="0" then %>checked<%end if%> class="clearBorder">No
			</td>
		</tr>
		<tr>
			<td style="padding-left: 10px;">
				Thumbnail height:&nbsp;
				<input name="cs_ImageHeight" type="text" value="<%=cs_ImageHeight%>" size="4">
				&nbsp;(enter &quot;0&quot; to preserve the image size)
			</td>
		</tr>
		<tr>
			<td style="padding-left: 10px;">
				Thumbnail width:&nbsp;
				<input name="cs_ImageWidth" type="text" value="<%=cs_ImageWidth%>" size="4">
				&nbsp;(enter &quot;0&quot; to preserve the image size)
			</td>
		</tr>
		<tr> 
			<td class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td>
				Text to show (shown right above the related products - max 255 characters):
				<div style="padding:5px;">
				<textarea name="crossSellText" cols="60" rows="3" wrap="virtual" maxlength="255"><%=crossSellText%></textarea>
				</div>
			</td>
		</tr>
		<tr>
			<td><hr></td>
		</tr>
		<tr> 
			<td class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td align="center">
				<input type="submit" name="SubmitSettings" value="Update" class="btn btn-primary">&nbsp;
				<% if idmain<>"1" then %>
				<input type="submit" name="RemoveSettings" value="Remove Product-specific Settings" onClick="return(confirm('You are about to remove product-specific cross selling settings for this product. The storewide cross selling settings will be used instead. Click OK to confirm the removal or CANCEL to keep the current settings.'));">&nbsp;
				<% end if %>
				<input type="button" class="btn btn-default"  name="back" value="Back" onClick="javascript:history.back()">&nbsp;
			</td>
		</tr>
		<tr> 
			<td align="center">
				<input type="button" class="btn btn-default"  value="Add New Relationship" onClick="location.href='crossSellAdd.asp'">&nbsp;
				<input type="button" class="btn btn-default"  value="View Existing Relationships" onClick="location.href='crossSellView.asp'">
			</td>
		</tr>
		<tr> 
			<td class="pcCPspacer"></td>
		</tr>		
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
