<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/SocialNetworkWidgetConstants.asp"-->
<% pageTitle="Generate E-Commerce Widget for Blogs" %>
<% section="layout" %>
<%
Dim pcv_strPageName, widgetURL
pcv_strPageName="genSocialNetworkWidget.asp"

msg=Request("msg")

if scSSL = "1" then
	widgetURL = scSslURL &"/"& scPcFolder & "/pc/"
else
	widgetURL = scStoreURL &"/"& scPcFolder & "/pc/"
end if

pcv_Path=replace(widgetURL, "//", "/")
pcv_Path=replace(pcv_Path,"http:/","http://")
pcv_Path=replace(pcv_Path,"https:/","https://")

If request("action")="add" Then

	Session("SNW_TYPE")=Request("exportType")
	Session("SNW_CATEGORY")=Request("catlist")
	tmpPrdCount=Request("prdcount")
	if tmpPrdCount="" OR tmpPrdCount="0" then
		tmpPrdCount="100"
	end if
	Session("SNW_MAX")=tmpPrdCount
	Session("SNW_AFFILIATE")=Request("affiliate")
	
	wdW=getUserInput(Request("wdw"), 8)
	if wdW="" OR wdW="0" then
		wdW=198
	End If
	If not validNum(wdW) then
		wdW=198
	End If
	
	wdH=getUserInput(Request("wdH"), 8)
	if wdH="" OR wdH="0" then
		wdH=438
	End If
	If not validNum(wdH) then
		wdH=438
	End If
	Session("SNW_WDW")=wdW
	Session("SNW_WDH")=wdH
	
	If Session("SNW_CATEGORY")="" Then
		call closeDb()
        response.redirect(pcv_strPageName&"?msg=You must select a category.")  
	End If	
	
	If Session("SNW_TYPE")="0" Then
		
		Dim pcv_strWidgetFlag, pcv_strWidgetError
		pcv_strWidgetFlag=""
		pcv_strWidgetError=""
		
		'// Generate Static XML File
		' call pcs_SaveSocialWidgetXML()  
		
		'// Generate Widget
		call pcs_SaveSocialWidgetJS()  

		'// Redirect with Message
		If pcv_strWidgetFlag<>"Success" Then
		'	>> Fail
			call closeDb()
            response.redirect(pcv_strPageName&"?msg=There was an error processing your request:" & pcv_strWidgetError)  
		Else	
		'	>> Success
			call closeDb()
            response.redirect("../includes/PageCreateSocialNetworkWidget.asp")
		End If
		
	Else

		'// Dynamic XML >>> Generate Constants		
		call closeDb()
        response.redirect("../includes/PageCreateSocialNetworkWidget.asp")
	
	End If

End If


Sub pcs_SaveSocialWidgetXML()
	
	Dim StringBuilderObj,pcv_intProductCount
	
	pcv_intProductCount=-1
	
	'// Category ID
	pIdCategory=Session("SNW_CATEGORY")
	
	'// Sort
	if ProdSort="" then
		ProdSort="19"
	end if
	
	select case ProdSort
		Case "19": query1 = " ORDER BY categories_products.POrder Asc, products.description Asc"
		Case "0": query1 = " ORDER BY products.SKU Asc"
		Case "1": query1 = " ORDER BY products.description Asc" 	
		Case "2": 
		If Session("customerType")=1 then
			query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) DESC"
		else
			query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) DESC"
		End if
	Case "3":
		If Session("customerType")=1 then
			query1 = " ORDER BY (CASE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE Products.bToBPrice WHEN 0 THEN Products.Price ELSE Products.bToBPrice END) ELSE (CASE (CASE WHEN Products.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultWPrice END) WHEN 0 THEN Products.pcProd_BTODefaultPrice ELSE Products.pcProd_BTODefaultWPrice END) END) ASC"
		else
			query1 = " ORDER BY (CASE (CASE WHEN Products.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE Products.pcProd_BTODefaultPrice END) WHEN 0 THEN Products.Price ELSE Products.pcProd_BTODefaultPrice END) ASC"
		End if 	 	
	end select
	
	'// Query Products
	query="SELECT TOP " & Session("SNW_MAX") & " products.idProduct, products.price, products.smallImageUrl, products.description FROM products, categories_products WHERE products.idProduct=categories_products.idProduct AND categories_products.idCategory="& pIdCategory &" AND active=-1 AND configOnly=0 and removed=0 " & query1
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)		
	if NOT rs.EOF then
		pcArray_Products = rs.getRows()
		pcv_intProductCount = UBound(pcArray_Products,2)
	end If	
	set rs=nothing
	
	'// Set Array
	set StringBuilderObj = new StringBuilder
	StringBuilderObj.append "var wPL = new Array();"

	pCnt=0
	if pcv_intProductCount=-1 then
	'	>> Fail
		call closeDb()
        response.redirect(pcv_strPageName&"?msg=The category you selected is empty.")
	end if
	
	For pCnt=0 to pcv_intProductCount
	
		pidProduct=""
		pDescription=""   
		pPrice=""
		pSmallImageUrl="" 
	
		pidProduct=pcArray_Products(0,pCnt) '// rs("idProduct")
		pPrice=pcArray_Products(1,pCnt) '// rs("price")
		pSmallImageUrl=pcArray_Products(2,pCnt) '// rs("smallImageUrl")
		pDescription=pcArray_Products(3,pCnt) '// rs("description") 
								
		if pSmallImageUrl="" OR isNULL(pSmallImageUrl) then
			pSmallImageUrl="no_image.gif"
		end if		

		pDescription=ClearHTMLTags2(pDescription,2)
		pDescription=replace(pDescription,"&quot;","""")
		If 44<len(pDescription) then
			pDescription=trim(left(pDescription,44)) & "..."
		End If
	
		StringBuilderObj.append "wPL[" & pCnt & "]=""" & replace(pDescription & "||" & money(pPrice) & "||" & pcv_Path&"catalog/"&pSmallImageUrl& "||" &pcv_Path&"viewPrd.asp?idproduct="&pidProduct,"""","&quot;") & """;" & vbcrlf
	Next
	
	'// Set Path
	if PPD="1" then
		pcStrFolder=Server.Mappath ("/"&scPcFolder&"/pc")
	else
		pcStrFolder=server.MapPath("../pc")
	end if

	'// Write XML File
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fs.CreateTextFile(pcStrFolder & "\pcSyndication.inc",True)
	a.Write(StringBuilderObj.toString())
	a.Close
	Set a=Nothing
	Set fs=Nothing
	
	If err.description="" Then
		pcv_strWidgetFlag="Success"
		pcv_strWidgetError=""
	Else
		pcv_strWidgetFlag=""
		pcv_strWidgetError=err.description
	End If 

End Sub


Sub pcs_SaveSocialWidgetJS()
	On Error Resume Next
	
	Dim pcv_strWidgetXML, pcv_NewLine, pcv_URL, pcv_WidgetExists
	pcv_WidgetExists=""
	pcv_URL=""
	pcv_NewLine=CHR(10)	
	pcv_strWidgetXML = ""
	pcv_strWidgetXML=pcv_strWidgetXML&""
	
	'// Set URL
	pcv_URL=pcv_Path&"pcSyndication_ShowItems.asp"
	
	Dim wdW, wdH
	wdW=getUserInput(Request("wdw"), 8)
	If not validNum(wdW) then
		wdW=198
	End If
	
	wdH=getUserInput(Request("wdH"), 8)
	If not validNum(wdH) then
		wdH=438
	End If	
	
	'// Set JS File
	pcv_strWidgetXML=pcv_strWidgetXML&"//Change 'idaffiliate' value with an AffiliateID"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var idaffiliate=0;"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_ad_width = " & wdW & ";"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_ad_height = " & wdH & ";"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_style_border = ""0px solid #FFFFFF"";"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_style_background = ""#FFFFFF"";"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_ad_frame = ""0"";"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_ad_src = """& pcv_URL & "?mh=""  + pcv_ad_height + ""&mw=""  + pcv_ad_width + ""&idaffiliate="" + idaffiliate" & pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_window=window;"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"var pcv_doc=document;"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"pcf_ShowItems(pcv_doc,pcv_window);"& pcv_NewLine

	pcv_strWidgetXML=pcv_strWidgetXML&"function pcf_ShowItems(pcv_window,pcv_doc) {"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('<style type=""text/css"">');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('.pcProductsFrame {');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('border: '+pcv_style_border+';');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('background-color: '+pcv_style_background+';');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('}');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('</style>');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('<iframe name=""pcv_ProductsFrame"" width=""'+pcv_ad_width+'"" class=""pcProductsFrame"" height=""'+pcv_ad_height+'"" frameborder=""'+pcv_ad_frame+'"" src=""'+pcv_ad_src+'"" marginwidth=""0"" marginheight=""0"" vspace=""0"" hspace=""0"" allowtransparency=""false"" scrolling=""no"">');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"	pcv_window.write('</iframe>');"& pcv_NewLine
	pcv_strWidgetXML=pcv_strWidgetXML&"}"& pcv_NewLine
	
	'response.Write(pcv_strWidgetXML)
	'response.End()

	'// Set Path
	if PPD="1" then
		pcStrFolder=Server.Mappath ("/"&scPcFolder&"/pc")
	else
		pcStrFolder=server.MapPath("../pc")
	end if	

	'// Write JS File
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set a=fs.CreateTextFile(pcStrFolder & "\pcSyndication.js",True)
	a.Write(pcv_strWidgetXML)
	a.Close
	Set a=Nothing
	Set fs=Nothing

	If err.description="" Then
		pcv_strWidgetFlag="Success"
		pcv_strWidgetError=""
	Else
		pcv_strWidgetFlag=""
		pcv_strWidgetError=err.description
	End If 

End Sub
%>
<!--#include file="AdminHeader.asp"-->
<style>
.pcCPOverview {
	background-color: #F5F5F5;
	border: 1px solid #FF9900;
	margin: 5px;
	padding: 5px;
	color: #666666;
	font-size:11px;
	text-align: left;
}
.pcCodeStyle {
	font-family: "Courier New", Courier, monospace;
	color: #FF0000;
	font-size: 9;
}
</style>
<% If msg="success" Then %>

<%
Session("SNW_TYPE")=""
Session("SNW_CATEGORY")=""
Session("SNW_MAX")=""
Session("SNW_AFFILIATE")=""
%>
<table class="pcCPcontent">
<tr>
	<td align="center">
		<div align="center">
			<strong>ProductCart E-Commerce Widget created successfully!</strong>
			<br />
      <br />
         	<div class="pcCPOverview">
            	Add the following snippet of JavaScript code to your blog or other page that supports JavaScript to display your widget (<a href="http://wiki.productcart.com/widgets/productcart_ecommerce_widget" target="_blank">User Guide</a>):
                <br />
                <br />
              	<span class="pcCodeStyle">
                	<%
					pcv_URL=pcv_Path&"pcSyndication.js"
					%>
					&lt;script type=&quot;text/javascript&quot; src=&quot;<%=pcv_URL%>&quot;&gt; &lt;/script&gt;            
              	</span>
              	<br />
       	</div>
			<br />
            <br />
			<a href="../pc/pcSyndication_Preview.asp?path=<%=pcv_URL%>" target="_blank">Preview Widget</a> | <a href="genSocialNetworkWidget.asp">Generate New E-Commerce Widget</a>
		</div>
	</td>
</tr>
<tr>
	<td class="pcCPspacer"></td>
</tr>
</table>

<% Else %>

<form method="post" name="form1" action="<%=pcv_strPageName%>?action=add" class="pcForms">
	<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2"> How it works</th>
	</tr>
	<%if msg<>"" then%>
	<tr>
		<td colspan="2">
			<div class="pcCPmessage">
				<%=msg%>
			</div>
		</td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">
			<p>The <strong>ProductCart E-Commerce Widget for Blogs</strong> allows you to take products you have for sale on your store, and show them on a blog or other Web page including popular social networks. The Widget can be styled to match the look and feel of any web site. Read the <a href="http://wiki.productcart.com/widgets/productcart_ecommerce_widget" target="_blank">ProductCart E-Commerce Widget User Guide</a> for more information.</p>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Creating or Updating your Widget</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td align="right">Widget Width:</td>
		<td nowrap="nowrap"><input type="text" name="wdw" size="3" value="<%if SNW_WDW<>"" then%><%=SNW_WDW%><%else%>198<%end if%>"> px</td>
	</tr>
	<tr>
		<td align="right">Widget Height:</td>
		<td nowrap="nowrap"><input type="text" name="wdh" size="3" value="<%if SNW_WDH<>"" then%><%=SNW_WDH%><%else%>438<%end if%>"> px</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2"><p>Show Affiliate Links: 
		    <input type="radio" name="affiliate" value="1" class="clearBorder" <% if SNW_AFFILIATE="1" then response.Write("checked") %>> 
		    Yes&nbsp;&nbsp;
		    <input type="radio" name="affiliate" value="0" class="clearBorder" <% if SNW_AFFILIATE="0" then response.Write("checked") %>> 
		    No</p></td>
            
            <input type="hidden" name="exportType" value="0">
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td align="right"><input type="text" name="prdcount" size="3" value="<%=SNW_MAX%>"></td>
		<td nowrap="nowrap">Maximum of products per Widget (e.g. 25)<br />
		  <em>The more products you publish with your Widget the longer it will take to load.</em></td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
    <tr> 
      <td align="right" valign="top" nowrap="nowrap">
      	Choose a Category:</td>
      <td width="90%" valign="top">
            <select size="8" name="catlist">
            <% query="SELECT idcategory, idParentCategory, categorydesc FROM categories ORDER BY categoryDesc ASC"
            set rstemp=conntemp.execute(query)
            if err.number <> 0 then                
                call closeDb()
                response.redirect "techErr.asp?error="& Server.Urlencode("Error in retreiving categories from database: "&Err.Description) 
            end If
            if rstemp.eof then
                catnum="0"
            end If
            if catnum<>"0" then
                
				do until rstemp.eof
                    idparentcategory=""
					idparentcategory=rstemp("idParentCategory")
                   	if idparentcategory<>"1" then 
						
						set rs2=server.createobject("adodb.recordset")
						query="SELECT categoryDesc FROM categories where idcategory="& idparentcategory & " AND iBTOHide<>1"						
						set rs2=conntemp.execute(query)
						if NOT rs2.EOF then
							parentDesc=rs2("categoryDesc")							
							%>                            
							<option value='<%response.write rstemp("idcategory")%>'><%response.write rstemp("categorydesc")&" - Parent: "&parentDesc %></option>
							<% 
						end if
                	
					else
						%>
                        <option value='<%response.write rstemp("idcategory")%>' <% if cint(SNW_CATEGORY)=cint(rstemp("idcategory")) then response.Write("selected") %>><%=rstemp("categorydesc")%></option>
                        <%
					end if
					rstemp.movenext
                loop
            End if %>
        </select>
        <br />
        <br />
        <span class="pcCPnotes"><strong>Tip:</strong> We recommend creating a hidden category called something like &quot;E-Commerce Widget&quot;. The hidden category should contain all the products you want displayed in your widget. Then, select the &quot;E-Commerce Widget&quot; category in the category list shown menu above.</span>
        <br />
      </td>
    </tr>
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr> 
		<td colspan="2" style="text-align: center;">
			<input name="submit" type="submit" class="btn btn-primary" value="Generate ProductCart E-Commerce Widget">
		</td>
	</tr>
	</table>
</form>
<%
End If
%><!--#include file="AdminFooter.asp"-->
