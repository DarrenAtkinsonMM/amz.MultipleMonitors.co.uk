<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin="1*2*3*"%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/pcSeoFunctions.asp"-->
<% 
pageTitle="Generate Storefront Category Navigation" 
pageIcon="pcv4_icon_www.gif"
section="layout"

Dim query1,rsC,i,intCount,pcv_exportType,pcv_numprd
pcv_cats1=""
set StringBuilderObj = new StringBuilder



Function genCatInfor(tmp_IDCAT,tmp_CatName)
Dim tmp_showIcon,k,pcv_first,pcv_prdcount,pcv_back,pcv_end,tsT,tmp_query,pcv_CatName,pcv_prds1,tmp1,tmp2,pcv_CatNameFull,ProNameFull
	tmp1=0
	tmp2=0
	tmp_showIcon=0
	pcv_first=0
	pcv_prdcount=0
	maxProNameL = 22
	pcv_CatName=ClearHTMLTags2(tmp_CatName,0)
	pcv_CatNameFull=pcv_CatName
	If maxProNameL<len(pcv_CatName) then
		pcv_CatName=trim(left(pcv_CatName,maxProNameL)) & "..."
        pcv_CatName=pcv_CatName
	End If
	if (pcv_exportType="0" or pcv_exportType="2") then
		pcv_end=intCount
		tmp1=0
		For k=0 to intCount
			pcv_back=pcv_end-k
			if cint(pcv_cats1(2,k))=tmp_IDCAT then
				tmp1=1
				exit for
			end if
			if pcv_cats1(2,pcv_back)=tmp_IDCAT then
				tmp1=1
				exit for
			end if
		Next
	end if
	
	if (pcv_exportType="1" or pcv_exportType="2") then
		tmp_query="SELECT products.idProduct,products.description FROM products INNER JOIN categories_products ON (products.idProduct=categories_products.idProduct AND products.removed=0 AND products.active=-1 AND products.configOnly=0) WHERE categories_products.idcategory=" & tmp_IDCAT & " ORDER BY categories_products.POrder asc,products.description ASC;"
		set rsT=connTemp.execute(tmp_query)
		if not rsT.eof then
			tmp2=1
			pcv_prds1=rsT.getRows()
			set rsT=nothing
		else
			tmp2=0
		end if
		set rsT=nothing
	end if
	if ((tmp1=1) and (pcv_exportType="0" or pcv_exportType="2")) or ((tmp2=1) and (pcv_exportType="1" or pcv_exportType="2")) then
		tmp_showIcon=1
	end if
	StringBuilderObj.append "<li>"

	'// SEO Links
	'// Build Navigation Category Link
	pcStrNavCatLink=pcv_CatNameFull & "-c" & tmp_IDCAT & ".htm"
	pcStrNavCatLink=removeChars(pcStrNavCatLink)
	if scSeoURLs<>1 then
		pcStrNavCatLink="viewCategories.asp?idCategory=" & tmp_IDCAT
	end if
	if tmp_showIcon=1 then
	StringBuilderObj.append "<a href=""" & tempURL & pcStrNavCatLink & """ class=""MenuBarItemSubmenu"">" & pcv_CatName & "</a>" & vbcrlf	
	else
	StringBuilderObj.append "<a href=""" & tempURL & pcStrNavCatLink & """>" & pcv_CatName & "</a>" & vbcrlf	
	end if		
	'//
	
	if tmp_showIcon=1 then
		StringBuilderObj.append "<ul>" & vbcrlf
		if ((tmp1=1) and (pcv_exportType="0" or pcv_exportType="2")) then
			For k=0 to intCount
				if cint(pcv_cats1(2,k))=tmp_IDCAT then
					pcv_first=1
					call genCatInfor(pcv_cats1(0,k),pcv_cats1(1,k))
				else
					if pcv_first=1 then
						exit for
					end if
				end if
			Next
		end if
		if ((tmp2=1) and (pcv_exportType="1" or pcv_exportType="2")) then
			if ubound(pcv_prds1,2)>pcv_numprd-1 then
				pcv_prdcount=pcv_numprd-1
			else
				pcv_prdcount=ubound(pcv_prds1,2)
			end if
			For k=0 to pcv_prdcount
				ProName=ClearHTMLTags2(pcv_prds1(1,k),0)
				ProNameFull=ProName
				If maxProNameL<len(ProName) then
					ProName=trim(left(ProName,maxProNameL)) & "..."
				End If


				pIntPrdId=pcv_prds1(0,k)
				'// SEO Links
				'// Build Navigation Product Link
				pcStrNavPrdLink=ProNameFull & "-" & tmp_IDCAT & "p" & pIntPrdId & ".htm"
				pcStrNavPrdLink=removeChars(pcStrNavPrdLink)
				if scSeoURLs<>1 then
					pcStrNavPrdLink="viewPrd.asp?idcategory=" & tmp_IDCAT & "&idproduct=" & pcv_prds1(0,k)
				end if
				'//
				StringBuilderObj.append "<li><a href=""" & tempURL & pcStrNavPrdLink & """>" & ProName & "</a></li>" & vbcrlf
			Next
			if ubound(pcv_prds1,2)>pcv_numprd-1 then
				'// SEO Links
				'// Build Navigation Category Link
				pcStrNavCatLink=pcv_CatNameFull & "-c" & tmp_IDCAT & ".htm"
				pcStrNavCatLink=removeChars(pcStrNavCatLink)
				if scSeoURLs<>1 then
					pcStrNavCatLink="viewCategories.asp?idCategory=" & tmp_IDCAT
				end if
				'//
				StringBuilderObj.append "<li><a href=""" & tempURL & pcStrNavCatLink & """>More Products...</a></li>" & vbcrlf
			end if
		end if
		StringBuilderObj.append "</ul>" & vbcrlf
	end if
	StringBuilderObj.append "</li>" & vbcrlf
End Function
%>
<!--#include file="AdminHeader.asp"-->
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype, pcvTempQuery
Response.Write(pcf_InitializePrototype())


IF request("action")="add" THEN

	pcv_exportType=request("exportType")
	if pcv_exportType="" then
		pcv_exportType="0"
	end if
	if pcv_exportType="3" then
		pcvTempQuery = "AND categories.idParentCategory=1 "
		else
		pcvTempQuery = ""
	end if
	
	pcv_numprd=request("prdcount")
	if pcv_numprd="" or pcv_numprd="0" then
		pcv_numprd=5
	end if
		
	pcIntJQNav = request("JQNav") ' 1 = horizontal, 2 = vertical
	pcv_JQNavID = request("JQNavID")
	if not validNum(pcIntJQNav) or pcIntJQNav=0  then
		session("pcIntJQNav")=0
		'// Unordered list, own settings
		pcv_topULid=request("topULid")
		session("pcv_topULid")=pcv_topULid
		pcv_topULclass=request("topULclass")
		session("pcv_topULclass")=pcv_topULclass
	else
		session("pcIntJQNav")=pcIntJQNav
		if pcIntJQNav = 1 then
			pcv_JQNavClass = "ddsmoothmenu"
		else
			pcv_JQNavClass = "ddsmoothmenu-v"
		end if					
	end if
	
	pcvStr_linkType=request("linkType")
	if pcvStr_linkType="absolute" then
		tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
		tempURL=replace(tempURL,"http:/","http://")
		tempURL=replace(tempURL,"https:/","https://")
	else
		tempURL=""
	end if

	query="SELECT idCategory,categoryDesc,idParentCategory FROM categories WHERE categories.idCategory<>1 " & pcvTempQuery & "AND categories.iBTOhide=0 AND categories.pccats_RetailHide=0 ORDER BY categories.idParentCategory ASC,categories.priority ASC,categories.categoryDesc ASC;"
	set rsC=connTemp.execute(query)
	if not rsC.eof then
		pcv_cats1=rsC.GetRows()
		intCount=ubound(pcv_cats1,2)
		set rsC=nothing
    	
		If pcIntJQNav = 1 Or pcIntJQNav = 2 Then
			StringBuilderObj.append "<div id="""&pcv_JQNavID&""" >" & vbcrlf
			StringBuilderObj.append "<ul class="""&pcv_JQNavClass&""">" & vbcrlf & "" & vbcrlf
		Else
			pcv_topULid = pcv_JQNavID
			StringBuilderObj.append "<ul id="""&pcv_topULid&""" class="""&pcv_topULclass&""">" & vbcrlf & "" & vbcrlf
		End If
		For i=0 to intCount
			if pcv_cats1(2,i)="1" then
				call genCatInfor(pcv_cats1(0,i),pcv_cats1(1,i))
			else
				exit for
			end if
		Next
		StringBuilderObj.append "</ul>" & vbcrlf
		If pcIntJQNav = 1 Or pcIntJQNav = 2 Then
			StringBuilderObj.append "</div>"
		End If
	end if
	set rsC=nothing
	
	if PPD="1" then
		pcStrFolder="/"&scPcFolder&"/pc"
	else
		pcStrFolder="../pc"
	end if
	
	if pcIntJQNav = 1 then
		StringBuilderObj.append vbcrlf & "<script type=text/javascript>" & vbcrlf
		StringBuilderObj.append "$(function() {"
        StringBuilderObj.append "ddsmoothmenu.init({"
		StringBuilderObj.append "mainmenuid: """ & pcv_JQNavID & """," & vbcrlf
		StringBuilderObj.append "orientation: 'h'," & vbcrlf
		StringBuilderObj.append "classname: 'ddsmoothmenu'," & vbcrlf
		StringBuilderObj.append "zindexvalue: 100," & vbcrlf
		StringBuilderObj.append "contentsource: ""markup""" & vbcrlf
		StringBuilderObj.append "});" & vbcrlf
        StringBuilderObj.append "});"
		StringBuilderObj.append "</script>" & vbcrlf
	elseif pcIntJQNav = 2 then
		StringBuilderObj.append vbcrlf & "<script type=text/javascript>" & vbcrlf
		StringBuilderObj.append "$(function() {"
        StringBuilderObj.append "ddsmoothmenu.init({"
		StringBuilderObj.append "mainmenuid: """ & pcv_JQNavID & """," & vbcrlf
		StringBuilderObj.append "orientation: 'v'," & vbcrlf
		StringBuilderObj.append "classname: 'ddsmoothmenu-v'," & vbcrlf
		StringBuilderObj.append "arrowswap: true," & vbcrlf
		StringBuilderObj.append "zindexvalue: 100," & vbcrlf
		StringBuilderObj.append "contentsource: ""markup""" & vbcrlf
		StringBuilderObj.append "});" & vbcrlf
        StringBuilderObj.append "});"
		StringBuilderObj.append "</script>" & vbcrlf
	end if
	
	call pcs_SaveUTF8(pcStrFolder & "\inc_RetailCatMenu.inc",pcStrFolder & "\inc_RetailCatMenu.inc",StringBuilderObj.toString())

	Set StringBuilderObj=Nothing
tmp1=0
tmp2=0
pcv_cats1=""
set StringBuilderObj = new StringBuilder

	query="SELECT idCategory,categoryDesc,idParentCategory FROM categories WHERE categories.idCategory<>1 AND categories.iBTOhide=0 ORDER BY categories.idParentCategory ASC,categories.priority ASC,categories.categoryDesc ASC;"
	set rsC=connTemp.execute(query)
	if not rsC.eof then
		pcv_cats1=rsC.GetRows()
		intCount=ubound(pcv_cats1,2)
		set rsC=nothing

		If pcIntJQNav = 1 Or pcIntJQNav = 2 Then
			StringBuilderObj.append "<div id="""&pcv_JQNavID&""" >" & vbcrlf
			StringBuilderObj.append "<ul class="""&pcv_JQNavClass&""">" & vbcrlf & "" & vbcrlf
		Else
			pcv_topULid = pcv_JQNavID
			StringBuilderObj.append "<ul id="""&pcv_topULid&""" class="""&pcv_topULclass&""">" & vbcrlf & "" & vbcrlf
		End If
        
        
        For i=0 to intCount
			if pcv_cats1(2,i)="1" then
				call genCatInfor(pcv_cats1(0,i),pcv_cats1(1,i))
			else
				exit for
			end if
		Next

		StringBuilderObj.append "</ul>" & vbcrlf
		If pcIntJQNav = 1 Or pcIntJQNav = 2 Then
			StringBuilderObj.append "</div>"
		End If
        
		if pcIntJQNav = 1 then
			StringBuilderObj.append vbcrlf & "<script type=text/javascript>" & vbcrlf
            StringBuilderObj.append "$(function() {"
			StringBuilderObj.append "ddsmoothmenu.init({"
			StringBuilderObj.append "mainmenuid: """ & pcv_JQNavID & """," & vbcrlf
			StringBuilderObj.append "orientation: 'h'," & vbcrlf
			StringBuilderObj.append "classname: 'ddsmoothmenu'," & vbcrlf
			StringBuilderObj.append "zindexvalue: 100," & vbcrlf
			StringBuilderObj.append "contentsource: ""markup""" & vbcrlf
			StringBuilderObj.append "})" & vbcrlf
            StringBuilderObj.append "});"
			StringBuilderObj.append "</script>" & vbcrlf
		elseif pcIntJQNav = 2 then
			StringBuilderObj.append vbcrlf & "<script type=text/javascript>" & vbcrlf
            StringBuilderObj.append "$(function() {"
			StringBuilderObj.append "ddsmoothmenu.init({"
			StringBuilderObj.append "mainmenuid: """ & pcv_JQNavID & """," & vbcrlf
			StringBuilderObj.append "orientation: 'v'," & vbcrlf
			StringBuilderObj.append "classname: 'ddsmoothmenu-v'," & vbcrlf
			StringBuilderObj.append "arrowswap: true," & vbcrlf
			StringBuilderObj.append "zindexvalue: 100," & vbcrlf
			StringBuilderObj.append "contentsource: ""markup""" & vbcrlf
			StringBuilderObj.append "})" & vbcrlf
            StringBuilderObj.append "});"
			StringBuilderObj.append "</script>" & vbcrlf
		end if
		
	end if
	set rsC=nothing

	if PPD="1" then
		pcStrFolder="/"&scPcFolder&"/pc"
	else
		pcStrFolder="../pc"
	end if

	call pcs_SaveUTF8(pcStrFolder & "\inc_WholeSaleCatMenu.inc",pcStrFolder & "\inc_WholeSaleCatMenu.inc",StringBuilderObj.toString())

	Set StringBuilderObj=Nothing
%>
<table class="pcCPcontent">
<tr>
	<td align="center">
		<div class="pcCPmessageSuccess">
			New Storefront Category Navigation was created successfully!
			<br /><br />
			<a href="../pc/viewcategories.asp" target="_blank">View Storefront</a>&nbsp;|&nbsp;
            <%
			if session("pcIntJQNav")=1 then
			%>
                <a href="../pc/menupreview.asp" target="_blank">Preview JQuery Horizontal Menu</a>
                &nbsp;|&nbsp;
                <%
				elseif session("pcIntJQNav") = 2 then
				%>
                <a href="../pc/menupreview.asp" target="_blank">Preview JQuery Vertical Menu</a>
                &nbsp;|&nbsp;
                <%
				else
			end if
			%>
            <a href="genCatNavigation.asp">Generate New Navigation</a>
		</div>
	</td>
</tr>
<tr>
	<td class="pcSpacer">&nbsp;</td>
</tr>
</table>

<%
session("pcIntJQNav")=""
ELSE
%>

<form method="post" name="form1" action="genCatNavigation.asp?action=add" class="pcForms">
	<table class="pcCPcontent">
	<tr>
		<td colspan="2">
        	<h2>How it works</h2>
            <div>
			ProductCart will generate a <u>static</u> file to store your navigation links. This improves storefront performance (less database queries). Remember to rerun this feature when you add/edit categories (and products if included in the navigation). Your &quot;<strong>header.asp</strong>&quot; or &quot;<strong>footer.asp</strong>&quot; file must include the file &quot;<strong>inc_catsmenu.asp</strong>&quot; in order for the navigation to show.<a href="http://wiki.productcart.com/how_to/add_category_navigation" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" width="16" height="16" border="0"></a>
            </div>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcSpacer">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2">
        <h2>Create or Update your storefront category navigation</h2>
        <p>When creating your navigation file you have three options:</p></td>
	</tr>
	<tr>
		<td align="right" width="10%"><input type="radio" id="exportType_0" name="exportType" value="0" checked class="clearBorder"></td>
		<td width="90%"><label for="exportType_0">Include categories and sub-categories, but no products</label></td>
	</tr>
	<tr>
		<td align="right"><input type="radio" id="exportType_2" name="exportType" value="2" class="clearBorder"></td>
		<td><label for="exportType_2">Include categories, sub-categories, and their products</label></td>
	</tr>
	<tr>
		<td align="right"><input type="radio" id="exportType_1" name="exportType" value="1" class="clearBorder"></td>
		<td><label for="exportType_1">Include only top-level categories and their products</label></td>
	</tr>
	<tr>
		<td align="right"><input type="radio" id="exportType_3" name="exportType" value="3" class="clearBorder"></td>
		<td><label for="exportType_3">Include only top-level categories</label></td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td align="right"><input type="text" name="prdcount" size="3" value="5"></td>
		<td>Maximum of products per category (e.g. 5)<br />
		A &quot;More products...&quot; link is automatically added if there are more products in the category than the number specified here.</td>
	</tr>
	<tr>
		<td colspan="2" class="pcSpacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2">Do you want to use relative or absolute links?&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=445"></a></td>
	</tr>
	<tr>
		<td align="right"><input type="radio" id="linkTypeRelative" name="linkType" value="relative" class="clearBorder"></td>
    <td><label for="linkTypeRelative">Relative Links</label></td>
  </tr>
  <tr>
		<td align="right"><input type="radio" id="linkTypeAbsolute" name="linkType" value="absolute" class="clearBorder" checked="checked"></td>
		<td><label for="linkTypeAbsolute">Absolute Links</label></td>
	</tr>
	<tr>
		<td colspan="2" class="pcSpacer"><hr></td>
	</tr>
	<tr> 
		<td valign="top" colspan="2">
        
            <div class="bs-callout bs-callout-info">
                If using the Basic Blue or Bootstrap Themes you need to select &quot;JQuery Vertical Menu Bar&quot; below.
            </div>
            <br />
            Choose an option for the <strong>JQuery menu bar</strong>: <a href="http://www.dynamicdrive.com/dynamicindex1/ddsmoothmenu.htm" target="_blank">Documentation</a>
        </td>
  </tr>
  <tr>
    <td align="right"><input type="radio" id="JQNavHorizontal" name="JQNav" value="1" onChange="selectShowNavOptions(this);"></td>
		<td><label for="JQNavHorizontal">JQuery Horizontal Menu Bar </label></td>
	</tr>
  <tr>
		<td align="right"><input type="radio" id="JQNavVertical" name="JQNav" value="2" onChange="selectShowNavOptions(this);" checked></td>
    <td><label for="JQNavVertical">JQuery Vertical Menu Bar</label></td>
	</tr>
	<tr>
		<td align="right"><input type="radio" id="JQNavOther" name="JQNav" value="0" onChange="selectShowNavOptions(this);"></td>
		<td><label for="JQNavOther">Other</label></td>
	</tr>
	<tr>
		<td colspan="2">
			<table class="pcCPcontent" id="JQNavOptions">
				<tr> 
					<td align="right">
						ID of JQuery menu bar:
					</td>
					<td>
						<input type="text" name="JQNavID" value="<%if session("pcvULID")<>"" then response.write(session("pcvULID")) else response.write "menubar99" end if%>" size="30"> <span class="pcSmallText">See JQuery menu bar documentation for details.</span>
					</td>
				</tr>
			</table>
			<table class="pcCPcontent" id="OtherNavOptions" style="display: none;">
				<tr><td colspan="2"><strong>Use your own CSS</strong></td></tr>
				<tr><td colspan="2"><i>You can assign a CSS class to the most relevant elements in the unordered list.</i></td></tr>
        	
          
				<tr><td align="right" width="15%">Top UL tag ID:</td><td><input type="text" name="topULid" value="<%=session("pcv_topULid")%>" size="30"></td></tr>
				<tr><td align="right" width="15%">Top UL tag Class:</td><td><input type="text" name="topULclass" value="<%if session("pcv_topULclass")<>"" then response.write(session("pcv_topULclass")) else response.write "dropdown-menu" end if%>" size="30"></td></tr>
				<tr><td colspan="2" class="pcCPspacer"></td></tr>
				<tr>
					<td colspan="2">
						<div class="pcCPmessage">NOTE: make sure that your store interface (header.asp, footer.asp) contains the CSS and JS files needed to style the unordered list.</div>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcSpacer"></td>
	</tr>
	<tr>
		<td colspan="2" class="pcSpacer"><hr></td>
	</tr>
	<tr> 
		<td colspan="2">
			<input name="submit" type="submit" class="btn btn-primary" value="Generate Storefront Category Navigation" onClick="pcf_Open_genCatNav();">
			<%
            '// Loading Window
            '	>> Call Method with OpenHS();
            response.Write(pcf_ModalWindow("Generating category navigation. Please wait...", "genCatNav", 300))
            %>
		</td>
	</tr>
	</table>
</form>
<%
END IF
%>
<!--#include file="AdminFooter.asp"-->
