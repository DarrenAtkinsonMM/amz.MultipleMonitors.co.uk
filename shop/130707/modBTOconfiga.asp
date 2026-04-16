<% 'CONFIGURATOR ONLY FILE 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Setup Configurable Products or Services" %>
<% section = "services" %>
<%PmAdmin=2%> 
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="pcCalculateBTODefaultPrices.asp" -->
<!--#include file="inc_UpdateDates.asp" -->
<!--#include file="AdminHeader.asp"-->
<script type=text/javascript>
function showlaptop(theTable)
{
     obj = document.getElementsByTagName('TR');
      for (i=0; i<obj.length; i++)
     {
          if (obj[i].id == theTable)
          obj[i].style.display = '';
     }
}
function hidelaptop(theTable)
{
     obj = document.getElementsByTagName('TR');
      for (i=0; i<obj.length; i++)
     {
          if (obj[i].id == theTable)
          obj[i].style.display = 'none';
     }
}
</script>
<% dim f

' form parameter 
pidProduct=request.Querystring("idProduct")
if pidProduct="" then
	pidProduct=request.Form("idProduct")
end if

if trim(pidProduct)="" then
	call closeDb()
	response.redirect "msg.asp?message=2"
end if

'check for customer categories
dim intCCExists
intCCExists=0 
query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories;"
SET rs=Server.CreateObject("ADODB.RecordSet")
SET rs=conntemp.execute(query)
if NOT rs.eof then 
	intCCExists=1
	CCArray=rs.getrows()
	intCountCC=ubound(CCArray,2)
end if
SET rs=nothing

If request.form("Submit2")<>"" then
'******* Update Configurable Items
	oCnt = request.Form("oCnt")
	for o = 1 to oCnt
		catSort=request.Form("catSort"&o)
		If catSort="" then
			catSort="0"
		End If
		requiredCategory=request.Form("requiredCategory"&o)
		If requiredCategory="" then
			requiredCategory="0"
		End If
		showInfo=request.Form("showInfo"&o)
		If showInfo="" then
			showInfo="0"
		End If
		displayQF=request.Form("displayQF"&o)
		If displayQF="" then
			displayQF="0"
		end If
		pcv_ShowDesc=request.Form("pcv_ShowDesc"&o)
		if pcv_ShowDesc="" then
			pcv_ShowDesc="0"
		end if
		pcv_ShowImg=request.Form("pcv_ShowImg"&o)
		if pcv_ShowImg="" then
			pcv_ShowImg="0"
		end if
		pcv_ImgWidth=request.Form("pcv_ImgWidth"&o)
		if pcv_ImgWidth="" or pcv_ImgWidth="0" then
			pcv_ImgWidth="35"
		end if
		pcv_ShowSKU=request.Form("pcv_ShowSKU"&o)
		if pcv_ShowSKU="" then
			pcv_ShowSKU="0"
		end if
		pcv_UseRadio=request.Form("pcv_UseRadio"&o)
		if pcv_UseRadio="" then
			pcv_UseRadio="0"
		end if
		If pcv_UseRadio="2" then
			pcv_UseRadio="0"
			multiSelect="1"
		else
			multiSelect="0"
		end If
		CATNotes=request.Form("Notes"&o)
		if CATNotes<>"" then
		CATNotes=replace(CATNotes,"'","''")
		end if
		idCategory=request.Form("CATID"&o)
		pCnt =request.Form("pCnt"&o)
		for p=1 to pCnt
			id=request.Form("id"&p&"_"&idCategory)

			If multiSelect="1" then
				cdefault = request.Form("cdefault"&idCategory&"_"&p)
				If cint(cdefault)=cint(p) then
					pcdefault="1"
				else
					pcdefault="0"
				end if
			Else
				cdefault = request.Form("cdefault"&idCategory)
				If cdefault="XX" then
					pcdefault="0"
				Else
					If cint(cdefault)=cint(p) then
						pcdefault="1"
					else
						pcdefault="0"
					end if
				End If
			End if

			prdSort=request.Form("prdSort"&p&"_"&id&"_"&idCategory)
			If prdSort="" then
				prdSort="0"
			End If
			Wprice=request.Form("Wprice"&p&"_"&id&"_"&idCategory)
			If Wprice="" then
				Wprice="0"
			End If

			rPrice=request.Form("rPrice"&p&"_"&id&"_"&idCategory)
			If rPrice<>"" then
				price=request.Form("rPrice"&p&"_"&id&"_"&idCategory)
				If price="" then
					price="0"
				End If
			Else
				price=request.Form("price"&p&"_"&id&"_"&idCategory)
				If price="" then
					price="0"
				End If
			End If
			Wprice = replacecomma(Wprice)
			price = replacecomma(price)
			'enter everything into the database
			query="UPDATE configSpec_products SET pcConfPro_ShowDesc=" & pcv_ShowDesc & ",pcConfPro_ShowImg=" & pcv_ShowImg & ",pcConfPro_ImgWidth=" & pcv_ImgWidth & ",pcConfPro_ShowSKU=" & pcv_ShowSKU & ",pcConfPro_UseRadio=" & pcv_UseRadio & ",price="&price&", Wprice="&Wprice&",cdefault="&pcdefault&", requiredCategory="&requiredCategory&", showInfo="&showInfo&", multiSelect="&multiSelect&", prdSort="&prdSort&", catSort="&catSort&",displayQF=" & displayQF & ",Notes=N'" & CATNotes & "' WHERE specProduct="&pidProduct&" AND configProduct="&id&" AND configProductCategory="&idCategory&";"
			set rs=Server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing
			'if cc exists
			if intCCExists=1 then
				For intCC=0 to intCountCC
					idcustomerCategory=CCArray(0,intCC)
					idCC_BTO_Price=request.Form("idCCBTO"&p&"_"&id&"_"&idCategory&"_"&idcustomerCategory)
					idCC_BTO_Price=replacecomma(idCC_BTO_Price)
					'get price
					pcCC_BTO_Price=request.Form("CCBTOPrice"&p&"_"&id&"_"&idCategory&"_"&idcustomerCategory)
					pcCC_BTO_Price=replacecomma(pcCC_BTO_Price)
					query="SELECT idCustomerCategory, idBTOProduct, idBTOItem FROM pcCC_BTO_Pricing WHERE idCustomerCategory=" & idcustomerCategory & " AND idBTOProduct=" & pidProduct & " AND idBTOItem=" & id & ";"
					set rs=connTemp.execute(query)
					if rs.eof then
						query="INSERT INTO pcCC_BTO_Pricing (idCustomerCategory, idBTOProduct, idBTOItem, pcCC_BTO_Price) VALUES ("&idcustomerCategory&","&pidProduct&","&id&","&pcCC_BTO_Price&");"
					else
						query="UPDATE pcCC_BTO_Pricing SET pcCC_BTO_Price="&pcCC_BTO_Price&" WHERE idCustomerCategory=" & idcustomerCategory & " AND idBTOProduct=" & pidProduct & " AND idBTOItem=" & id & ";"
					end if
					set rs=nothing
					set rs=Server.CreateObject("ADODB.RecordSet")
					set rs=conntemp.execute(query)
					set rs=nothing

				Next
			end if
		next
	next
	
	'****** End of Update Configurable Items
	
	'****** Update Additional Charges
	
	oCnt = request.Form("CHGoCnt")
	for o = 1 to oCnt
	
		catSort=request.Form("CHGcatSort"&o)
		If not validNum(catSort) then catSort="0"

		requiredCategory=request.Form("CHGrequiredCategory"&o)
		If not validNum(requiredCategory) then requiredCategory="0"

		showInfo=request.Form("CHGshowInfo"&o)
		If not validNum(showInfo) then showInfo="0"

		displayQF=request.Form("CHGdisplayQF"&o)
		If not validNum(displayQF) then displayQF="0"

		pcv_AShowDesc=request.Form("CHGpcv_AShowDesc"&o)
		if not validNum(pcv_AShowDesc) then pcv_AShowDesc="0"

		pcv_AShowImg=request.Form("CHGpcv_AShowImg"&o)
		if not validNum(pcv_AShowImg) then pcv_AShowImg="0"
		
		pcv_AImgWidth=request.Form("CHGpcv_AImgWidth"&o)
		if not validNum(pcv_AImgWidth) then pcv_AImgWidth="35"
		
		pcv_AShowSKU=request.Form("CHGpcv_AShowSKU"&o)
		if not validNum(pcv_AShowSKU) then pcv_AShowSKU="0"
		
		pcv_UseRadio=request.Form("CHGpcv_UseRadio"&o)
		if not validNum(pcv_UseRadio) then pcv_UseRadio="0"
			If pcv_UseRadio="2" then
				pcv_UseRadio="0"
				multiSelect="1"
			else
				multiSelect="0"
			end If
		CHCATNotes=request.Form("CHNotes"&o)
		if CHCATNotes<>"" then
		CHCATNotes=replace(CHCATNotes,"'","''")
		end if
		idCategory=request.Form("CHGCATID"&o)
		pCnt =request.Form("CHGpCnt"&o)
		for p=1 to pCnt
			id=request.Form("CHGid"&p&"_"&idCategory)
			If multiSelect="1" then
			Else
			End if
			cdefault = request.Form("CHGcdefault"&idCategory)
			If cdefault="XX" then
				pcdefault="0"
			Else
				If cint(cdefault)=cint(p) then
					pcdefault="1"
				else
					pcdefault="0"
				end if
			End If
			prdSort=request.Form("CHGprdSort"&p&"_"&id&"_"&idCategory)
			If prdSort="" then
				prdSort="0"
			End If
			Wprice=request.Form("CHGWprice"&p&"_"&id&"_"&idCategory)
			If Wprice="" then
				Wprice="0"
			End If

			rPrice=request.Form("CHGrPrice"&p&"_"&id&"_"&idCategory)
			If rPrice<>"" then
				price=request.Form("CHGrPrice"&p&"_"&id&"_"&idCategory)
				If price="" then
					price="0"
				End If
			Else
				price=request.Form("CHGprice"&p&"_"&id&"_"&idCategory)
				If price="" then
					price="0"
				End If
			End If

			price=replacecomma(price)
			Wprice=replacecomma(Wprice)
			'enter everything into the database
			query="UPDATE configSpec_Charges SET pcConfCha_ShowDesc=" & pcv_AShowDesc & ",pcConfCha_ShowImg=" & pcv_AShowImg & ",pcConfCha_ImgWidth=" & pcv_AImgWidth & ",pcConfCha_ShowSKU=" & pcv_AShowSKU & ",pcConfCha_UseRadio=" & pcv_UseRadio & ",price="&price&", Wprice="&Wprice&",cdefault="&pcdefault&", requiredCategory="&requiredCategory&", showInfo="&showInfo&", multiSelect="&multiSelect&", prdSort="&prdSort&", catSort="&catSort&",displayQF=" & displayQF & ",Notes=N'" & CHCATNotes & "' WHERE specProduct="&pidProduct&" AND configProduct="&id&" AND configProductCategory="&idCategory&";"
			set rs=Server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing
			'if cc exists
			if intCCExists=1 then
				For intCC=0 to intCountCC
					idcustomerCategory=CCArray(0,intCC)
					idCC_BTO_Price=request.Form("CHGidCCBTO"&p&"_"&id&"_"&idCategory&"_"&idcustomerCategory)
					'get price
					pcCC_BTO_Price=request.Form("CHGCCBTOPrice"&p&"_"&id&"_"&idCategory&"_"&idcustomerCategory)
					idCC_BTO_Price=replacecomma(idCC_BTO_Price)
					pcCC_BTO_Price=replacecomma(pcCC_BTO_Price)
					query="SELECT idCustomerCategory, idBTOProduct, idBTOItem FROM pcCC_BTO_Pricing WHERE idCustomerCategory=" & idcustomerCategory & " AND idBTOProduct=" & pidProduct & " AND idBTOItem=" & id & ";"
					set rs=connTemp.execute(query)
					if rs.eof then
						query="INSERT INTO pcCC_BTO_Pricing (idCustomerCategory, idBTOProduct, idBTOItem, pcCC_BTO_Price) VALUES ("&idcustomerCategory&","&pidProduct&","&id&","&pcCC_BTO_Price&");"
					else
						query="UPDATE pcCC_BTO_Pricing SET pcCC_BTO_Price="&pcCC_BTO_Price&" WHERE idCustomerCategory=" & idcustomerCategory & " AND idBTOProduct=" & pidProduct & " AND idBTOItem=" & id & ";"
					end if
					set rs=nothing
				Next
			end if
		next
	next
	
	call updPrdEditedDate(pidProduct)
	
	
'****** End of Update Additional Charges
	call closeDb()
	response.redirect "modBTOconfiga.asp?idProduct="& pidProduct
End If

If (request.form("DelButton")<>"") OR (request.form("DelButton1")<>"") then
	aCnt=request("aCnt")
	if IsNumeric(aCnt) then
		For q=1 to aCnt
			id = request("DP" & q)
			if id<>"" then
				id=split(id,"_")
				if id(0)="BTO" then
					query="DELETE FROM configSpec_products WHERE configProduct="& id(1) &" AND specProduct="& pidProduct
					set rs=Server.CreateObject("ADODB.RecordSet")
					set rs=conntemp.execute(query)
					set rs=nothing
				else
					if id(0)="CHG" then
						query="DELETE FROM configSpec_Charges WHERE configProduct="& id(1) &" AND specProduct="& pidProduct
						set rs=Server.CreateObject("ADODB.RecordSet")
						set rs=conntemp.execute(query)
						set rs=nothing
					end if
				end if
			end if
		Next
	end if
	call closeDb()
	response.redirect "modBTOconfiga.asp?idProduct="& pidProduct
End If

dim showBtm
showBtm="0"

Sub TestAvailableItems(tempBTOID)
	Dim rs,query,rs1
	Dim TestPrd,strPrd
	Dim TestCat,strCat
	Dim pcArr,pcArr1,i,j,intCount,intCount1
	
	TestPrd=0
	strPrd=""
	TestCat=0
	strCat=""
	
	query="SELECT pcBR_IDSourcePrd,pcBR_ID,pcBR_Must_Exists,pcBR_CanNot_Exists,pcBR_CatMust_Exists,pcBR_CatCanNot_Exists FROM pcBTORules WHERE pcBR_IDBTOPrd=" & tempBTOID
	set rs=connTemp.execute(query)
	
	if rs.eof then
		set rs=nothing
		exit sub
	else
		pcArr=rs.getRows()
		intCount=ubound(pcArr,2)
		set rs=nothing
		
		For i=0 to intCount
			if pcArr(2,i)="1" then
				query="SELECT pcBRMust_Item,description FROM products INNER JOIN pcBRMust ON products.idproduct=pcBRMust.pcBRMust_Item WHERE pcBR_ID=" & pcArr(1,i)
				set rs=connTemp.execute(query)
				if not rs.eof then
					pcArr1=rs.getRows()
					intCount1=ubound(pcArr1,2)
					set rs=nothing
					For j=0 to intCount1
						query="SELECT configProduct FROM configSpec_products WHERE configProduct=" & pcArr1(0,j) & " AND specProduct=" & tempBTOID
						set rs=connTemp.execute(query)
						if rs.eof then
							TestPrd=1
							if strPrd<>"" then
								strPrd=strPrd & ", "
							end if
							strPrd=strStr & """" & pcArr1(1,j) & """"
						end if
						set rs=nothing
					Next
				end if
				set rs=nothing
			end if
			
			if pcArr(3,i)="1" then
				query="SELECT pcBRCanNot_Item,description FROM products INNER JOIN pcBRCanNot ON products.idproduct=pcBRCanNot.pcBRCanNot_Item WHERE pcBR_ID=" & pcArr(1,i)
				set rs=connTemp.execute(query)
				if not rs.eof then
					pcArr1=rs.getRows()
					intCount1=ubound(pcArr1,2)
					set rs=nothing
					For j=0 to intCount1
						query="SELECT configProduct FROM configSpec_products WHERE configProduct=" & pcArr1(0,j) & " AND specProduct=" & tempBTOID
						set rs=connTemp.execute(query)
						if rs.eof then
							TestPrd=1
							if strPrd<>"" then
								strPrd=strPrd & ", "
							end if
							strPrd=strStr & """" & pcArr1(1,j) & """"
						end if
						set rs=nothing
					Next
				end if
				set rs=nothing
			end if
			
			if pcArr(4,i)="1" then
				query="SELECT pcBRCatMust_Item,categoryDesc FROM categories INNER JOIN pcBRCatMust ON categories.idCategory=pcBRCatMust.pcBRCatMust_Item WHERE pcBR_ID=" & pcArr(1,i)
				set rs=connTemp.execute(query)
				if not rs.eof then
					pcArr1=rs.getRows()
					intCount1=ubound(pcArr1,2)
					set rs=nothing
					For j=0 to intCount1
						query="SELECT configProductCategory FROM configSpec_products WHERE configProductCategory=" & pcArr1(0,j) & " AND specProduct=" & tempBTOID
						set rs=connTemp.execute(query)
						if rs.eof then
							TestCat=1
							if strCat<>"" then
								strCat=strCat & ", "
							end if
							strCat=strCat & """" & pcArr1(1,j) & """"
						end if
						set rs=nothing
					Next
				end if
				set rs=nothing
			end if
			
			if pcArr(5,i)="1" then
				query="SELECT pcBRCatCanNot_Item,categoryDesc FROM categories INNER JOIN pcBRCatCanNot ON categories.idCategory=pcBRCatCanNot.pcBRCatCanNot_Item WHERE pcBR_ID=" & pcArr(1,i)
				set rs=connTemp.execute(query)
				if not rs.eof then
					pcArr1=rs.getRows()
					intCount1=ubound(pcArr1,2)
					set rs=nothing
					For j=0 to intCount1
						query="SELECT configProductCategory FROM configSpec_products WHERE configProductCategory=" & pcArr1(0,j) & " AND specProduct=" & tempBTOID
						set rs=connTemp.execute(query)
						if rs.eof then
							TestCat=1
							if strCat<>"" then
								strCat=strCat & ", "
							end if
							strCat=strCat & """" & pcArr1(1,j) & """"
						end if
						set rs=nothing
					Next
				end if
				set rs=nothing
			end if
		Next
		
		if (TestPrd=1) and (strPrd<>"") then
			%>
			<div class="pcCPmessage">
				Rule Conflict: Item <%=strPrd%> appears in one ore more rules for this product, but is no longer part of the product configuration. Add it back to the product configuration or remove it from the rules that contain it. Otherwise the storefront pages will malfunction.
			</div>
			<%
		end if
		
		if (TestCat=1) and (strCat<>"") then
			%>
			<br>
			<div class="pcCPmessage">
				Rule Conflict: Category <%=strCat%> appears in one ore more rules for this product, but is no longer part of the product configuration. Add it back to the product configuration or remove it from the rules that contain it. Otherwise the storefront pages will malfunction.
			</div>
			<%
		end if
		
	end if
		
End Sub

Call RunCalBDPC()

%>
<% ' START show message, if any %>
	<div align="center"><!--#include file="pcv4_showMessage.asp"--></div>
<% 	' END show message %>
<form method="post" name="modifyProduct" action="modBTOconfiga.asp" class="pcForms">
	<table class="pcCPcontent">

<%
			strSQL="SELECT categories.categoryDesc, configSpec_products.price AS cPrice, configSpec_products.Wprice, configSpec_products.cdefault, configSpec_products.prdSort, configSpec_products.catSort, configSpec_products.requiredCategory, configSpec_products.multiSelect, categories.idCategory, configSpec_products.configProduct, products.price AS price, configSpec_products.configProductCategory,products.description, configSpec_products.showInfo, configSpec_products.displayQF, configSpec_products.pcConfPro_ShowDesc, configSpec_products.pcConfPro_ShowImg, configSpec_products.pcConfPro_ImgWidth, configSpec_products.pcConfPro_ShowSKU,products.active,products.removed,products.stock,products.nostock,products.pcProd_BackOrder,configSpec_products.pcConfPro_UseRadio,configSpec_products.notes FROM (configSpec_products INNER JOIN categories ON configSpec_products.configProductCategory = categories.idCategory) INNER JOIN products ON configSpec_products.configProduct = products.idProduct WHERE (((configSpec_products.specProduct)="&pidProduct&")) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort, products.description;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(strSQL)
			if rs.eof then 
			ShowBtm="1"
			set rs=nothing
%>
	<tr>
		<td colspan="8" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td colspan="8" align="center">
			<div class="pcCPmessage">This product/service has not yet been configured.</div>
		</td>
	</tr>
	<tr>
		<td colspan="8" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td colspan="8" align="center">
		<input type="button" class="btn btn-default"  value="Start Configuration" onClick="location.href='AddBTOcat.asp?idProduct=<%=pidProduct%>'">
		&nbsp;
		<input type="button" class="btn btn-default"  value="Copy from another configurable product" onClick="location.href='ApplyBTOConfiga.asp?idProduct=<%=pidProduct%>'">
		<br /><br />
		<input type="button" class="btn btn-default"  value="Locate another configurable Product" onClick="location.href='LocateProducts.asp?cptype=1'">
		</td>
	</tr>
<% else	%>
	<tr>
		<td colspan="8">
        <a name="top"></a>
		<% if showBtm="0" then%>
		<h2>You are configuring: <strong><%=pDescription%></strong></h2>
		<% If statusCM="1" Then %>

			<a href="modifyProduct.asp?idProduct=<%=pidProduct%>&prdType=bto"><%=pDescription%></a></strong> (<a href="../pc/configurePrd.asp?idproduct=<%=pidProduct%>&adminpreview=1" target="_blank">Preview &gt;&gt;</a>)</p>

			<table cellpadding="6" cellspacing="0" border="0" align="center" width="700" style="border-color:#e1e1e1; border-style: solid; border-width:thin; background-color: #f5f5f5; margin: 10px 0 0 0;">
			<tr class="normal">
				<td align="right">
					<a href="AddBTOcat.asp?idProduct=<%=pidProduct%>"><img src="images/add.gif" width="14" height="16" hspace="5" border="0"></a>
				</td>
				<td align="left">
					<a href="AddBTOcat.asp?idProduct=<%=pidProduct%>">Add Item Category</a>
				</td>
				<td align="right">&nbsp;</td>
				<td align="right">
					<a href="AddBTOchg.asp?idProduct=<%=pidProduct%>"><img src="images/add.gif" width="14" height="16" hspace="5" border="0"></a>
				</td>
				<td align="left">
					<a href="AddBTOchg.asp?idProduct=<%=pidProduct%>">Add Additional Charges</a>
				</td>
				<td align="right">&nbsp;</td>
				<td align="right">
					<a href="bto_managerules.asp?idProduct=<%=pidProduct%>"><img src="images/quick.gif" width="24" height="18" hspace="5" border="0"></a>
				</td>
				<td align="left">
					<a href="bto_managerules.asp?idProduct=<%=pidProduct%>">Conflict Management</a>
				</td>
				<td align="right">&nbsp;</td>
				<td align="right">
					<a href="../pc/configurePrd.asp?idproduct=<%=pidProduct%>" target="_blank"><img src="images/move.gif" width="25" height="18" hspace="5" border="0"></a>
				</td>
				<td align="left">
					<a href="../pc/configurePrd.asp?idproduct=<%=pidProduct%>&adminpreview=1" target="_blank">Preview Configuration Page</a>
				</td>
				<td align="right">
					<a href="ApplyBTOConfigMulti.asp?idProduct=<%=pidProduct%>"><img src="images/quick.gif" width="24" height="18" hspace="5" border="0"></a>
				</td>
				<td align="left">
					<a href="ApplyBTOConfigMulti.asp?idProduct=<%=pidProduct%>">Apply Configuration to other Configurable Products</a>
				</td>
			</tr>
			</table>

		<% Else %>
        	<div class="cpOtherLinks"><a href="AddBTOcat.asp?idProduct=<%=pidProduct%>">Add a Category</a> : <a href="AddBTOchg.asp?idProduct=<%=pidProduct%>">Add Additional Charges</a> : <a href="#addcharges">Manage Additional Charges</a> : <a href="ApplyBTOConfigMulti.asp?idProduct=<%=pidProduct%>">Apply Configuration to another Configurable Product</a> : <a href="modifyProduct.asp?idProduct=<%=pidProduct%>&prdType=bto">Edit Product</a> : <a href="../pc/viewprd.asp?idproduct=<%=pidProduct%>&adminpreview=1" target="_blank">Preview</a></div>


		<% End If %>

		<% end if %>
		</td>
	</tr>
	<% If statusCM="1" Then %>
	<tr>
		<td colspan="8" align="center">
			<%call TestAvailableItems(pidProduct)%>
		</td>
	</tr>
	<tr>
		<td colspan="8" class="pcCPspacer"></td>
	</tr>
	<% End If %>
	<tr>
		<th colspan="8">Selectable Items - <a href="AddBTOcat.asp?idProduct=<%=pidProduct%>">Add more</a></th>
	</tr>
	<%
	aCnt=0
	oCnt = 0
	pCnt = 0
	btm="0"
	pcArr=rs.getRows()
	intCount=ubound(pcArr,2)
	set rs=nothing
	For i=0 to intCount
		strCategoryDesc=pcArr(0,i)
		dblcPrice=pcArr(1,i)
		dblWprice=pcArr(2,i)
		cdefault=pcArr(3,i)
		strPrdSort=pcArr(4,i)
		strCatSort=pcArr(5,i)
		requiredCategory=pcArr(6,i)
		multiSelect=pcArr(7,i)
		tempCat1 = pcArr(8,i)
		intConfigProduct=pcArr(9,i)
		dblprice=pcArr(10,i)
		configProductCategory=pcArr(11,i)
		strDesc=pcArr(12,i)
		pcv_active=pcArr(19,i)
		pcv_removed=pcArr(20,i)
		pcv_stock=pcArr(21,i)
		pcv_nostock=pcArr(22,i)
		pcv_backorder=pcArr(23,i)
		pcv_HideMsg=""
		if pcv_removed<>"0" then
			pcv_HideMsg="(Hidden - Removed)"
		end if
		if pcv_HideMsg="" then
			if pcv_active="0" then
				pcv_HideMsg="(Hidden - Inactive)"
			end if
		end if
		if pcv_HideMsg="" then
			if (scOutOfStockPurchase="-1") AND (iBTOOutofStockPurchase="-1") then
				if (clng(pcv_stock)<=0) AND (pcv_nostock="0") AND (pcv_backorder="0") then
					pcv_HideMsg="(Hidden - Out of Stock)"
				end if
			end if
		end if
		if pcv_HideMsg<>"" then
			pcv_HideMsg="<b><i>" & pcv_HideMsg & "</i></b>"
		end if
		pCnt = pCnt+1
		if Cint(tempCat2) <> Cint(tempCat1) then 
			if btm="1" then
%>
			<input type="hidden" name="pCnt<%=oCnt%>" value="<%=pCnt-1%>">
<% 
			end if
			oCnt = oCnt + 1
			pCnt = 1

			multiSelectVar="0"
			showInfoVar="0"
			requiredCategoryVar="0"
			pcv_ShowDesc="0"
			pcv_ShowImg="0"
			pcv_ImgWidth="35"
			pcv_ShowSKU="0"
			pcv_UseRadio="0"
				if multiSelect=True then
					multiSelectVar="1"
				end if
				if pcArr(13,i)=True then
					showInfoVar="1"
				end if
				if requiredCategory=True then
					requiredCategoryVar="1"
				end if
				if (pcArr(14,i)=-1) or (pcArr(14,i)=1) then
					displayQF="1"
				else
					displayQF="0"	
				end if
				if pcArr(15,i)="1" then
					pcv_ShowDesc="1"
				end if
				if pcArr(16,i)="1" then
					pcv_ShowImg="1"
				end if
				if pcArr(17,i)>"0" then
					pcv_ImgWidth=pcArr(17,i)
				end if
				if pcArr(18,i)="1" then
					pcv_ShowSKU="1"
				end if
				if pcArr(24,i)="1" then
					pcv_UseRadio="1"
				end if
				CATNotes=pcArr(25,i)
			%>
			<tr>
				<td colspan="8">&nbsp;</td>
			</tr>
			<tr bgcolor="#e5e5e5"> 
				<td colspan="6"> 
					<span style="font-size:14px; font-weight: bold;"><%=strCategoryDesc%></span> 
					<input type="hidden" name="CATID<%=oCnt%>" value="<%=tempCat1%>">
					&nbsp;&nbsp;
					<a href="AddBTOprd.asp?catSort=<%=strCatSort%>&idProduct=<%=pidProduct%>&idCategory=<%=tempCat1%>">Add more items</a>
					&nbsp;|&nbsp;
					<a href="javascript:if (confirm('You are about to remove this category from the product configuration. Are you sure you want to complete this action?')) location='DelBTOCat.asp?BTOproduct=<%=pidProduct%>&idCategory=<%=tempCat1%>&CatType=1';">Remove</a>
				</td>
				<td colspan="2" align="right">
					Order: <input type="text" name="catSort<%=oCnt%>" size="1" maxlength="3" value="<%=strCatSort%>" style="text-align: right; font-size: 8pt; font-weight: bold; color: #000000; background-color: #99CCFF">
				</td>
			</tr>

			<tr> 
				<td>Settings:</td>
				<td colspan="7">
				<% if showInfoVar="1" then %>
					<input type="checkbox" name="showInfo<%=oCnt%>" value="1" checked class="clearBorder">
				<% else %>
					<input type="checkbox" name="showInfo<%=oCnt%>" value="1" class="clearBorder">
				<% end if %>
				Show Details 
				<% if requiredCategoryVar="1" then %>
					<input type="checkbox" name="requiredCategory<%=oCnt%>" value="1" checked class="clearBorder">
				<% else %>
					<input type="checkbox" name="requiredCategory<%=oCnt%>" value="1" class="clearBorder">
				<% end if %>
				Required Category
				<% if displayQF="1" then %>
					<input type="checkbox" name="displayQF<%=oCnt%>" value="1" checked class="clearBorder">
				<% else %>
					<input type="checkbox" name="displayQF<%=oCnt%>" value="1" class="clearBorder">
				<% end if %>
				Display Quantity Field
				</td>
			</tr>
			<tr>
				<td></td>
				<td colspan="7">
				<%if pcv_ShowDesc="1" then %>
					<input type="checkbox" name="pcv_ShowDesc<%=oCnt%>" value="1" checked class="clearBorder">
				<% else %>
					<input type="checkbox" name="pcv_ShowDesc<%=oCnt%>" value="1" class="clearBorder">
				<% end if %>
				Show Product Description&nbsp;
				<%if pcv_ShowImg="1" then %>
					<input type="checkbox" name="pcv_ShowImg<%=oCnt%>" value="1" checked class="clearBorder">
				<% else %>
					<input type="checkbox" name="pcv_ShowImg<%=oCnt%>" value="1" class="clearBorder">
				<% end if %>
				Show Item Image&nbsp;&nbsp;-&nbsp;&nbsp;
				Image Width = <input type="text" name="pcv_ImgWidth<%=oCnt%>" value="<%=pcv_ImgWidth%>" size="4">&nbsp;&nbsp;
				<%if pcv_ShowSKU="1" then %>
					<input type="checkbox" name="pcv_ShowSKU<%=oCnt%>" value="1" checked class="clearBorder">
				<% else %>
					<input type="checkbox" name="pcv_ShowSKU<%=oCnt%>" value="1" class="clearBorder">
				<% end if %>
				Show Item SKU<br>
				<input type="radio" name="pcv_UseRadio<%=oCnt%>" value="0" <%if pcv_UseRadio<>"1" AND multiSelectVar<>"1" then%>checked<%end if%> class="clearBorder"> Display choices using radio buttons&nbsp;&nbsp;
				<input type="radio" name="pcv_UseRadio<%=oCnt%>" value="1" <%if pcv_UseRadio="1" AND multiSelectVar<>"1" then%>checked<%end if%> class="clearBorder"> Display choices using drop down menus
				<input type="radio" name="pcv_UseRadio<%=oCnt%>" value="2" <%if multiSelectVar="1" then%>checked<%end if%> class="clearBorder"> Display choices using check boxes
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td colspan="7">Configuration instructions:<br />
				<textarea rows="3" name="Notes<%=oCnt%>" cols="100"><%=CATNotes%></textarea></td>
			</tr> 
			<tr>
				<td colspan="8"><hr></td>
			</tr>
			<tr>
				<td colspan="8">Selectable items within the category:</td>
			</tr>                           
			<tr class="pcSmallText">
				<td nowrap>ORDER</td>
				<td nowrap>DEFAULT</td>
				<td nowrap width="60%">ITEM NAME</td>
				<td nowrap align="center" colspan="2">USE REG. PRICE</td>
				<td align="center">PRICE</td>
				<td align="center">WHOLESALE</td>
				<td align="center">DEL</td>
			</tr>

			<% if multiSelectVar<>"1" then %>
			<tr> 
				<td>&nbsp;</td>
				<td align="right"><input type="radio" name="cdefault<%=tempCat1%>" value="XX" checked class="clearBorder"></td>
				<td colspan="6">No Default Item</td>
			</tr>
			<% end if %>
							
		<%
		end if 
		'multiSelectVar="0"
		showInfoVar="0"
		requiredCategoryVar="0"
		tempCat2 = tempCat1
		%>

		<tr> 
			<td align="center">
				<input type="text" name="prdSort<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>" size="1" maxlength="3" value="<%=strPrdSort%>">
			</td>
			<td align="right">
				<% if cdefault=True then %>          
					<input <%if multiSelectVar<>"1" then%>type="radio"<%else%>type="checkbox"<%end if%> name="cdefault<%=tempCat1%><%if multiSelectVar="1" then%>_<%=pCnt%><%end if%>" value="<%=pCnt%>" checked class="clearBorder">
				<% else %>
					<input <%if multiSelectVar<>"1" then%>type="radio"<%else%>type="checkbox"<%end if%> name="cdefault<%=tempCat1%><%if multiSelectVar="1" then%>_<%=pCnt%><%end if%>" value="<%=pCnt%>" class="clearBorder">
				<% end if %>
			</td>
			<td>  
				<a href="FindProductType.asp?id=<%=intConfigProduct%>" target="_blank"><%=strDesc %></a> <%=pcv_HideMsg%><% if intCCExists=1 then %>&nbsp;&nbsp;<a href="#" onClick="showlaptop('laptop<%=intConfigProduct%>');return false;"><img src="images/pc_expand.gif" alt="Show Customer Pricing Categories" width="16" height="16">Show</a>&nbsp;&nbsp;<% end if %>
			</td>
			<td width="5%"> 
				<input type="checkbox" name="rPrice<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>" value="<%=money(dblprice)%>" class="clearBorder">
			</td>
			<td width="18%" nowrap><%=scCurSign%> <%=money(dblprice)%></td>
			<td width="18%" nowrap> 
				<div align="center"><%=scCurSign%> <input type="text" name="price<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>" value="<%=money(dblcPrice)%>" size="4" maxlength="10" style="text-align: right; font-family: Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #000000">
				<input type="hidden" name="id<%=pCnt%>_<%=tempCat1%>" value="<%=intConfigProduct%>">
				</div>
			</td>
			<td width="18%" nowrap> 
				<div align="center"><%=scCurSign%> <input type="text" name="Wprice<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>" value="<%=money(dblWprice)%>" size="4" maxlength="10" style="text-align: right; font-family: Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #000000">
				</div>
			</td>
			<td width="6%"> 
				<div align="center">
				<%aCnt=aCnt+1%>
				<input type="checkbox" name="DP<%=aCnt%>" value="BTO_<%=intConfigProduct%>" class="clearBorder">
				</div>
			</td>
		</tr>
<%
			if intCCExists=1 then
			For intCC=0 to intCountCC
				query="SELECT pcCC_BTO_Pricing.idCC_BTO_Price, pcCC_BTO_Pricing.idcustomerCategory, pcCC_BTO_Pricing.idBTOProduct, pcCC_BTO_Pricing.idBTOItem, pcCC_BTO_Pricing.pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE (((pcCC_BTO_Pricing.idcustomerCategory)="&CCArray(0,intCC)&") AND ((pcCC_BTO_Pricing.idBTOProduct)="&pidProduct&") AND ((pcCC_BTO_Pricing.idBTOItem)="&intConfigProduct&"));"
				SET rsCCObj=Server.CreateObject("ADODB.RecordSet")
				SET rsCCObj=conntemp.execute(query) 
				if rsCCObj.eof then 
					query="SELECT pcCC_Pricing.idcustomerCategory, pcCC_Pricing.idProduct, pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE (((pcCC_Pricing.idcustomerCategory)="&CCArray(0,intCC)&") AND ((pcCC_Pricing.idProduct)="&intConfigProduct&"));"
					SET rsPriceObj=server.CreateObject("ADODB.RecordSet")
					SET rsPriceObj=conntemp.execute(query)
					if rsPriceObj.eof then
						idCC_BTO_Price=0
						pcCC_BTO_Price=0
						
						' Find out if there is a wholesale price
						if (dblWprice>"0") then
							tempPrice=dblWprice
						else
							tempPrice=dblprice
						end if
						
						' Calculate the "across the board" price
						if CCArray(2,intCC)="ATB" then
							if CCArray(4,intCC)="Retail" then
								pcCC_BTO_Price=dblprice-(pcf_Round(dblprice*(cdbl(CCArray(3,intCC))/100),2))
							else
								pcCC_BTO_Price=tempPrice-(pcf_Round(tempPrice*(cdbl(CCArray(3,intCC))/100),2))
							end if							
						end if					
					else
						idCC_BTO_Price=0					
						pcCC_BTO_Price=rsPriceObj("pcCC_Price")
						pcCC_BTO_Price=pcf_Round(pcCC_BTO_Price, 2)
					end if
					SET rsPriceObj=nothing
				else
					idCC_BTO_Price=rsCCObj("idCC_BTO_Price")
					pcCC_BTO_Price=rsCCObj("pcCC_BTO_Price")
				end if
				set rsCCObj=nothing
%>
				<tr id="laptop<%=intConfigProduct%>" style="display: none; background-color:#e5e5e5;">
					<td colspan="6" align="right"><%=CCArray(1,intCC)%>:</td>
					<td width="18%" nowrap align="center"> 
						<%=scCurSign%> 
						<input type="text" name="CCBTOprice<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>_<%=CCArray(0,intCC)%>" value="<%=money(pcCC_BTO_Price)%>" size="4" maxlength="10" style="text-align: right; font-family: Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #000000">
						<input type="hidden" name="idCCBTO<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>_<%=CCArray(0,intCC)%>" value="<%=idCC_BTO_Price%>">
					</td>
					<td>&nbsp;</td>
				</tr>
			<% next %>
				<tr id="laptop<%=intConfigProduct%>" style="display: none;">
					<td colspan="7" align="right"><a href="#" onClick="hidelaptop('laptop<%=intConfigProduct%>');return false;"><img src="images/pc_collapse.gif" alt="Hide Customer Pricing Categories">Hide</a></td>
					<td>&nbsp;</td>
				</tr>
			<% end if
			
			btm="1"
		Next
		set rs=nothing
		%>
		<input type="hidden" name="pCnt<%=oCnt%>" value="<%=pCnt%>">
		<% end if %>

	<% if showBtm="0" then%>
		<input type="hidden" name="oCnt" value="<%=oCnt%>">
		<tr> 
			<td colspan="8"><hr></td>
		</tr>
		<tr> 
			<td colspan="8" align="center">
			<input type="submit" name="Submit2" value="Update" class="btn btn-primary">&nbsp;
			<input type="button" class="btn btn-default"  value="Preview" onClick="window.open('../pc/configurePrd.asp?idproduct=<%=pidProduct%>&adminpreview=1')">&nbsp;
			<input type="button" class="btn btn-default"  value="Add New Category" onClick="location.href='AddBTOcat.asp?idProduct=<%=pidProduct%>'">
			&nbsp;
			<input type="button" class="btn btn-default"  value="Locate another Configurable Product" onClick="location.href='LocateProducts.asp?cptype=1'">
            &nbsp;
			<input type="submit" name="DelButton" value="Delete Selected" onClick="javascript:return(confirm('Are you sure you want to delele selected items?'));" class="btn btn-default">
					&nbsp;
            <input type="button" class="btn btn-default"  value="Back to the Top" onClick="location.href='#top'">
			</td>
		</tr>
		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>

			<%
			strSQL="SELECT categories.categoryDesc, configSpec_Charges.price AS cPrice, configSpec_Charges.Wprice, configSpec_Charges.cdefault, configSpec_Charges.prdSort, configSpec_Charges.catSort, configSpec_Charges.requiredCategory, configSpec_Charges.multiSelect,categories.idCategory, configSpec_Charges.configProduct, products.price AS price, configSpec_Charges.configProductCategory, products.description, configSpec_Charges.showInfo, configSpec_Charges.displayQF, configSpec_Charges.pcConfCha_ShowDesc,configSpec_Charges.pcConfCha_ShowImg,configSpec_Charges.pcConfCha_ImgWidth,configSpec_Charges.pcConfCha_ShowSKU,products.active,products.removed,products.stock,products.nostock,products.pcProd_BackOrder,configSpec_Charges.pcConfCha_UseRadio,configSpec_Charges.Notes FROM (configSpec_Charges INNER JOIN categories ON configSpec_Charges.configProductCategory = categories.idCategory) INNER JOIN products ON configSpec_Charges.configProduct = products.idProduct WHERE (((configSpec_Charges.specProduct)="&pidProduct&")) ORDER BY configSpec_Charges.catSort, categories.idCategory, configSpec_Charges.prdSort, products.description;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(strSQL)
			showChg="0"
			if rs.eof then
				showChg="1"
				set rs=nothing%>
				<tr> 
					<td colspan="8"><a name="addcharges"></a><hr></td>
				</tr>
				<tr> 
					<td colspan="8">No Additional Charges - <a href="AddBTOchg.asp?idProduct=<%=pidProduct%>">Add New</a></td>
				</tr>
			<%end if%>
			<% if showChg="0" then%>
				<tr>
					<td colspan="8" class="pcCPspacer"><a name="addcharges"></a></td>
				</tr>
				<tr class="normal"> 
					<th colspan="8">Additional Charges - <a href="AddBTOchg.asp?idProduct=<%=pidProduct%>">Add more</a></th>
				</tr>
				<% oCnt = 0
				pCnt = 0
				btm="0"
				pcArr=rs.getRows()
				intCount=ubound(pcArr,2)
				set rs=nothing
				For i=0 to intCount
					strCategoryDesc=pcArr(0,i)
					dblcPrice=pcArr(1,i)
					dblWprice=pcArr(2,i)
					cdefault=pcArr(3,i)
					strPrdSort=pcArr(4,i)
					strCatSort=pcArr(5,i)
					requiredCategory=pcArr(6,i)
					multiSelect=pcArr(7,i)
					tempCat1 = pcArr(8,i)
					intConfigProduct=pcArr(9,i)
					dblprice=pcArr(10,i)
					configProductCategory=pcArr(11,i)
					strDesc=pcArr(12,i)
					pcv_active=pcArr(19,i)
					pcv_removed=pcArr(20,i)
					pcv_stock=pcArr(21,i)
					pcv_nostock=pcArr(22,i)
					pcv_backorder=pcArr(23,i)
					pcv_HideMsg=""
					if pcv_removed<>"0" then
						pcv_HideMsg="(Hidden - Removed)"
					end if
					if pcv_HideMsg="" then
						if pcv_active="0" then
							pcv_HideMsg="(Hidden - Inactive)"
						end if
					end if
					if pcv_HideMsg="" then
						if (scOutOfStockPurchase="-1") AND (iBTOOutofStockPurchase="-1") then
							if (clng(pcv_stock)<=0) AND (pcv_nostock="0") AND (pcv_backorder="0") then
								pcv_HideMsg="(Hidden - Out of Stock)"
							end if
						end if
					end if
					if pcv_HideMsg<>"" then
						pcv_HideMsg="<b><i>" & pcv_HideMsg & "</i></b>"
					end if
					pCnt = pCnt+1

					if Cint(tempCat2) <> Cint(tempCat1) then 
						if btm="1" then %>
							<input type="hidden" name="CHGpCnt<%=oCnt%>" value="<%=pCnt-1%>">
							<% end if
							oCnt = oCnt + 1
							pCnt = 1
							%>
							<tr>
								<td colspan="8">&nbsp;</td>
							</tr>
							<% 'check for multselect and required
							multiSelectVar="0"
							showInfoVar="0"
							requiredCategoryVar="0"
							pcv_AShowDesc="0"
							pcv_AShowImg="0"
							pcv_AImgWidth="35"
							pcv_AShowSKU="0"
							pcv_UseRadio="0"
								if multiSelect=True then
									multiSelectVar="1"
								end if
								if requiredCategory=True then
									requiredCategoryVar="1"
								end if
								if pcArr(13,i)=True then
									showInfoVar="1"
								end if
								if (pcArr(14,i)=-1) or (pcArr(14,i)=1) then
									displayQF="1"
								else
									displayQF="0"	
								end if
								if pcArr(15,i)="1" then
									pcv_AShowDesc="1"
								end if
								if pcArr(16,i)="1" then
									pcv_AShowImg="1"
								end if
								if pcArr(17,i)>"0" then
									pcv_AImgWidth=pcArr(17,i)
								end if
								if pcArr(18,i)="1" then
									pcv_AShowSKU="1"
								end if
								if pcArr(24,i)="1" then
									pcv_UseRadio="1"
								end if
								CHCATNotes=pcArr(25,i)
							%>
							<tr bgcolor="#e5e5e5"> 
								<td colspan="6"> 
									<span style="font-size:14px; font-weight: bold;"><%=strCategoryDesc%></span>
									<input type="hidden" name="CHGCATID<%=oCnt%>" value="<%=tempCat1%>">
									&nbsp;&nbsp;
									<a href="AddBTOchg1.asp?catSort=<%=strCatSort%>&idProduct=<%=pidProduct%>&idCategory=<%=tempCat1%>">Add more items</a>
									&nbsp;|&nbsp;
									<a href="javascript:if (confirm('You are about to remove this category from the product configuration. Are you sure you want to complete this action?')) location='DelBTOCat.asp?BTOproduct=<%=pidProduct%>&idCategory=<%=tempCat1%>&CatType=2';">Remove</a>
								</td>
								<td colspan="2" align="right">
									Order: <input type="text" name="CHGcatSort<%=oCnt%>" size="1" maxlength="3" value="<%=strCatSort%>" style="text-align: right; font-family: Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: bold; color: #000000; background-color: #99CCFF"> 
								</td>
							</tr>
									<tr> 
										<td>Settings:</td>
										<td colspan="7">
										<% if showInfoVar="1" then %>
											<input type="checkbox" name="CHGshowInfo<%=oCnt%>" value="1" checked class="clearBorder">
										<% else %>
											<input type="checkbox" name="CHGshowInfo<%=oCnt%>" value="1" class="clearBorder">
										<% end if %>
										Show Details&nbsp;&nbsp; 
										<% if requiredCategoryVar="1" then %>
											<input type="checkbox" name="CHGrequiredCategory<%=oCnt%>" value="1" checked class="clearBorder">
										<% else %>
											<input type="checkbox" name="CHGrequiredCategory<%=oCnt%>" value="1" class="clearBorder">
										<% end if %>
										Required Category&nbsp;&nbsp;

										<input type="hidden" name="CHGdisplayQF<%=oCnt%>" value="0">
										</td>
									</tr>
									<tr>
										<td></td>
										<td colspan="7">
											<%if pcv_AShowDesc="1" then %>
												<input type="checkbox" name="CHGpcv_AShowDesc<%=oCnt%>" value="1" checked class="clearBorder">
											<% else %>
												<input type="checkbox" name="CHGpcv_AShowDesc<%=oCnt%>" value="1" class="clearBorder">
											<% end if %>
											Show Product Description&nbsp;
											<%if pcv_AShowImg="1" then %>
												<input type="checkbox" name="CHGpcv_AShowImg<%=oCnt%>" value="1" checked class="clearBorder">
											<% else %>
												<input type="checkbox" name="CHGpcv_AShowImg<%=oCnt%>" value="1" class="clearBorder">
											<% end if %>
											Show Item Image&nbsp;&nbsp;-&nbsp;&nbsp;
											Image Width = <input type="text" name="CHGpcv_AImgWidth<%=oCnt%>" value="<%=pcv_AImgWidth%>" size="4">&nbsp;&nbsp;
											<%if pcv_AShowSKU="1" then %>
												<input type="checkbox" name="CHGpcv_AShowSKU<%=oCnt%>" value="1" checked class="clearBorder">
											<% else %>
												<input type="checkbox" name="CHGpcv_AShowSKU<%=oCnt%>" value="1" class="clearBorder">
											<% end if %>
											Show Item SKU<br>
											<input type="radio" name="CHGpcv_UseRadio<%=oCnt%>" value="0" <%if pcv_UseRadio<>"1" AND multiSelectVar<>"1" then%>checked<%end if%> class="clearBorder"> Display choices using radio buttons&nbsp;&nbsp;
											<input type="radio" name="CHGpcv_UseRadio<%=oCnt%>" value="1" <%if pcv_UseRadio="1" AND multiSelectVar<>"1" then%>checked<%end if%> class="clearBorder"> Display choices using drop down menus
											<input type="radio" name="CHGpcv_UseRadio<%=oCnt%>" value="2" <%if multiSelectVar="1" then%>checked<%end if%> class="clearBorder"> Display choices using check boxes
										</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td colspan="7">Configuration instructions:<br>
										<textarea rows="3" name="CHNotes<%=oCnt%>" cols="100"><%=CHCATNotes%></textarea></td>
									</tr>
									<tr>
										<td colspan="8"><hr></td>
									</tr>
									<tr class="pcSmallText">
										<td nowrap>ORDER</td>
										<td nowrap>DEFAULT</td>
										<td nowrap width="60%">ITEM NAME</td>
										<td nowrap align="center" colspan="2">USE REG. PRICE</td>
										<td align="center">PRICE</td>
										<td align="center">WHOLESALE</td>
										<td align="center">DEL</td>
									</tr>
									<tr class="small" style="font-family: Arial, Helvetica, sans-serif; font-size: 9pt; font-weight: normal; color: #000000"> 
										<td>&nbsp;</td>
										<td align="right"><input type="radio" name="CHGcdefault<%=tempCat1%>" value="XX" checked class="clearBorder"></td>
										<td colspan="6">No Default Item</td>
									</tr>
								<% end if 
								multiSelectVar="0"
								showInfoVar="0"
								requiredCategoryVar="0"
								tempCat2 = tempCat1 %>
								<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 9pt; font-weight: normal; color: #000000"> 
									<td align="center">
									<input type="text" name="CHGprdSort<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>" size="1" maxlength="3" value="<%=strPrdSort%>" style="text-align: right; font-family: Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #000000"></td>
									<td align="right">                  
									<% if cdefault=True then %>          
										<input type="radio" name="CHGcdefault<%=tempCat1%>" value="<%=pCnt%>" checked class="clearBorder">        
									<% else %>
										<input type="radio" name="CHGcdefault<%=tempCat1%>" value="<%=pCnt%>" class="clearBorder">
									<% end if %>
									</td>
									<td>  
									<a href="FindProductType.asp?id=<%=intConfigProduct%>" target="_blank"><%=strDesc%></a> <%=pcv_HideMsg%>
									<% if intCCExists=1 then %>
									&nbsp;&nbsp;<a href="#" onClick="showlaptop('laptop<%=intConfigProduct%>');return false;"><img src="images/pc_expand.gif" alt="Show Customer Pricing Categories" width="16" height="16">Show</a>&nbsp;&nbsp;
									<% end if %></td>
									<td width="5%"> 
									<input type="checkbox" name="CHGrPrice<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>" value="<%=dblprice%>" class="clearBorder">
									</td>
									<td width="18%" nowrap><div align="right"><%=scCurSign%><%=money(dblprice)%></div></td>
									<td width="18%" nowrap> 
										<div align="center"><%=scCurSign%> 
										<input type="text" name="CHGprice<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>" value="<%=money(dblcPrice)%>" size="4" maxlength="10" style="text-align: right; font-family: Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #000000">
										<input type="hidden" name="CHGid<%=pCnt%>_<%=tempCat1%>" value="<%=intConfigProduct%>">
										</div>
									</td>
									<td width="12%" nowrap> 
										<div align="center"><%=scCurSign%> 
										<input type="text" name="CHGWprice<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>" value="<%=money(dblWprice)%>" size="4" maxlength="10" style="text-align: right; font-family: Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #000000">
										</div>
									</td>
									<td width="6%"> 
									<div align="center">
									<%aCnt=aCnt+1%>
									<input type="checkbox" name="DP<%=aCnt%>" value="CHG_<%=intConfigProduct%>" class="clearBorder">
									</div>
									</td>
								</tr>
								
							<% if intCCExists=1 then
								For intCC=0 to intCountCC
									query="SELECT pcCC_BTO_Pricing.idCC_BTO_Price, pcCC_BTO_Pricing.idcustomerCategory, pcCC_BTO_Pricing.idBTOProduct, pcCC_BTO_Pricing.idBTOItem, pcCC_BTO_Pricing.pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE (((pcCC_BTO_Pricing.idcustomerCategory)="&CCArray(0,intCC)&") AND ((pcCC_BTO_Pricing.idBTOProduct)="&pidProduct&") AND ((pcCC_BTO_Pricing.idBTOItem)="&intConfigProduct&"));"	
									
									SET rsCCObj=Server.CreateObject("ADODB.RecordSet")
									SET rsCCObj=conntemp.execute(query) 
									if rsCCObj.eof then 
										query="SELECT pcCC_Pricing.idcustomerCategory, pcCC_Pricing.idProduct, pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE (((pcCC_Pricing.idcustomerCategory)="&CCArray(0,intCC)&") AND ((pcCC_Pricing.idProduct)="&intConfigProduct&"));"
										SET rsPriceObj=server.CreateObject("ADODB.RecordSet")
										SET rsPriceObj=conntemp.execute(query)
										if rsPriceObj.eof then
											idCC_BTO_Price=0
											pcCC_BTO_Price=0
											
											' Find out if there is a wholesale price
											if (dblWprice>"0") then
												tempPrice=dblWprice
											else
												tempPrice=dblprice
											end if
											
											' Calculate the "across the board" price
											if CCArray(2,intCC)="ATB" then
												if CCArray(4,intCC)="Retail" then
													pcCC_BTO_Price=dblprice-(pcf_Round(dblprice*(cdbl(CCArray(3,intCC))/100),2))
												else
													pcCC_BTO_Price=tempPrice-(pcf_Round(tempPrice*(cdbl(CCArray(3,intCC))/100),2))
												end if												
											end if					
										else
											idCC_BTO_Price=0					
											pcCC_BTO_Price=rsPriceObj("pcCC_Price")
											pcCC_BTO_Price=pcf_Round(pcCC_BTO_Price, 2)
										end if
										SET rsPriceObj=nothing
									else
										idCC_BTO_Price=rsCCObj("idCC_BTO_Price")
										pcCC_BTO_Price=rsCCObj("pcCC_BTO_Price")
									end if
									set rsCCObj=nothing
%>
									
									<tr id="laptop<%=intConfigProduct%>" style="display: none;">
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										<td colspan="4"><div align="right"><%=CCArray(1,intCC)%>:</div></td>
										<td width="18%" nowrap> 
											<div align="center"><%=scCurSign%> 
											<input type="text" name="CHGCCBTOprice<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>_<%=CCArray(0,intCC)%>" value="<%=money(pcCC_BTO_Price)%>" size="4" maxlength="10" style="text-align: right; font-family: Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #000000">
											<input type="hidden" name="CHGidCCBTO<%=pCnt%>_<%=intConfigProduct%>_<%=tempCat1%>_<%=CCArray(0,intCC)%>" value="<%=idCC_BTO_Price%>">
											</div></td>
										<td>&nbsp;</td>
									</tr>
								<% next %>
								
								<tr id="laptop<%=intConfigProduct%>" style="display: none;">
									  <td>&nbsp;</td>
									  <td>&nbsp;</td>
									  <td colspan="4">&nbsp;</td>
									  <td nowrap><div align="center"><a href="#" onClick="hidelaptop('laptop<%=intConfigProduct%>');return false;"><img src="images/pc_collapse.gif" alt="Hide Customer Pricing Categories">Hide</a></div></td>
									  <td>&nbsp;</td>
							  </tr>

							<% end if %>
								<% btm="1"
							Next 
							set rs=nothing
							 %>
							<input type="hidden" name="CHGpCnt<%=oCnt%>" value="<%=pCnt%>">
						<% end if %>

			<% if showChg="0" then%>
			<input type="hidden" name="CHGoCnt" value="<%=oCnt%>">
			<input type="hidden" name="CHGidProduct" value="<%=pidProduct%>">
				<tr> 
					<td colspan="8"><hr align="center" noshade color="#e1e1e1" size="1"></td>
				</tr>
			<tr> 
				<td align="center" colspan="8">
					<input type="submit" name="Submit2" value="Update" class="btn btn-primary">&nbsp;
					<input type="button" class="btn btn-default"  value="Add New Charges" onClick="location.href='AddBTOchg.asp?idProduct=<%=pidProduct%>'">
                    &nbsp;
					<input type="button" class="btn btn-default"  value="Locate another Configurable Product" onClick="location.href='LocateProducts.asp?cptype=1'">
                    &nbsp;
					<input type="submit" name="DelButton1" value="Delete Selected" onClick="javascript:return(confirm('Are you sure you want to delele selected items?'));" class="btn btn-default">
					&nbsp;
                    <input type="button" class="btn btn-default"  value="Back to the Top" onClick="location.href='#top'">
				</td>
			</tr>
			<% end if %>
		<% end if %>
	</table>
	<input type="hidden" name="aCnt" value="<%=aCnt%>">
	<input type="hidden" name="idProduct" value="<%=pidProduct%>">
</form>
<!--#include file="Adminfooter.asp"-->