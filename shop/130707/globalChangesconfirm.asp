<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" --> 
<!--#include file="inc_UpdateDates.asp" -->
<% 
Dim rstemp1

if request("action")="update" then
	CP=request("CP1")
	'Start SDBA
	tmpquery=""
	if CP="9" then
		if request("pcIDDropshipper")<>"0" then
			tmpquery=",pcDropShippersSuppliers "
		end if
	end if
	'End SDBA
	if CP="2" then
		if request("idcategory")<>"0" then
			query="SELECT DISTINCT products.pcProd_Apparel,products.weight,products.idProduct,products.price,products.listprice,products.btoBprice,products.cost,products.pcprod_Addprice,products.pcprod_AddWprice,products.pcProd_ParentPrd FROM products, categories_products" & tmpquery & " WHERE products.removed=0 "
		else
			query="SELECT DISTINCT products.pcProd_Apparel,products.weight,products.idProduct,products.price,products.listprice,products.btoBprice,products.cost,products.pcprod_Addprice,products.pcprod_AddWprice,products.pcProd_ParentPrd FROM products" & tmpquery & " WHERE products.removed=0 "
		end if
	else
		if CP="1" then
			query="SELECT DISTINCT products.pcProd_Apparel,products.weight,products.idProduct,products.price,products.listprice,products.btoBprice,products.cost,products.pcprod_Addprice,products.pcprod_AddWprice,products.pcProd_ParentPrd FROM products" & tmpquery & " WHERE products.removed=0 "
		else
			query="SELECT DISTINCT products.pcProd_Apparel,products.weight,products.idProduct,products.price,products.listprice,products.btoBprice,products.cost,products.pcprod_Addprice,products.pcprod_AddWprice,products.pcProd_ParentPrd FROM products" & tmpquery & " WHERE products.removed=0 "
		end if
	end if
	
	if request("pcv_incSPs")="2" then
		query=query & " and products.pcProd_ParentPrd<>0 "
	else
		if ((request("priceSelect")="2") OR (request("priceSelect1")="2") OR (request("priceSelect2")="2")) then
		else
			query=query & " and products.pcProd_ParentPrd=0 "
		end if
	end if

	if CP="2" then
		if request("idcategory")<>"0" then
			idcategory=request("idcategory")
			if request("incSubCats")<>"1" then
				query=query & " AND categories_products.idCategory=" &idcategory & " AND products.idProduct=categories_products.idProduct"
			else
				query=query & " AND categories_products.idCategory IN (" & request("TmpCatList") & ") AND products.idProduct=categories_products.idProduct"
			end if
		end if
	end if
	
	if CP="3" then
		query=query & "AND products.sku like '%"&replace(request("sku"),"'","''")&"%'"
	end if
	
	if CP="4" then
		query=query & "AND ((products.description like '%"&replace(request("nd"),"'","''")&"%') OR (products.details like '%" &replace(request("nd"),"'","''")& "%'))"
	end if	
	
	if CP="5" then
		if request("hpType")="2" then
			query=query & "AND products.listprice>="&replacecomma(request("hprice"))
		else
			query=query & "AND products.price>="&replacecomma(request("hprice"))
		end if
	end if	

	if CP="6" then
		if request("lpType")="2" then
			query=query & "AND products.listprice<="&replacecomma(request("lprice"))
		else
			query=query & "AND products.price<="&replacecomma(request("lprice"))
		end if
	end if
	
	if CP="7" then
		query=query & "AND products.IDBrand="&request("IDBrand")
	end if
	
	'Start SDBA
	if CP="8" then
		if request("pcIDSupplier")>"0" then
			query=query & "AND products.pcSupplier_ID="&request("pcIDSupplier")
		end if
	end if
	if CP="9" then
		if request("pcIDDropshipper")<>"0" then
			pcArr=split(request("pcIDDropshipper"),"_")
			query=query & "AND products.pcDropShipper_ID=" & pcArr(0)
			query=query & "AND pcDropShippersSuppliers.idproduct=products.idproduct AND pcDropShippersSuppliers.pcDS_IsDropShipper=" & pcArr(1) & " AND products.pcDropShipper_ID=" & pcArr(0)
		end if
	end if
	'End SDBA
	
	if CP="10" then
		if request("pcv_instock")>"0" then
			query=query & "AND products.stock>0"
		else
			query=query & "AND products.stock<=0"
		end if
	end if
	
	if CP="11" then
		pcv_tmp1=0
		if (request("pcv_prdtype1")<>"") and (request("pcv_prdtype2")="") and (request("pcv_prdtype3")="") then
			query=query & " AND products.serviceSpec=0 AND products.configOnly=0"
			pcv_tmp1=1
		end if
		if (request("pcv_prdtype1")="") and (request("pcv_prdtype2")<>"") and (request("pcv_prdtype3")="") then
			query=query & " AND products.serviceSpec<>0"
			pcv_tmp1=2
		end if
		if (request("pcv_prdtype1")="") and (request("pcv_prdtype2")="") and (request("pcv_prdtype3")<>"") then
			query=query & " AND products.configOnly<>0"
			pcv_tmp1=3
		end if
		if (request("pcv_prdtype1")<>"") and (request("pcv_prdtype2")<>"") and (request("pcv_prdtype3")="") then
			query=query & " AND products.configOnly=0"
			pcv_tmp1=4
		end if
		if (request("pcv_prdtype1")<>"") and (request("pcv_prdtype2")="") and (request("pcv_prdtype3")<>"") then
			query=query & " AND products.serviceSpec=0"
			pcv_tmp1=5
		end if
		if (request("pcv_prdtype1")="") and (request("pcv_prdtype2")<>"") and (request("pcv_prdtype3")<>"") then
			query=query & " AND (products.serviceSpec<>0 OR products.configOnly<>0)"
			pcv_tmp1=6
		end if
		if (request("pcv_prdtype1")="") and (request("pcv_prdtype2")="") and (request("pcv_prdtype3")="") then
			pcv_tmp1=7
		end if
		if (request("pcv_prdtype1")<>"") and (request("pcv_prdtype2")<>"") and (request("pcv_prdtype3")<>"") then
			pcv_tmp1=8
		end if
		if (request("pcv_prdtype4")<>"") then
			if (request("pcv_prdtype5")<>"") then
				query=query & " AND ((products.Downloadable=1)"
			else
				query=query & " AND products.Downloadable=1"
			end if
			pcv_tmp1=9
		end if
		if (request("pcv_prdtype5")<>"") then
			if pcv_tmp1=9 then
				query=query & " OR (products.pcprod_GC=1))"
			else
				query=query & " AND products.pcprod_GC=1"
			end if
		end if		
	end if
	
	if CP="13" then
		query=query & "AND products.idproduct NOT IN (SELECT DISTINCT categories_products.idproduct FROM categories_products)"
	end if

	smSalesQuery=""
	'Prevent Products in a Current Sales from being modified
	smSalesQuery = " AND (products.idProduct NOT IN (SELECT pcSales_BackUp.idProduct FROM pcSales_BackUp WHERE pcSales_BackUp.idProduct = products.IdProduct))"
	
	UP=request("UP1")

	if ((UP="1") and (instr(request("priceSelect"),"CC_")>0)) OR ((UP="2") and (instr(request("priceSelect1"),"CC_")>0)) then
		query1=split(query,"FROM")
		query=query & smSalesQuery & " UNION (SELECT DISTINCT products.pcProd_Apparel,products.weight,products.idProduct,products.price,products.listprice,products.btoBprice,products.cost,products.pcprod_Addprice,products.pcprod_AddWprice,products.pcProd_ParentPrd FROM Products WHERE products.pcProd_ParentPrd IN (SELECT DISTINCT Products.idproduct FROM " & query1(1) & ") AND products.removed=0" & smSalesQuery & ")"
	else
		query=query & smSalesQuery
	end if

	set rstemp=connTemp.execute(query)
	
	count=0
	
	UP=request("UP1")
	
	Dim pcArrGC, intCountGC, m
	pcArrGC=rstemp.getRows()
	set rstemp=nothing
	intCountGC=ubound(pcArrGC,2)
	
	FOR m=0 TO intCountGC

	pcv_Apparel=pcArrGC(0,m)
	if pcv_Apparel<>"" then
	else
		pcv_Apparel=0
	end if

	pweight=pcArrGC(1,m)
	pidproduct=pcArrGC(2,m)
	pcv_Price=pcArrGC(3,m)
	pcv_ListPrice=pcArrGC(4,m)
	if (pcv_ListPrice="0") AND ((request("priceSelect")="2") OR (request("priceSelect1")="2") OR (request("priceSelect2")="2")) AND (pcArrGC(9,m)>"0") then
		query="SELECT ListPrice FROM Products WHERE idProduct=" & pcArrGC(9,m) &";"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			pcv_ListPrice=rsQ("ListPrice")
		end if
		set rsQ=nothing
	end if
	
	pcv_BtoBPrice=pcArrGC(5,m)
	pcv_Cost=pcArrGC(6,m)

	pcv_PcProdAddprice=pcArrGC(7,m)
	pcv_PcProdAddWprice=pcArrGC(8,m)
	pcv_fieldname="idproduct"
	pcv_tmpTable="products "
	
	pcv_coption=request("coption")
	if NOT isNumeric(pcv_coption) OR pcv_coption="" then
		pcv_coption=0
	end if
	pcv_roption=request("roption")
	if NOT isNumeric(pcv_roption) OR pcv_roption="" then
		pcv_roption=0
	end if
	pcv_numoptions=request("numoptions")
	if NOT isNumeric(pcv_numoptions) OR pcv_numoptions="" then
		pcv_numoptions=0
	end if
	pcv_stroptions=request("stroptions")
	if NOT isNumeric(pcv_stroptions) OR pcv_stroptions="" then
		pcv_stroptions=0
	end if

	if ((UP="3") and (int(pcv_coption)=14 or int(pcv_coption)=15)) OR ((UP="4") and (int(pcv_roption)=14 or int(pcv_roption)=15)) OR ((UP="7") and int(pcv_numoptions)=6) OR ((UP="11") and int(pcv_stroptions)>=1 and int(pcv_stroptions)<=8) then
		pcv_tmpTable="DProducts "
	end if

	if ((UP="3") and (int(pcv_coption)>=17 and int(pcv_coption)<=22)) OR ((UP="4") and int(pcv_roption)=17) OR ((UP="7") and int(pcv_numoptions)=7) OR ((UP="11") and int(pcv_stroptions)>=9 and int(pcv_stroptions)<=10) then
		pcv_tmpTable="pcGC "
		pcv_fieldname="pcGC_IDProduct"
	end if

	if ((UP="1") and (instr(request("priceSelect"),"CC_")>0)) OR ((UP="2") and (instr(request("priceSelect1"),"CC_")>0)) then
		pcv_tmpTable="pcCC_Pricing "
		if instr(request("priceSelect"),"CC_")>0 then
			tmp_Arr=split(request("priceSelect"),"CC_")
		else
			tmp_Arr=split(request("priceSelect1"),"CC_")
		end if
		pcv_fieldname="idcustomerCategory=" & tmp_Arr(1) & " AND idProduct"
	end if
		
	query1="update " & pcv_tmpTable
	query3=" where " & pcv_fieldname & "=" & pidproduct

	if pcv_tmpTable="products " then
		
		query4=" where pcprod_ParentPrd=" & pidproduct & " AND removed=0 AND active=0" 
	
		if (request("pcv_incSPs")="1") AND (int(pcv_coption)<>6) AND (int(pcv_roption)<>6) AND UP<>"12" then
			if ((request("priceSelect")="2") OR (request("priceSelect1")="2") OR (request("priceSelect2")="2")) then
			else
			query3=query3 & " or pcProd_ParentPrd=" & pidproduct
			end if
		end if

	end if

	if UP="1" then
		priceSelect=request("priceSelect")
		if priceSelect="1" then
			tempPrice=cdbl(pcv_Price)
			if request("cpriceType")="1" then
				tempPrice=tempPrice+cdbl(tempPrice*cdbl(replacecomma(request("cprice")))*0.01)
			else
				if request("cpriceType")="2" then
					tempPrice=tempPrice+cdbl(replacecomma(request("cprice")))
				end if
			end if
			if request("cpriceRound")="1" then
				tempPrice=round(tempPrice)
			else
				if request("cpriceRound")="2" then
					tempPrice=round(tempPrice,2)
				end if
			end if
			query2 ="set price=" & tempPrice
		end if
		if priceSelect="2" then
			tempPrice=cdbl(pcv_ListPrice)
			if request("cpriceType")="1" then
				tempPrice=tempPrice+cdbl(tempPrice*cdbl(replacecomma(request("cprice")))*0.01)
			else
				if request("cpriceType")="2" then
					tempPrice=tempPrice+cdbl(replacecomma(request("cprice")))
				end if
			end if
			if request("cpriceRound")="1" then
				tempPrice=round(tempPrice)
			else
				if request("cpriceRound")="2" then
					tempPrice=round(tempPrice,2)
				end if
			end if
			query2 ="set listprice=" & tempPrice
		end if
		if priceSelect="3" then
			tempPrice=cdbl(pcv_BtoBPrice)
			if request("cpriceType")="1" then
				tempPrice=tempPrice+cdbl(tempPrice*cdbl(replacecomma(request("cprice")))*0.01)
			else
				if request("cpriceType")="2" then
					tempPrice=tempPrice+cdbl(replacecomma(request("cprice")))
				end if
			end if
			if request("cpriceRound")="1" then
				tempPrice=round(tempPrice)
			else
				if request("cpriceRound")="2" then
					tempPrice=round(tempPrice,2)
				end if
			end if
			query2 ="set btoBPrice=" & tempPrice
		end if
	
		if priceSelect="4" then
			tempPrice=cdbl(pcv_Price)
			tempAddPrice=cdbl(pcv_PcProdAddprice)
			tempPrice=tempPrice-tempAddPrice
			if request("cpriceType")="1" then
				tempAddPrice=tempAddPrice+cdbl(tempAddPrice*cdbl(replacecomma(request("cprice")))*0.01)
			else
				if request("cpriceType")="2" then
					tempAddPrice=tempAddPrice+cdbl(replacecomma(request("cprice")))
				end if
			end if
			tempPrice=tempPrice+tempAddPrice
			if request("cpriceRound")="1" then
				tempAddPrice=round(tempAddPrice)
				tempPrice=round(tempPrice)
			end if
			query2 ="set pcprod_Addprice=" & tempAddPrice & ",price=" & tempPrice
		end if

		if priceSelect="5" then
			tempPrice=cdbl(pcv_BtoBPrice)
			if tempPrice=0 then
				tempPrice=cdbl(pcv_Price)
			end if
			tempAddWPrice=cdbl(pcv_PcProdAddWprice)
			tempPrice=tempPrice-tempAddWPrice
			if request("cpriceType")="1" then
				tempAddWPrice=tempAddWPrice+cdbl(tempAddWPrice*cdbl(replacecomma(request("cprice")))*0.01)
			else
				if request("cpriceType")="2" then
					tempAddWPrice=tempAddWPrice+cdbl(replacecomma(request("cprice")))
				end if
			end if
			tempPrice=tempPrice+tempAddWPrice
			if request("cpriceRound")="1" then
				tempAddWPrice=round(tempAddWPrice)
				tempPrice=round(tempPrice)
			end if
			query2 ="set pcprod_AddWprice=" & tempAddWPrice & ",btoBPrice=" & tempPrice
		end if

		if instr(priceSelect,"CC_")>0 then
			tmp_Arr=split(priceSelect,"CC_")
			tempPrice=0
			tmp_BTOTable=0
			query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idproduct=" & pidproduct & " AND idcustomerCategory=" & tmp_Arr(1)
			set rstemp1=connTemp.execute(query)
			if not rstemp1.eof then
				tempPrice=rstemp1("pcCC_Price")
				tempPrice=pcf_Round(tempPrice, 2)
				if IsNull(tempPrice) or tempPrice="" then
					tempPrice=0
				end if
			else
				query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idBTOItem=" & pidproduct & " AND idcustomerCategory=" & tmp_Arr(1)
				set rstemp1=connTemp.execute(query)
				if not rstemp1.eof then
					tempPrice=rstemp1("pcCC_BTO_Price")
					if IsNull(tempPrice) or tempPrice="" then
						tempPrice=0
					end if
					tmp_BTOTable=1
				end if
			end if
			set rstemp1=nothing
			
			if tempPrice<>"0" then
			else
					query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory=" & tmp_Arr(1)
					SET rstemp1=Server.CreateObject("ADODB.RecordSet")
					SET rstemp1=conntemp.execute(query)
					if NOT rstemp1.eof then 
						intIdcustomerCategory=rstemp1("idcustomerCategory")
						strpcCC_Name=rstemp1("pcCC_Name")
						strpcCC_CategoryType=rstemp1("pcCC_CategoryType")
						intpcCC_ATBPercentage=rstemp1("pcCC_ATB_Percentage")
						intpcCC_ATB_Off=rstemp1("pcCC_ATB_Off")
						if intpcCC_ATB_Off="Retail" then
							intpcCC_ATBPercentOff=0
						else
							intpcCC_ATBPercentOff=1
						end if
						
						SP_price=pcv_Price
						SP_wprice=pcv_BtoBPrice
		
						if (SP_wprice>"0") then
							SPtempPrice=SP_wprice
						else
							SPtempPrice=SP_price
						end if
						' Calculate the "across the board" price
						if strpcCC_CategoryType="ATB" then
							if intpcCC_ATBPercentOff=0 then
								tempPrice=SP_price-(pcf_Round(SP_price*(cdbl(intpcCC_ATBPercentage)/100),2))
							else
								tempPrice=SPtempPrice-(pcf_Round(SPtempPrice*(cdbl(intpcCC_ATBPercentage)/100),2))
							end if
						end if
					end if
			end if
			
			'if tempPrice<>"0" then
				if request("cpriceType")="1" then
					tempPrice=tempPrice+cdbl(tempPrice*cdbl(replacecomma(request("cprice")))*0.01)
				else
					if request("cpriceType")="2" then
						tempPrice=tempPrice+cdbl(replacecomma(request("cprice")))
					end if
				end if
				if request("cpriceRound")="1" then
					tempPrice=round(tempPrice)
				else
					if request("cpriceRound")="2" then
						tempPrice=round(tempPrice,2)
					end if
				end if
				if tmp_BTOTable=1 then
					query2="UPDATE pcCC_BTO_Pricing SET pcCC_BTO_Price=" & tempPrice & " WHERE idBTOItem=" & pidproduct & " AND idcustomerCategory=" & tmp_Arr(1)
					set rstemp1=connTemp.execute(query2)
					set rstemp1=nothing
					query2 =""
				end if
				query2="SELECT idproduct FROM pcCC_Pricing WHERE idproduct=" & pidproduct & " AND idcustomerCategory=" & tmp_Arr(1)
				set rstemp1=connTemp.execute(query2)
				if rstemp1.eof then
					query2="INSERT INTO pcCC_Pricing (idproduct,idcustomerCategory,pcCC_Price) VALUES (" & pidproduct & "," & tmp_Arr(1) & "," & tempPrice & ");"
					set rstemp1=connTemp.execute(query2)
					query2 =""
				else
					query2 ="set pcCC_Price=" & tempPrice
				end if
			'else
			'	query2 ="set idproduct=idproduct"
			'end if
		end if				
	end if
	if UP="2" then
		priceSelect1=request("priceSelect1")
		priceSelect2=request("priceSelect2")
		tempPrice=0
		tmp_BTOTable=0
		if instr(priceSelect2,"CC_")>0 then
			tmp_Arr=split(priceSelect2,"CC_")
			query="SELECT pcCC_Price FROM pcCC_Pricing WHERE idproduct=" & pidproduct & " AND idcustomerCategory=" & tmp_Arr(1)
			set rstemp1=connTemp.execute(query)
			if not rstemp1.eof then
				tempPrice=rstemp1("pcCC_Price")
				tempPrice=pcf_Round(tempPrice, 2)
				if IsNull(tempPrice) or tempPrice="" then
					tempPrice=0
				end if
			else
				query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idBTOItem=" & pidproduct & " AND idcustomerCategory=" & tmp_Arr(1)
				set rstemp1=connTemp.execute(query)
				if not rstemp1.eof then
					tempPrice=rstemp1("pcCC_BTO_Price")
					if IsNull(tempPrice) or tempPrice="" then
						tempPrice=0
					end if
					tmp_BTOTable=1
				end if
			end if
			set rstemp1=nothing
		end if	
		if priceSelect1="1" then
		Select Case priceSelect2
		Case "1": tempPrice=cdbl(pcv_Price)
		Case "2": tempPrice=cdbl(pcv_ListPrice)
		Case "3": tempPrice=cdbl(pcv_BtoBPrice)
		Case "4": tempPrice=cdbl(pcv_Cost)
		End Select
			if ((priceSelect2="4") and (cdbl(tempPrice)<>0)) or (priceSelect2<>"4") then
				tempPrice=tempPrice*cdbl(replacecomma(request("wprice")))*0.01
				if request("cpriceRound1")="1" then
					tempPrice=round(tempPrice)
				else
					if request("cpriceRound1")="2" then
						tempPrice=round(tempPrice,2)
					end if
				end if
				query2 ="set Price=" & tempPrice
			else
				query2 ="set Price=Price"
			end if
		end if
		if priceSelect1="2" then
		Select Case priceSelect2
		Case "1": tempPrice=cdbl(pcv_Price)
		Case "2": tempPrice=cdbl(pcv_ListPrice)
		Case "3": tempPrice=cdbl(pcv_BtoBPrice)
		Case "4": tempPrice=cdbl(pcv_Cost)
		End Select
			if ((priceSelect2="4") and (cdbl(tempPrice)<>0)) or (priceSelect2<>"4") then
				tempPrice=tempPrice*cdbl(replacecomma(request("wprice")))*0.01
				if request("cpriceRound1")="1" then
					tempPrice=round(tempPrice)
				else
					if request("cpriceRound1")="2" then
						tempPrice=round(tempPrice,2)
					end if
				end if
				query2 ="set listPrice=" & tempPrice
			else
				query2 ="set listPrice=listPrice"
			end if
		end if
		if priceSelect1="3" then
		Select Case priceSelect2
		Case "1": tempPrice=cdbl(pcv_Price)
		Case "2": tempPrice=cdbl(pcv_ListPrice)
		Case "3": tempPrice=cdbl(pcv_BtoBPrice)
		Case "4": tempPrice=cdbl(pcv_Cost)
		End Select
			if ((priceSelect2="4") and (cdbl(tempPrice)<>0)) or (priceSelect2<>"4") then
				tempPrice=tempPrice*cdbl(replacecomma(request("wprice")))*0.01
				if request("cpriceRound1")="1" then
					tempPrice=round(tempPrice)
				else
					if request("cpriceRound1")="2" then
						tempPrice=round(tempPrice,2)
					end if
				end if
				query2 ="set btoBPrice=" & tempPrice
			else
				query2 ="set btoBPrice=btoBPrice"
			end if
		end if
	
		if priceSelect1="4" then
			Select Case priceSelect2
				Case "1": tempAddPrice=cdbl(pcv_Price)
				Case "3": tempAddPrice=cdbl(pcv_BtoBPrice)
				Case "4": tempAddPrice=cdbl(pcv_Cost)
				Case "5": tempAddPrice=cdbl(pcv_PcProdAddprice)
				Case "6": tempAddPrice=cdbl(pcv_PcProdAddWprice)
			End Select
			if ((priceSelect2="4") and (cdbl(tempPrice)<>0)) or (priceSelect2<>"4") then
				tempPrice=cdbl(pcv_Price)
				tempPrice=tempPrice-cdbl(pcv_PcProdAddprice)
				tempAddPrice=tempAddPrice*cdbl(replacecomma(request("wprice")))*0.01
				tempPrice=tempPrice+tempAddPrice
				if request("cpriceRound")="1" then
					tempAddPrice=round(tempAddPrice)
					tempPrice=round(tempPrice)
				end if
				query2 ="set pcprod_Addprice=" & tempAddPrice & ",price=" & tempPrice
			else
				query2 ="set pcprod_Addprice=pcprod_Addprice"
			end if
		end if

		if priceSelect1="5" then
			Select Case priceSelect2
				Case "1": tempAddPrice=cdbl(pcv_Price)
				Case "3": tempAddPrice=cdbl(pcv_BtoBPrice)
				Case "4": tempAddPrice=cdbl(pcv_Cost)
				Case "5": tempAddPrice=cdbl(pcv_PcProdAddprice)
				Case "6": tempAddPrice=cdbl(pcv_PcProdAddWprice)
			End Select
			if ((priceSelect2="4") and (cdbl(tempPrice)<>0)) or (priceSelect2<>"4") then
				tempPrice=cdbl(pcv_BtoBPrice)
				if tempPrice="" then
					tempPrice=cdbl(pcv_Price)
				end if
				tempPrice=tempPrice-cdbl(pcv_PcProdAddWprice)
				tempAddPrice=tempAddPrice*cdbl(replacecomma(request("wprice")))*0.01
				tempPrice=tempPrice+tempAddPrice
				if request("cpriceRound")="1" then
					tempAddPrice=round(tempAddPrice)
					tempPrice=round(tempPrice)
				end if
				query2 ="set pcprod_AddWprice=" & tempAddPrice & ",bToBprice=" & tempPrice
			else
				query2 ="set pcprod_AddWprice=pcprod_AddWprice"
			end if
		end if

		if instr(priceSelect1,"CC_")>0 then
			tmp_Arr1=split(priceSelect1,"CC_")
			Select Case priceSelect2
				Case "1": tempPrice=cdbl(pcv_Price)
				Case "2": tempPrice=cdbl(pcv_ListPrice)
				Case "3": tempPrice=cdbl(pcv_BtoBPrice)
				Case "4": tempPrice=cdbl(pcv_Cost)
			End Select
			'if cdbl(tempPrice)<>0 then
				tempPrice=tempPrice*cdbl(replacecomma(request("wprice")))*0.01
				if request("cpriceRound1")="1" then
					tempPrice=round(tempPrice)
				else
					if request("cpriceRound1")="2" then
						tempPrice=round(tempPrice,2)
					end if
				end if
				query2="SELECT idBTOItem FROM pcCC_BTO_Pricing WHERE idBTOItem=" & pidproduct & " AND idcustomerCategory=" & tmp_Arr1(1)
				set rstemp1=connTemp.execute(query2)
				if not rstemp1.eof then
					query2="UPDATE pcCC_BTO_Pricing SET pcCC_BTO_Price=" & tempPrice & " WHERE idBTOItem=" & pidproduct & " AND idcustomerCategory=" & tmp_Arr1(1)
					set rstemp1=connTemp.execute(query2)
					set rstemp1=nothing
					query2 =""
				end if
				query2="SELECT idproduct FROM pcCC_Pricing WHERE idproduct=" & pidproduct & " AND idcustomerCategory=" & tmp_Arr1(1)
				set rstemp1=connTemp.execute(query2)
				if rstemp1.eof then
					query2="INSERT INTO pcCC_Pricing (idproduct,idcustomerCategory,pcCC_Price) VALUES (" & pidproduct & "," & tmp_Arr1(1) & "," & tempPrice & ");"
					set rstemp1=connTemp.execute(query2)
					query2 =""
				else
					query2 ="set pcCC_Price=" & tempPrice
				end if
				set rstemp1=nothing
			'else
			'	query2 =""
			'end if
		end if
		
	end if
	if UP="3" then
		COption=pcv_coption
		Select Case COption
		Case "1":
			query2 ="set listHidden=-1"
		Case "2":
			query2 ="set hotDeal=-1"
		Case "3":
			query2 ="set notax=-1"
		Case "4":
			pEmailText=replace(request("nfsmsg"),"""","&quot;")
			pEmailText=replace(pEmailText,"'","''")
			query2 ="set formQuantity=-1,emailText='" & pEmailText & "'"
		Case "5":
			query2 ="set noshipping=-1"	
		Case "6":
			if request("pcv_incSPs")="2" then
				query2 ="set pcProd_SPInActive=0,active=0"
			else
				query2 ="set active=-1"
			end if
		Case "7":
			query2 ="set noStock=-1"						
		Case "8":
			query2 ="set noshippingtext=-1"
		Case "9":
			query2 ="set pcProd_BackOrder=1"
		Case "10":
			query2 ="set pcProd_NotifyStock=1"
		Case "11":
			query2 ="set pcProd_IsDropShipped=1"
		Case "12":
			pOverSizeSpec="NO"
			pOS_height=request("pcv_height")
			pOS_width=request("pcv_width")
			pOS_length=request("pcv_length")
			if pOS_length="" OR pOS_width="" OR pOS_length="" then
				pOverSizeSpec="NO"
			else
				pOS_girth=((pOS_width*2)+(pOS_height*2)+pOS_length)
				if pWeight<30 and pOS_girth<108 and pOS_girth>84 then
					pOSX=1
				else
					if pWeight<70 and pOS_girth>108 then
						pOSX=2
					else
						pOSX=0
					end if
				end if
				pOverSizeSpec= pOS_width&"||"&pOS_height&"||"&pOS_length&"||"&pOSX&"||"&pWeight
			end if
			query2 ="set OverSizeSpec='" & pOverSizeSpec & "' "
		Case "13":
			query2="set Downloadable=1 "
		Case "14":
			query2="set DProducts.URLExpire=1 "
		Case "15":
			query2="set DProducts.License=1 "
		Case "16":
			query2="set pcprod_GC=1 "
		Case "17":
			query2="set pcGC.pcGC_Exp=0 "
		Case "18":
			query2="set pcGC.pcGC_Exp=1 "
		Case "19":
			query2="set pcGC.pcGC_Exp=2 "
		Case "20":
			query2="set pcGC.pcGC_EOnly=1 "
		Case "21":
			query2="set pcGC.pcGC_CodeGen=0 "
		Case "22":
			query2="set pcGC.pcGC_CodeGen=1 "
		Case "23":
			query2="set pcprod_hidebtoprice=1 "
		Case "24":
			query2="set pcprod_HideDefConfig=1 "
		Case "25":
			query2="set NoPrices=1 "
		Case "26":
			query2="set NoPrices=2 "
		Case "27":
			query2="set pcProd_SkipDetailsPage=1 "
		Case "28":
			query2="set showInHome=-1 "
		Case "29":
			query2="set pcProd_HideSKU=1 "
		Case "30":
			query2="set pcPrd_MojoZoom=1 "
		Case "31":
			query2="set pcProd_ShowBTOCMMsg=1 "
		Case "32":
			query2="set pcProd_AdditionalImages=1 "
		End Select
	end if
	if UP="4" then
		ROption=pcv_roption
		Select Case ROption
		Case "1":
			query2 ="set listHidden=0"
		Case "2":
			query2 ="set hotDeal=0"
		Case "3":
			query2 ="set notax=0"
		Case "4":
			query2 ="set formQuantity=0,emailText=''"
		Case "5":
			query2 ="set noshipping=0"	
		Case "6":
			if request("pcv_incSPs")="2" then
				query2 ="set pcProd_SPInActive=1,active=0"
			else
				query2 ="set active=0"
			end if
		Case "7":
			query2 ="set noStock=0"						
		Case "8":
			query2 ="set noshippingtext=0"
		Case "9":
			query2 ="set pcProd_BackOrder=0"
		Case "10":
			query2 ="set pcProd_NotifyStock=0"
		Case "11":
			query2 ="set pcProd_IsDropShipped=0"
		Case "12":
			query2 ="set OverSizeSpec='NO' "
		Case "13":
			query2="set Downloadable=0 "
		Case "14":
			query2="set DProducts.URLExpire=0 "
		Case "15":
			query2="set DProducts.License=0 "
		Case "16":
			query2="set pcprod_GC=0 "
		Case "17":
			query2="set pcGC.pcGC_EOnly=0 "
		Case "18":
			query2="set pcprod_hidebtoprice=0 "
		Case "19":
			query2="set pcprod_HideDefConfig=0 "
		Case "20":
			query2="set NoPrices=0 "
		Case "21":
			query2="set pcProd_SkipDetailsPage=0 "
		Case "22":
			query2 ="set showInHome=0 "
		Case "23":
			query2 ="set pcProd_HideSKU=0 "
		Case "24":
			query2="set pcPrd_MojoZoom=0 "
		Case "31":
			query2="set pcProd_ShowBTOCMMsg=0 "
		Case "32":
			query2="set pcProd_AdditionalImages=0 "
		End Select
	end if
	if UP="5" then
		PTOption=request("ptoption")
		Select Case PTOption
		Case "1":
			query2 ="set configOnly=0, serviceSpec=0"
		Case "2":
			query2 ="set configOnly=0, serviceSpec=1"
		Case "3":
			if SQL_Format="1" then
				query2 ="set configOnly=-1, serviceSpec=0"
			else
				query2 ="set configOnly=1, serviceSpec=0"
			end if
		End Select
	end if
	if UP="6" then
		pWeight=request("weight")
		If pWeight="" then
			pWeight="0"
		End if		
		pWeight_oz=request("weight_oz")
		If pWeight_oz="" then
			pWeight_oz="0"
		End if
		pWeight=((Int(pWeight)*16)+Int(pWeight_oz))
		if scShipFromWeightUnit="KGS" then
			pWeight_kg=request("weight_kg")
			if pWeight_kg="" then
				pWeight_kg="0"
			end if
			pWeight_g=request("weight_g")
			if pWeight_g="" then
				pWeight_g="0"
			end if
			pWeight=((Int(pWeight_kg)*1000)+Int(pWeight_g))
		end if
		if pWeight="" then
			pWeight="0"
		End If
		weight_units=request("weight_units")
		if weight_units="" then
			weight_units="0"
		end if
		query2 ="set weight=" & pweight & ",pcprod_QtyToPound=" & weight_units
	end if
	ChangeStock=0
	if UP="7" then
		numvalue=request("numvalue")
		if numvalue="" then
			numvalue="0"
		end if
		Select Case pcv_numoptions
			Case "1": query2 ="set stock=" & numvalue
			ChangeStock=1
			Case "2": query2 ="set cost=" & replacecomma(numvalue)
			Case "3": query2 ="set pcProd_ReorderLevel=" & numvalue
			Case "4": query2 ="set pcProd_ShipNDays=" & numvalue
			Case "5": query2 ="set iRewardPoints=" & numvalue
			Case "6": query2 ="set DProducts.ExpireDays=" & numvalue
			Case "7": query2 ="set pcGC.pcGC_ExpDays=" & numvalue
		End Select
	end if
	if UP="8" then
		minimumqty=request("minimumqty")
		if minimumqty<>"" then
		else
		minimumqty="0"
		end if
		qtyvalidate=request("qtyvalidate")
		if qtyvalidate<>"" then
		else
			qtyvalidate="0"
		end if
		query2 ="set pcprod_minimumqty=" & minimumqty & ", pcprod_qtyvalidate=" & qtyvalidate & " "
		if qtyvalidate="1" then
			query2=query2 & ",pcProd_multiQty=" & minimumqty & " "
		end if
	end if
	if UP="9" then
		pcv_IDSupplier=request("pcToIDSupplier")
		if (pcv_IDSupplier<>"") and (pcv_IDSupplier<>"0") then
			query2="set pcSupplier_ID=" & pcv_IDSupplier
		else
			query2="set pcSupplier_ID=pcSupplier_ID"
		end if
	end if
	
	if UP="10" then
		if request("pcToIDDropshipper")<>"" then
			pcArr=split(request("pcToIDDropshipper"),"_")
			query2="set pcDropShipper_ID=" & pcArr(0)
			tmpquery="DELETE FROM pcDropShippersSuppliers WHERE idproduct=" & pidproduct
			set rstemp1=connTemp.execute(tmpquery)
			tmpquery="INSERT INTO pcDropShippersSuppliers (idproduct,pcDS_IsDropShipper) VALUES (" & pidproduct & "," & pcArr(1) & ");"
			set rstemp1=connTemp.execute(tmpquery)
			set rstemp1=nothing
		else
			query2="set pcDropShipper_ID=pcDropShipper_ID"
		end if
	end if
	if UP="11" then
		strvalue=request("strvalue")
		Select Case pcv_stroptions
			Case "1":
				query2="set DProducts.ProductURL='" & strvalue & "' "
			Case "2":
				query2="set DProducts.LocalLG='" & strvalue & "' "
			Case "3":
				query2="set DProducts.RemoteLG='" & strvalue & "' "
			Case "4":
				query2="set DProducts.LicenseLabel1='" & strvalue & "' "
			Case "5":
				query2="set DProducts.LicenseLabel2='" & strvalue & "' "
			Case "6":
				query2="set DProducts.LicenseLabel3='" & strvalue & "' "
			Case "7":
				query2="set DProducts.LicenseLabel4='" & strvalue & "' "
			Case "8":
				query2="set DProducts.LicenseLabel5='" & strvalue & "' "
			Case "9":
				query2="set pcGC.pcGC_GenFile='" & strvalue & "' "
			Case "10":
				query2="set pcGC.pcGC_ExpDate='" & strvalue & "' "
		End Select
	end if
	if UP="12" then
		pcv_IDBrand=request("pcToIDBrand")
		if pcv_IDBrand="" then
			pcv_IDBrand="0"
		end if
		query2="set IDBrand=" & pcv_IDBrand & " "
	end if
	
	if UP="13" then
		query2="set pcProd_DisplayLayout='" & request("pcv_displayLayout") & "' "
	end if
	
	if UP="14" then
		tmpquery="SELECT idproduct FROM categories_products WHERE idproduct=" & pidproduct & " AND idcategory=" & request("ToIDCategory")
		set rstemp1=connTemp.execute(tmpquery)
		if rstemp1.eof then
			tmpquery="INSERT INTO categories_products (idproduct,idcategory) VALUES (" & pidproduct & "," & request("ToIDCategory") & ");"
			set rstemp1=connTemp.execute(tmpquery)
		end if
		set rstemp1=nothing
		query2=""
	end if
	
	if UP="15" then
		Select Case request("goSett")
			Case "1": query2="set pcProd_GoogleCat='" & replace(request("goValue"),"'","''") & "' "
			Case "2": query2="set pcProd_GoogleGender='" & replace(request("goValue"),"'","''") & "' "
	        Case "3": query2="set pcProd_GoogleAge='" & replace(request("goValue"),"'","''") & "' "
			Case "4": query2="set pcProd_GoogleColor='" & replace(request("goValue"),"'","''") & "' "
			Case "5": query2="set pcProd_GoogleSize='" & replace(request("goValue"),"'","''") & "' "
			Case "6": query2="set pcProd_GooglePattern='" & replace(request("goValue"),"'","''") & "' "
			Case "7": query2="set pcProd_GoogleMaterial='" & replace(request("goValue"),"'","''") & "' "
		End Select
	end if
	
	query=query1 & query2 & query3
	if query2<>"" then
		set rstemp1=connTemp.execute(query)
	end if
	
	if err.number<>0 then
		err.number=0
		err.description=""
	else
		count=count+1
	end if
	
	'BackInStock-S
	if ChangeStock=1 then
		Call pcs_hookInStockEvent(pidproduct, "")
	end if
	'BackInStock-E
	
	call updPrdEditedDate(pidproduct)
	
	'// Update sub-products instead of category updates
	if (pcv_Apparel="1") AND (pcv_tmpTable="products ") AND (UP<>"14") then
		
		if ((request("priceSelect")="2") OR (request("priceSelect1")="2") OR (request("priceSelect2")="2")) then
		else
			if (query2<>"set active=-1") AND (query2<>"set active=0") then
				query=query1 & query2 & query4
				set rs1=connTemp.execute(query)
				set rs1=nothing
			else
				if request("pcv_incSPs")="1" then
					if (query2="set active=0") then
						query=query1 & "set pcProd_SPInActive=1,active=0" & query4
					else
						query=query1 & "set pcProd_SPInActive=0,active=0" & query4
					end if
					set rs1=connTemp.execute(query)
					set rs1=nothing
				end if
			end if
		end if
		
		'BackInStock-S
		if ChangeStock=1 then
			'Call pcs_hookInStockEvent(pidproduct, "")
		end if
		'BackInStock-E

		query="Select price,btoBPrice from Products where idproduct=" & pidproduct
		set rs1=connTemp.execute(query)
		pPrice=rs1("Price")
	
		if pPrice<>"" then
		else
			pPrice="0"
		end if

		pBToBPrice=rs1("bToBPrice")
		if (pBToBPrice<>"") and (pBToBPrice<>"0") then
		else
			pBToBPrice=pPrice
		end if

		query="Select idproduct,pcprod_addprice,pcprod_addWprice,pcprod_Relationship from Products where pcprod_ParentPrd=" & pidproduct & " AND removed=0 AND active=0"
		set rs1=connTemp.execute(query)

		do while not rs1.eof
			pcv_SIdproduct=rs1("idproduct")
			pcv_Addprice=rs1("pcprod_Addprice")
			if pcv_Addprice<>"" then
			else
				pcv_Addprice=0
			end if
	
			pcv_AddWprice=rs1("pcprod_AddWprice")
			if (pcv_AddWprice<>"") and (pcv_AddWprice<>"0") then
			else
				pcv_AddWprice="0"
			end if
	
			pcv_PPrice=cdbl(pPrice)+cdbl(pcv_Addprice)
	
			pcv_PWPrice=cdbl(pBToBPrice)+cdbl(pcv_AddWprice)

			query="update products set price=" & pcv_PPrice & ",btoBPrice=" & pcv_PWPrice & " where idproduct=" & pcv_SIdproduct
			set rs=connTemp.execute(query)
			set rs=nothing
			
			call pcs_hookProductModified(pcv_SIdproduct, "")
			
			rs1.MoveNext
		loop
		set rs1=nothing
	end if
	'// End update sub-products
	NEXT

set rstemp1= nothing

If statusAPP="1" OR scAPP=1 Then
	'// Update parent products inventory levels if necessary 
	%>
	<!--#include file="../pc/app-updstock.asp"-->
	<% 
End If

pageTitle="Global Product Changes"

if request("nav")="1" then
section="services"
else
section="products"
end if
%>
<!--#include file="AdminHeader.asp"-->
	<br>
	    <div class="pcCPmessageSuccess"><%=count%> products were successfully updated. <a href="globalChanges.asp?nav=<%=request("nav")%>">New Global Change</a></div>
        
        <% if scBTO=1 and (UP="1" or UP="2") then %>
			<div class="pcCPmessage"><strong>About Configurator Price Updates</strong><br><br>Please note that default prices are not updated when Global Changes are made. That's because the price of an item assigned to a configurable product may be different from its price as a stand-alone product, and different among different configurable products.<br><br>To update prices across multiple products, please use the following features:
            <ul>
            	<li><a href="updBTOPrdPrices.asp">Update Base Prices</a></li>
                <li><a href="updBTODefaultPrices.asp">Update Default Prices</a></li>
                <li><a href="updateBTOprices.asp">Update Configuration Prices</a></li>
            </ul>
            </div>
        <% end if %>
    <br>
<!--#include file="AdminFooter.asp"--> 
<%
End if
%>