<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="pcCalculateBTODefaultPrices.asp" -->
<!--#include file="inc_UpdateDates.asp" -->
<%
dim pageTitle, section, f
pageTitle="Product Updated"
section="products"

'// LOAD VARIABLES - START

	pIdProduct=request("idProduct")
	pSku=request("sku")
	pSku=replace(pSku,"'","''")
	pSku=replace(pSku,"""","&quot;")
	origsku=request("origsku")

	'// Determine product type: std, bto, item
	pcv_ProductType=lcase(trim(request("prdType")))

		'// pserviceSpec = 1 ONLY when the product is configurable
		if pcv_ProductType="bto" then
			pserviceSpec="1"
		else
			pserviceSpec="0"
		end if
	
		'// pconfigOnly = 1 ONLY when the product is a configurable-only item
		if pcv_ProductType="item" then
			pconfigOnly="1"
		else
			pconfigOnly="0"
		end if
	
	pDescription=request("description")    
	pDescription=pcf_ReplaceCharacters(pDescription)
    pDescription=replace(pDescription,"""","&quot;")
	pDetails=pcf_ReplaceCharacters(request("details"))
	psDesc=pcf_ReplaceCharacters(request("sDesc"))


	pImageUrl=request("imageUrl")
	pSmallImageUrl=request("smallImageUrl")
	pLargeImageUrl=request("largeImageUrl")
	

	'Downloadable Product Information
	pdownloadable=request("downloadable")
	if trim(pdownloadable)="" then
		pdownloadable="0"
	end if
	purlexpire=request("urlexpire")
	pexpiredays=request("expiredays")
	if not validNum(pexpiredays) then
		pexpiredays="0"
	end if
	plicense=request("license")
	pproducturl=replace(request("producturl"),"'","''")
	plocallg=replace(request("locallg"),"'","''")
	premotelg=replace(request("remotelg"),"'","''")
	if ucase(premotelg)="HTTP://" then
		premotelg=""
	end if
	plicenselabel1=replace(request("licenselabel1"),"'","''")
	plicenselabel2=replace(request("licenselabel2"),"'","''")
	plicenselabel3=replace(request("licenselabel3"),"'","''")
	plicenselabel4=replace(request("licenselabel4"),"'","''")
	plicenselabel5=replace(request("licenselabel5"),"'","''")
	paddtomail=replace(request("addtomail"),"'","''")

	' GGG add-on start
	pGC=request("GC")
	if pGC="" then
		pGC="0"
	end if
	pGCExp=request("GCExp")
	if pGCExp="" then
		pGCExp="0"
	end if
	pGCExpDate=request("GCExpDate")
	if pGCExp<>"1" then
		pGCExpDate="01/01/1900"
	end if
	pGCExpDay=request("GCExpDay")
	if pGCExp<>"2" then
		pGCExpDay="0"
	end if
	pGCEOnly=request("GCEOnly")
	if pGCEOnly="" then
		pGCEOnly="0"
	end if
	pGCGen=request("GCGen")
	if pGCGen="" then
		pGCGen="0"
	end if
	pGCGenFile=request("GCGenFile")
	if pGCGen<>"1" then
		pGCGenFile=""
	end if
	'GGG add-on end
	
	pPrice=replacecomma(request("price"))
	if pPrice="" then
		pPrice="0"
	end if
	pListPrice=replacecomma(request("listPrice"))
	if pListPrice="" then
		pListPrice="0"
	end if
	pBtoBPrice=replacecomma(request("btoBPrice"))
	if pBtoBPrice="" then
		pBtoBPrice="0"
	end If

	'Start SDBA
	pCost=replacecomma(request("cost"))
	if pCost="" then
		pCost="0"
	end If
	
	pcbackorder=replacecomma(request("pcbackorder"))
	if (pcbackorder="") or (not IsNumeric(pcbackorder)) then
		pcbackorder="0"
	end If
	
	pcShipNDays=replacecomma(request("pcShipNDays"))
	if (pcShipNDays="") or (not IsNumeric(pcShipNDays)) then
		pcShipNDays="0"
	end If
	
	pcnotifystock=replacecomma(request("pcnotifystock"))
	if (pcnotifystock="") or (not IsNumeric(pcnotifystock)) then
		pcnotifystock="0"
	end If
	
	pcreorderlevel=replacecomma(request("pcreorderlevel"))
	if (pcreorderlevel="") or (not IsNumeric(pcreorderlevel)) then
		pcreorderlevel="0"
	end If
	
	pcIDSupplier=replacecomma(request("pcIDSupplier"))
	if (pcIDSupplier="") or (not IsNumeric(pcIDSupplier)) then
		pcIDSupplier="0"
	end If
	pIdSupplier=request("idSupplier")
	
	
	pcIsdropshipped=replacecomma(request("pcIsdropshipped"))
	if (pcIsdropshipped="") or (not IsNumeric(pcIsdropshipped)) then
		pcIsdropshipped="0"
	end If
	
	pcIDDropShipper=request("pcIDDropShipper")
	if (pcIDDropShipper="") then
		pcIDDropShipper="0"
	end If
	
	pcDropShipperSupplier=0
	
	if pcIDDropShipper<>"0" then
		pcArr=split(pcIDDropShipper,"_")
		pcIDDropShipper=pcArr(0)
		pcDropShipperSupplier=pcArr(1)
	end if
	
	'End SDBA
	
	' Hide prices
	pnoprices=request("noprices")
	if pnoprices="" then
		pnoprices="0"
	end if	

	pListHidden=request("listHidden")
	if pListHidden="" then
		plistHidden="0"
	end if
	
	pActive=request("active")
	if pActive="" then
		pactive="0"
	end if
	
	pHotDeal=request("hotDeal")
	if pHotDeal="" then
		photDeal="0"
	end if
	
	pShowInHome=request("showInHome")
	if pShowInHome="" then
		pShowInHome="0"
	end if
	
	pIDBrand=request("IDBrand")
	if not validNum(pIDBrand) then
		pIDBrand="0"
	end if

	pDisplayLayout=lcase(request("displayLayout"))
	if pDisplayLayout<>"custom" and pDisplayLayout<>"stand" and pDisplayLayout<>"computer" and pDisplayLayout<>"monitor" and pDisplayLayout<>"traderpc" and pDisplayLayout<>"traderpropc" and pDisplayLayout<>"charterpc" then
		pDisplayLayout="custom"
	end if
		
	pWeight=request("weight")
	if pWeight="" then
		pWeight="0"
	End If
	
	pWeight_oz=request("weight_oz")
	If pWeight_oz="" then
		pWeight_oz="0"
	End if
	
	pcv_QtyToPound=request("QtyToPound")
	if NOT isNumeric(pcv_QtyToPound) or pcv_QtyToPound="" then
		pcv_QtyToPound=0
	end if
	
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

	pStock=request("stock")
	if pStock="" then
		pStock="0"
	end if
	
	pNoStock=request("noStock")
	if pNoStock="" then
		pNoStock="0"
	end if 
	
	pDeliveringTime=request("deliveringTime")
	if pDeliveringTime="" then
		pDeliveringtime="0"
	end If

	pnotax=request("notax")
	if pnotax<>"-1" then
	  pnotax="0"
	end if
	
	pnoshipping=request("noshipping")
	if pnoshipping="" then
		pnoshipping="0"
	end if
	
	pnoshippingtext=request("noshippingtext")
	if pnoshippingtext="" then
		pnoshippingtext="0"
	end if

	'GGG Add-on start
	'Electronic gift certificates are non-taxable and are not shipped
	if (pGCEOnly="1") and (pGC="1") then
		pnotax="-1"
		pnoshipping="-1"
		pNoStock="-1"
	end if
	'GGG Add-on end

	'Not for sale
	pFormQuantity=request("formQuantity")
	If pFormQuantity="" then
		pFormQuantity="0"
	End If
	
	'Not for sale message
	pEmailText=replace(request("emailText"),"""","&quot;")
	pEmailText=replace(pEmailText,"'","''")
	
	'Reward Points
	iRewardPoints = Request("iRewardPoints")
	if scDecSign="," then
		iRewardPoints=replace(iRewardPoints,".","")
	else
		iRewardPoints=replace(iRewardPoints,",","")
	end if
	if iRewardPoints="" then
		iRewardPoints=0
	end if
	iRewardPoints=round(iRewardPoints)

	pOverSizeSpec=request("OverSizeSpec")
	if pOverSizeSpec="YES" then
		pOS_height=request("os_height")
		pOS_width=request("os_width")
		pOS_length=request("os_length")
		if pOS_height="" OR pOS_width="" OR pOS_length="" then
			pOverSizeSpec="NO"
		else
			'find girth (width*2)+(height*2)+length
			pOS_girth=((pOS_width*2)+(pOS_height*2)+pOS_length)
			'response.write pOS_girth&"<BR>"
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
	end if
	
	'Surcharge
	pcv_Surcharge1=replacecomma(request("surcharge1"))
	if pcv_Surcharge1="" then
		pcv_Surcharge1="0"
	end if

	pcv_Surcharge2=replacecomma(request("surcharge2"))
	if pcv_Surcharge2="" then
		pcv_Surcharge2="0"
	end if

	'get Hide Configurator Price option
	pcv_intHideBTOPrice=request("hidebtoprice")
		if pcv_intHideBTOPrice="" then
			pcv_intHideBTOPrice="0"
		end if
		
	'get Hide default Configuration on the product details page
	intHideDefConfig=request("hideDefConfig")
		if intHideDefConfig="" then
			intHideDefConfig="0"
		end if
		
	'get Skip product details page
	pcv_intSkipDetailsPage=request("pcv_intSkipDetailsPage")
	if pcv_intSkipDetailsPage="" then
		pcv_intSkipDetailsPage="0"
	end if
	
	pcv_MaxSelect=request("maxselect")
	if pcv_MaxSelect="" then
		pcv_MaxSelect=0
	end if
	if Not IsNumeric(pcv_MaxSelect) then
		pcv_MaxSelect=0
	end if
	
	'MojoZoom image magnifier
	pcv_IntMojoZoom=request("MojoZoom")
	if pcv_IntMojoZoom="" then
		pcv_IntMojoZoom=0
	end if
	if Not validNum(pcv_IntMojoZoom) then
		pcv_IntMojoZoom=0
	end if
	pAltTagText=pcf_ReplaceCharacters(request("AltTagText"))
	
	pcv_IntAdditionalImages=request("AdditionalImages")
	if pcv_IntAdditionalImages="" then
		pcv_IntAdditionalImages=0
	end if
	if Not validNum(pcv_IntAdditionalImages) then
		pcv_IntAdditionalImages=0
	end if
		
	'get Validate Quantity option
	pcv_intQtyValidate=request("QtyValidate")
		if pcv_intQtyValidate="" then
			pcv_intQtyValidate="0"
		end if
	pcv_lngMinimumQty=request("MinimumQty")
		if pcv_lngMinimumQty="" then
			pcv_lngMinimumQty="0"
		end if
	pcv_lngMultiQty=request("multiQty")
		if pcv_lngMultiQty="" then
			pcv_lngMultiQty="0"
		end if
		if pcv_lngMultiQty=0 then
			pcv_intQtyValidate=0
		end if
	pcv_intHideSKU=request("hideSKU")
		if pcv_intHideSKU="" then
			pcv_intHideSKU=0
		end if

	pcv_showBtoCmMsg=request("showBtoCmMsg")
	if pcv_showBtoCmMsg="" then
		pcv_showBtoCmMsg="0"
	end if
	
	if not IsNumeric(pcv_showBtoCmMsg) then
		pcv_showBtoCmMsg=0
	end if

	'//Retrieve Product Meta Tag related fields
	pcv_StrPrdMetaTitle=getUserInput(request.Form("PrdMetaTitle"), 0)
	pcv_StrPrdMetaDesc=getUserInput(request.Form("PrdMetaDesc"), 0)
	pcv_StrPrdMetaKeywords=getUserInput(request.Form("PrdMetaKeywords"), 0)

	'Get Apparel Product Settings

	pcv_Apparel=request("apparel")
	if pcv_Apparel="" then
		pcv_Apparel="0"
	end if

	pcv_ApparelRadio=request("pcv_ApparelRadio")
	if pcv_ApparelRadio="" then
		pcv_ApparelRadio="0"
	end if

	pcv_ShowStockMsg=request("showstockmsg")
	if pcv_ShowStockMsg="" then
		pcv_ShowStockMsg="0"
	end if
	pcv_StockMsg=request("stockmsg")
	if pcv_StockMsg<>"" then
		pcv_StockMsg=replace(pcv_StockMsg,"'","''")
	end if

	pcv_SizeLink=request("sizelink")
	if pcv_SizeLink<>"" then
		pcv_SizeLink=replace(pcv_SizeLink,"'","''")
	end if

	pcv_SizeInfo=request("sizeinfo")
	if pcv_SizeInfo<>"" then
		pcv_SizeInfo=replace(pcv_SizeInfo,"'","''")
	end if

	pcv_SizeImg=request("sizeImg")
	pcv_SizeURL=request("sizeurl")
	if ucase(pcv_SizeURL)="HTTP://" then
		pcv_SizeURL=""
	end if	

	'//Get Google Shopping Settings
	pcv_GPC=request("pcv_GPC")
	if pcv_GPC="" then
		pcv_GPC="0"
	end if
	if pcv_GPC<>"0" then
		pcv_GCat=request("pcv_GCat")
		if pcv_GCat="" then
			pcv_GCat=request("pcv_GCatO")
		end if
	else
		pcv_GCat=""
	end if
	if pcv_GCat<>"" then
		pcv_GCat=replace(pcv_GCat,"'","''")
	end if
	pcv_GGen=request("pcv_GGen")
	if pcv_GGen<>"" then
		pcv_GGen=replace(pcv_GGen,"'","''")
	end if
	pcv_GAge=request("pcv_GAge")
	if pcv_GAge<>"" then
		pcv_GAge=replace(pcv_GAge,"'","''")
	end if
	pcv_GSize=request("pcv_GSize")
	if pcv_GSize<>"" then
		pcv_GSize=replace(pcv_GSize,"'","''")
	end if
	pcv_GColor=request("pcv_GColor")
	if pcv_GColor<>"" then
		pcv_GColor=replace(pcv_GColor,"'","''")
	end if
	pcv_GPat=request("pcv_GPat")
	if pcv_GPat<>"" then
		pcv_GPat=replace(pcv_GPat,"'","''")
	end if
	pcv_GMat=request("pcv_GMat")
	if pcv_GMat<>"" then
		pcv_GMat=replace(pcv_GMat,"'","''")
	end if
	pExemptDisc=request("exemptDisc")
	if pExemptDisc="" then
		pExemptDisc="0"
	end if
	
	'// Validate - Start

	'required fields
	if pDescription="" or pDetails="" then
		call closeDb()
		response.redirect "msg.asp?message=12"
	end if

	' numbers
	if isNumeric(pPrice)=false or isNumeric(pListPrice)=false or isNumeric(pbtobPrice)=false or isNumeric(pStock)=false or isNumeric(pWeight)=false then
		call closeDb()
		response.redirect "msg.asp?message=13"
	end if
	
	' reward points
	if not isNumeric(iRewardPoints) then
		call closeDb()
		response.redirect "msg.asp?message=14"
	end if
	
	ppTop=replace(request("ppTop"),"'","''")
	ppTopLeft=replace(request("ppTopLeft"), "'", "''")
	ppTopRight=replace(request("ppTopRight"),"'","''")
	ppMiddle=replace(request("ppMiddle"),"'","''")
	ppTabs=replace(request("ppTabs"),"'","''")
	if ppTabs<>"" then
		ppTabs=replace(ppTabs,vbCr,"")
		ppTabs=replace(ppTabs,vbLf,"")
	end if
	ppBottom=replace(request("ppBottom"),"'","''")
	
	pcv_AvalaraTaxCode=request("AvalaraTaxCode")
	if pcv_AvalaraTaxCode<>"" then
		pcv_AvalaraTaxCode=replace(pcv_AvalaraTaxCode,"'","''")
	end if
	'// Validate - End
	
'// LOAD VARIABLES - END

'// UPDATE PRODUCT INFORMATION - START


	'check if SKU already exists and flag
	dim DupSKU
	DupSKU=0
	if origsku<>pSku then
		query="SELECT sku FROM products WHERE sku='" &pSku& "';"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if NOT rs.eof then
			DupSKU=1
		end if
		set rs=nothing
	end if

	' Build main query
	query="UPDATE products SET pcProd_Top='" & ppTop & "',pcProd_TopLeft='" & ppTopLeft & "',pcProd_TopRight='" & ppTopRight & "',pcProd_Middle='" & ppMiddle & "',pcProd_Tabs=N'" & ppTabs & "',pcProd_Bottom='" & ppBottom & "',pcProd_GoogleCat='" & pcv_GCat & "',pcProd_GoogleGender='" & pcv_GGen & "',pcProd_GoogleAge='" & pcv_GAge & "',pcProd_GoogleSize='" & pcv_GSize & "',pcProd_GoogleColor='" & pcv_GColor & "',pcProd_GooglePattern='" & pcv_GPat & "',pcProd_GoogleMaterial='" & pcv_GMat & "',pcProd_MaxSelect=" & pcv_MaxSelect & ",pcProd_multiQty=" & pcv_lngMultiQty & ",iRewardPoints=" &iRewardPoints&", IDBrand=" & pIDBrand & ", sku=N'" &pSku& "', description=N'" &pDescription& "', details=N'" &pDetails& "', serviceSpec=" &pserviceSpec& ", configOnly=" &pconfigOnly& ", price=" &pprice& ", listPrice=" &pListPrice& ", cost=" &pCost& ", imageUrl='" &pImageUrl& "', weight=" &pWeight& ", stock=" &pStock& ", listHidden=" &plistHidden& ", hotDeal=" &pHotDeal& ", active=" &pActive& ", showInHome=" &pShowInHome& ", idSupplier= "&pIdSupplier& ", emailText=N'" &pEmailText& "', bToBPrice=" &pBToBPrice& ", formQuantity=" &pFormQuantity&", smallImageUrl='" &pSmallImageUrl& "', largeImageUrl='" &pLargeImageUrl& "', notax=" &pnotax& ",noshipping=" &pnoshipping& ",noprices=" &pnoprices& ",OverSizeSpec='" &pOverSizeSpec& "', sdesc=N'" & psDesc & "',downloadable=" & pdownloadable & ", noStock="&pNoStock&", noshippingtext=" &pnoshippingtext& ", pcprod_HideBTOPrice=" & pcv_intHideBTOPrice & ", pcprod_QtyValidate=" & pcv_intQtyValidate & ", pcprod_MinimumQty=" & pcv_lngMinimumQty & ", pcprod_QtyToPound="&pcv_QtyToPound&", pcprod_HideDefConfig=" & intHideDefConfig & ", pcProd_BackOrder=" & pcbackorder & ", pcProd_ShipNDays=" & pcShipNDays & ",pcProd_NotifyStock=" & pcnotifystock & ",pcProd_ReorderLevel=" & pcreorderlevel & ",pcSupplier_ID=" & pcIDSupplier & ", pcProd_IsDropShipped=" & pcIsdropshipped & ",pcDropShipper_ID=" & pcIDDropShipper & ", pcprod_GC=" & pGC & ", pcProd_SkipDetailsPage=" & pcv_intSkipDetailsPage & ", pcprod_DisplayLayout='" & pDisplayLayout & "', pcprod_MetaTitle=N'" &pcv_StrPrdMetaTitle&"', pcprod_MetaDesc=N'" &pcv_StrPrdMetaDesc&"', pcprod_MetaKeywords=N'" &pcv_StrPrdMetaKeywords&"', pcProd_HideSKU=" & pcv_intHideSKU & ",pcProd_showBtoCmMsg=" & pcv_showBtoCmMsg & ", pcprod_Apparel=" & pcv_Apparel & ",pcprod_ShowStockMsg=" & pcv_ShowStockMsg & ",pcprod_StockMsg=N'" & pcv_StockMsg & "',pcprod_SizeLink='" & pcv_SizeLink & "',pcprod_SizeInfo=N'" & pcv_SizeInfo & "',pcprod_SizeImg='" & pcv_SizeImg & "',pcprod_SizeURL='" & pcv_SizeURL & "',pcProd_ApparelRadio=" & pcv_ApparelRadio & ", pcProd_Surcharge1=" & pcv_Surcharge1 & ", pcProd_Surcharge2=" & pcv_Surcharge2 & ", pcPrd_MojoZoom=" & pcv_IntMojoZoom & ", pcProd_AvalaraTaxCode='" & pcv_AvalaraTaxCode & "', pcProd_AdditionalImages=" & pcv_IntAdditionalImages & ", pcProd_AltTagText='" & pAltTagText & "' WHERE idProduct=" &pIdProduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
	
	call pcs_hookProductModified(pIdProduct, "")
	
	
	if request("saveDefault")="1" then
		query="SELECT * FROM pcDefaultPrdLayout;"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			query="UPDATE pcDefaultPrdLayout SET pcDPL_Top='" & ppTop & "',pcDPL_TopLeft='" & ppTopLeft & "',pcDPL_TopRight='" & ppTopRight & "',pcDPL_Middle='" & ppMiddle & "',pcDPL_Tabs=N'" & ppTabs & "',pcDPL_Bottom='" & ppBottom & "';"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		else
			query="INSERT INTO pcDefaultPrdLayout (pcDPL_Top,pcDPL_TopLeft,pcDPL_TopRight,pcDPL_Middle,pcDPL_Tabs,pcDPL_Bottom) VALUES ('" & ppTop & "','" & ppTopLeft & "','" & ppTopRight & "','" & ppMiddle & "',N'" & ppTabs & "','" & ppBottom & "');"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		end if
	end if
	
	'query="SELECT pcFPE_IdProduct FROM pcDFProdsExempt WHERE pcFPE_IdProduct="&pIdProduct
	'set rs=conntemp.execute(query)
	'if pExemptDisc="1" then
	'	if rs.eof then
	'		query="INSERT INTO pcDFProdsExempt(pcFPE_IdProduct) VALUES ("&pIdProduct&")"
	'		set rs=conntemp.execute(query)
	'	end if
	'else
	'	if not rs.eof then
	'		query="DELETE FROM pcDFProdsExempt WHERE pcFPE_IdProduct="&pIdProduct
	'		set rs=conntemp.execute(query)
	'	end if
	'end if
	'set rs=nothing

    '// Did Stock Change?
	nmstock=request("nmstock")
	if nmstock="" then
		nmstock="0"
	end if	
	if Clng(pstock)>Clng(nmstock) then
		call pcs_hookStockChanged(pIdProduct, "")
	end if

	call updPrdEditedDate(pIdProduct)

	'Start SDBA
	'Delete record if it is existing
	query="DELETE FROM pcDropShippersSuppliers WHERE idProduct=" & pIdProduct
	set rstemp=connTemp.execute(query)
	set rstemp=nothing

	'Insert a new record to know the Supplier is also a Drop-shipper or not
	query="INSERT INTO pcDropShippersSuppliers (idProduct,pcDS_IsDropShipper) VALUES (" & pIdProduct & "," & pcDropShipperSupplier & ")"
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)
	set rstemp=nothing
	'End SDBA
	
	pcv_prdnotes=request("prdnotes")
	if pcv_prdnotes<>"" then
		pcv_prdnotes=replace(pcv_prdnotes,"'","''")
		pcv_prdnotes=replace(pcv_prdnotes,"""","&quot;")
		pcv_prdnotes=replace(pcv_prdnotes,"<","&lt;")
		pcv_prdnotes=replace(pcv_prdnotes,">","&gt;")
	end if
	query="UPDATE Products SET pcProd_PrdNotes=N'" & pcv_prdnotes & "' WHERE idproduct=" & pIdProduct & ";"
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
	
	'Update Product Search Fields
	SFData=request("SFData")
	query="DELETE FROM pcSearchFields_Products WHERE idproduct=" & pIdProduct & ";"
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
	if SFData<>"" then
		tmp1=split(SFData,"||")
		For i=0 to ubound(tmp1)
			if tmp1(i)<>"" then
				tmp2=split(tmp1(i),"^^^")
				if tmp2(2)="-1" then
					query="SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & tmp2(1) & " AND pcSearchDataName like '" & tmp2(3) & "';"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						query="UPDATE pcSearchData SET idSearchField=" & tmp2(1) & ",pcSearchDataName=N'" & tmp2(3) & "',pcSearchDataOrder=" & tmp2(4) & " WHERE idSearchField=" & tmp2(1) & " AND pcSearchDataName like '" & tmp2(3) & "';"
						set rsQ=connTemp.execute(query)
					else
						query="INSERT INTO pcSearchData (idSearchField,pcSearchDataName,pcSearchDataOrder) VALUES (" & tmp2(1) & ",N'" & tmp2(3) & "'," & tmp2(4) & ");"
						set rsQ=connTemp.execute(query)
					end if
					set rsQ=nothing

					query="SELECT idSearchData FROM pcSearchData WHERE pcSearchDataName like '" & tmp2(3) & "';"
					set rsQ=connTemp.execute(query)
					idSearchData=rsQ("idSearchData")
					set rsQ=nothing
				else
					idSearchData=tmp2(2)
				end if
				query="INSERT INTO pcSearchFields_Products (idproduct,idSearchData) VALUES (" & pIdProduct & "," & idSearchData & ");"
				set rsQ=connTemp.execute(query)
				set rsQ=nothing
			end if
		Next
	end if

'check if there are customer categories
query="SELECT idcustomerCategory, pcCC_CategoryType FROM pcCustomerCategories;"
SET rs=Server.CreateObject("ADODB.RecordSet")
SET rs=conntemp.execute(query)
if NOT rs.eof then 
	do until rs.eof 
		intIdcustomerCategory=rs("idcustomerCategory")
		intpcCC_Price=request("pcCC_"&intIdcustomerCategory)
		intpcCC_Price=replacecomma(intpcCC_Price)
	
		if isNumeric(intpcCC_Price) then
			if intpcCC_Price>0 then
				'// INSERT or UPDATE the Customer Pricing Category Price
				query="SELECT idCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&intIdcustomerCategory&" AND idProduct="&pIdProduct&";"
				SET rsPBPObj=Server.CreateObject("ADODB.RecordSet")
				SET rsPBPObj=conntemp.execute(query)
				if rsPBPObj.eof then
					query="INSERT INTO pcCC_Pricing (idcustomerCategory, idProduct, pcCC_Price) VALUES ("&intIdcustomerCategory&","&pIdProduct&","&intpcCC_Price&");"
				else
					intIdCC_Price=rsPBPObj("idCC_Price")
					query="UPDATE pcCC_Pricing SET pcCC_Price="&intpcCC_Price&" WHERE idCC_Price="&intIdCC_Price&";"
				end if
				SET rsIObj=Server.CreateObject("ADODB.RecordSet")
				SET rsIObj=conntemp.execute(query)

				SET rsIObj=nothing
				SET rsPBPObj=nothing
			else
				'// Price = 0 -> REMOVE Customer Pricing Category Price = Default price for that pricing category will be used
				query="DELETE FROM pcCC_Pricing WHERE idcustomerCategory="&intIdcustomerCategory&" AND idProduct="&pIdProduct&";"
				SET rsPBPObj=Server.CreateObject("ADODB.RecordSet")
				SET rsPBPObj=conntemp.execute(query)
				SET rsPBPObj=nothing
			end if
		end if
		rs.moveNext
	loop
end if
SET rs=nothing

'// Check if there are customer categories
If statusAPP="1" OR scAPP=1 Then

	query="SELECT idcustomerCategory, pcCC_CategoryType, pcCC_ATB_Off, pcCC_ATB_Percentage FROM pcCustomerCategories;"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if NOT rs.eof then 
		do until rs.eof 
			intIdcustomerCategory=rs("idcustomerCategory")
			pcv_strCategoryType=rs("pcCC_CategoryType")
			pcv_strCCATBOff=rs("pcCC_ATB_Off")
			pcv_strCCATBPercentage=rs("pcCC_ATB_Percentage")
			
			intpcCC_Price=request("pcCC_"&intIdcustomerCategory)
			intpcCCOrig_Price=request("pcCCOrig_"&intIdcustomerCategory)
			if NOT isNumeric(intpcCCOrig_Price) then
				intpcCCOrig_Price=0
			end if
			intpcCC_Price=replacecomma(intpcCC_Price)
		
			'if validNum(intpcCC_Price) then
			if isNumeric(intpcCC_Price) AND (intpcCCOrig_Price<>intpcCC_Price) then
				if Cdbl(intpcCC_Price)<>0 then
					query="DELETE FROM pcCC_Pricing WHERE idcustomerCategory="&intIdcustomerCategory&" AND idproduct IN (SELECT idproduct FROM Products WHERE pcprod_ParentPrd=" & pIdProduct & ");"
					SET rsIObj=Server.CreateObject("ADODB.RecordSet")
					SET rsIObj=conntemp.execute(query)
					SET rsIObj=nothing
					
					query="SELECT idproduct,pcprod_addprice,pcprod_addWprice,pcprod_Relationship FROM Products WHERE pcprod_ParentPrd=" & pIdProduct & " AND removed=0;"
					SET rsIObj=Server.CreateObject("ADODB.RecordSet")
					SET rsIObj=conntemp.execute(query)
					if not rsIObj.eof then
						pcArr=rsIObj.getRows()
						SET rsIObj=nothing
						intCount=ubound(pcArr,2)
						for i=0 to intCount	
						
							'// Calculate the differential
							pcv_SIdproduct=pcArr(0,i)
							If pcv_strCCATBOff="Retail" then
								pcv_Addprice=pcArr(1,i)
							Else
								pcv_Addprice=pcArr(2,i)
							End If										
							if pcv_Addprice<>"" then
							else
							pcv_Addprice=0
							end if								
							
							'// ATB: Make Adjustments
							if pcv_strCategoryType="ATB" then
								'// Apply the ATB to the differential	
								pcv_Addprice=pcv_Addprice-(pcf_Round(pcv_Addprice*(cdbl(pcv_strCCATBPercentage)/100),2))
							end if
							
							'// Calculate the Final Price				
							pcv_PPrice=cdbl(intpcCC_Price)+cdbl(pcv_Addprice)												
	
							query="INSERT INTO pcCC_Pricing (idcustomerCategory, idProduct, pcCC_Price) VALUES ("&intIdcustomerCategory&"," & pcv_SIdproduct & ","&pcv_PPrice&");"
							SET rsIObj=Server.CreateObject("ADODB.RecordSet")
							SET rsIObj=conntemp.execute(query)
							SET rsIObj=nothing
						next
					end if
					SET rsIObj=nothing
					
				else
					query="DELETE FROM pcCC_Pricing WHERE idcustomerCategory=" & intIdcustomerCategory & " AND idproduct IN (SELECT idproduct FROM Products WHERE pcprod_ParentPrd=" & pIdProduct & ");"
					set rsIObj=connTemp.execute(query)
					SET rsIObj=nothing
				end if
			end if
			rs.moveNext
		loop
	end if
	SET rs=nothing

End If

'update downloadable product
if (pdownloadable<>"") and (pdownloadable="1") then
	query="Select * from DProducts where idproduct=" & pIdproduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	if not rs.eof then
		query="Update DProducts set ProductURL='" & pProductURL & "',URLExpire=" & pURLExpire & ",ExpireDays=" & pExpireDays & ",License=" & pLicense & ",LocalLG='" & pLocalLG & "',RemoteLG='" & pRemoteLG & "',LicenseLabel1='" & pLicenseLabel1 & "',LicenseLabel2='" & pLicenseLabel2 & "',LicenseLabel3='" & pLicenseLabel3 & "',LicenseLabel4='" & pLicenseLabel4 & "',LicenseLabel5='" & pLicenseLabel5 & "',AddToMail='" & pAddtoMail & "' where idproduct=" & pIdproduct
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
	else
		query="Insert into DProducts (IdProduct,ProductURL,URLExpire,ExpireDays,License,LocalLG,RemoteLG,LicenseLabel1,LicenseLabel2,LicenseLabel3,LicenseLabel4,LicenseLabel5,AddToMail) values (" & pIdProduct & ",'" & pproducturl & "'," & pURLExpire & "," & pExpireDays & "," & pLicense & ",'" & pLocalLG & "','" & pRemoteLG & "',N'" & plicenselabel1 & "',N'" & plicenselabel2 & "',N'" & plicenselabel3 & "',N'" & plicenselabel4 & "',N'" & plicenselabel5 & "','" & pAddtoMail & "')"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
	end if
else
	query="delete from DProducts where idproduct=" & pIdproduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
end if

'GGG Add-on start
if (pGC<>"") and (pGC="1") then

	if SQL_Format="1" then
		pGCExpDate=(day(pGCExpDate)&"/"&month(pGCExpDate)&"/"&year(pGCExpDate))
	else
		pGCExpDate=(month(pGCExpDate)&"/"&day(pGCExpDate)&"/"&year(pGCExpDate))
	end if

	query="Select pcGC_IDProduct from pcGC where pcGC_idproduct=" & pIdproduct
	set rstemp=conntemp.execute(query)

	IF not rstemp.eof then
		query="Update pcGC set pcGC_Exp=" & pGCExp & ",pcGC_ExpDate='" & pGCExpDate & "',pcGC_ExpDays=" & pGCExpDay & ",pcGC_EOnly=" & pGCEOnly & ",pcGC_CodeGen=" & pGCGen & ",pcGC_GenFile='" & pGCGenFile & "' where pcGC_idproduct=" & pIdproduct
		set rstemp=conntemp.execute(query)
	ELSE
		query="Insert into pcGC (pcGC_IdProduct,pcGC_Exp,pcGC_ExpDate,pcGC_ExpDays,pcGC_EOnly,pcGC_CodeGen,pcGC_GenFile) values (" & pIdProduct & "," & pGCExp & ",'" & pGCExpDate & "'," & pGCExpDay & "," & pGCEOnly & "," & pGCGen & ",'" & pGCGenFile & "')"
		set rstemp=conntemp.execute(query)
	END IF
else
	query="delete from pcGC where pcGC_idproduct=" & pIdproduct
	set rstemp=conntemp.execute(query)
end if
'GGG Add-on end

'//Update sub-products
If statusAPP="1" OR scAPP=1 Then

	if pcv_Apparel="1" then

		query="Select idproduct,pcprod_addprice,pcprod_addWprice,pcprod_Relationship from Products where pcprod_ParentPrd=" & pidproduct & " AND removed=0 AND active=0"
		set rstemp=connTemp.execute(query)
		
		pPrice1=pPrice
		pBtoBPrice1=pBtoBPrice
		if (pBtoBPrice1="") or (pBtoBPrice1="0") then
			pBtoBPrice1=pPrice1
		end if
		
		do while not rstemp.eof
			pcv_SIdproduct=rstemp("idproduct")
			
			pcv_Addprice=rstemp("pcprod_Addprice")	
			if NOT (pcv_Addprice<>"") then
				pcv_Addprice=0
			end if
			
			pcv_AddWprice=rstemp("pcprod_AddWprice")
			if NOT ((pcv_AddWprice<>"") and (pcv_AddWprice<>"0")) then
				pcv_AddWprice="0"
			end if
			
			pcv_PPrice=cdbl(pPrice1)+cdbl(pcv_Addprice)	
			pcv_PWPrice=cdbl(pBToBPrice1)+cdbl(pcv_AddWprice)
			pcv_TempArr=split(rstemp("pcprod_Relationship"),"_")	
			
			pcv_newName="("	
			For i=1 to ubound(pcv_TempArr)
				pcv_Opt1=pcv_TempArr(i)
		
				query="select idOption from options_optionsGroups where idoptoptgrp=" & pcv_Opt1
				set rs=connTemp.execute(query)
		
				if not rs.eof then
					pcv_ROpt1=rs("idOption")
				end if
		
				query="select optionDescrip from Options where idOption=" & pcv_ROpt1
				set rs=connTemp.execute(query)
		
				if not rs.eof then
					pcv_Code1=rs("optionDescrip")
				end if
		
				pcv_newName=pcv_newName & pcv_Code1 
				if (i<ubound(pcv_TempArr)) then
					pcv_newName=pcv_newName & " - "
				end if
			Next
			pcv_newName=pcv_newName & ")"
			pcv_newName=replace(pcv_newName,"'","''")
			pcv_newName=pDescription & " " & pcv_newName
			
			query="UPDATE products SET description=N'" & pcv_newName & "',price=" & pcv_PPrice & ",btoBPrice=" & pcv_PWPrice & ",cost=" &pCost& ",notax=" &pnotax& ",noshipping=" &pnoshipping& ",noshippingtext=" & pnoshippingtext & ",OverSizeSpec='" &pOverSizeSpec& "',pcProd_NotifyStock=" & pcnotifystock & ",pcProd_ReorderLevel=" & pcreorderlevel & ",pcSupplier_ID=" & pcIDSupplier & ",pcProd_IsDropShipped=" & pcIsdropshipped & ",pcDropShipper_ID=" & pcIDDropShipper & ",pcprod_qtyvalidate=" & pcv_intQtyValidate & ",pcprod_minimumqty=" & pcv_lngMinimumQty & " WHERE idproduct=" & pcv_SIdproduct
			set rs=connTemp.execute(query)
			call pcs_hookProductModified(pcv_SIdproduct, "")
			
			'Start SDBA
			'Delete record if it is existing
			query="DELETE FROM pcDropShippersSuppliers WHERE idProduct=" & pcv_SIdproduct
			set rs=connTemp.execute(query)
			set rs=nothing
		
			'Insert a new record to know the Supplier is also a Drop-shipper or not
			query="INSERT INTO pcDropShippersSuppliers (idProduct,pcDS_IsDropShipper) VALUES (" & pcv_SIdproduct & "," & pcDropShipperSupplier & ")"
			set rs=connTemp.execute(query)
			set rs=nothing
			'End SDBA
			
			rstemp.MoveNext
		
		loop
	
	end if

End IF

Call RunCalBDPC()

If statusAPP="1" OR scAPP=1 Then
	'// Update parent products inventory levels if necessary 
	%>
	<!--#include file="../pc/app-updstock.asp"-->
	<%
End IF

'// Get "tab" querystring, if it exists:
tab = request("tab")
if len(tab)>0 then
	tabQS = "&tab=" & tab & "#TabbedPanels1"
else
	tabQS = ""
end if

if request("re1")="0" then
	call closeDb()
	response.redirect "FindProductType.asp?id=" & pIdproduct & tabQS
end if 

if err.number <> 0 then
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&Err.Description) 
end If
%>

<!--#include file="AdminHeader.asp"-->

<table class="pcCPcontent">
	<tr>
		<td colspan="2">
			<div class="pcCPmessageSuccess">&quot;<%=removeSQ(pDescription)%>&quot; was successfully modified. <a href="FindProductType.asp?id=<%=pIdProduct%>"><strong>Edit it again</strong></a>.</div>
		</td>
	</tr>
	<tr>
		<td width="50%" valign="top">
			<% if DupSKU=1 then %>
				<p><span class="pcCPnotes">Warning: DUPLICATE SKU - The part number (SKU) that you have entered already exists in the database. It is up to you to decide whether to keep it 'as is' or to change it to a unique one. To edit it, select 'Modify this product again' from the menu below.</span></p>
			<% end if %>
			<ul class="pcListIcon">
				<% if pcv_ProductType<>"item" then %>
				<li><a href="../pc/viewPrd.asp?idproduct=<%=pIdProduct%>&adminPreview=1" target="_blank">Preview in your storefront &gt;&gt;</a></li>
				<li style="padding-bottom: 10px;"><a href="FindProductType.asp?id=<%=pIdProduct%>">Modify this product again</a></li>
				<% end if %>
				
				<% if pcv_ProductType="std" then %>
					<li><a href="modPrdOpta.asp?idproduct=<%=pIdProduct%>">Add/Modify product options</a> (e.g. sizes, colors, etc.)</li>
					<% if pcv_Apparel="1" then %>
						<li><a href="app-subPrdsMngAll.asp?idproduct=<%=pIdProduct%>">Manage sub-products</a></li>
					<% end if %>
				<% elseif pcv_ProductType="bto" then %>
					<li style="padding-top: 10px"><a href="modBTOconfiga.asp?idProduct=<%=pIdProduct%>"><strong>Setup</strong> this configurable product or service</a></li>
				<% end if %>
				
				<% if pcv_ProductType<>"item" then %>
				<li><a href="AdminCustom.asp?idproduct=<%=pIdProduct%>">Add/Modify custom search or input fields</a></li>
				<li><a href="crossSellAddb.asp?action=source&prdlist=<%=pIdProduct%>">Add a cross-selling relationship</a></li>
				<% end if %>

				<% if pcv_Apparel="1" then %>
					<li><a href="viewDisca.asp">View/Add quantity discounts</a></li>
				<% else %>	
					<li><a href="FindProductQtyDisc.asp?idproduct=<%=pIdProduct%>">View/Add quantity discounts</a></li>
				<% end if %>
                
                <% if pcv_ProductType<>"item" then %>
                <li style="padding-top: 10px;"><a href="prv_ManageReviews.asp?IDProduct=<%=pIdProduct%>&nav=2">View/Manage reviews for this product</a></li>
                <% end if %>
             </ul>
			<% if pcv_ProductType="item" then %>
            <div style="margin: 20px; padding: 10px; border: 1px solid #CCC; color:#999;">Some of the links that are available when you add or edit a Standard or Configurable product are not available when the product is a Configurable-Only Item.</div>
            <% end if %>
          </td>
          <td width="50%" valign="top">
			<ul class="pcListIcon">
				<%
				query="SELECT idcategory FROM categories_products WHERE idproduct=" & pIdProduct & ";"
				set rs=connTemp.execute(query)
				if not rs.eof then
					pcv_ParentCatID=rs("idcategory")
					query="SELECT products.idproduct,products.serviceSpec,products.configOnly FROM products INNER JOIN categories_products ON products.idproduct=categories_products.idproduct WHERE idcategory=" & pcv_ParentCatID & " ORDER BY categories_products.POrder ASC,products.SKU ASC,products.description ASC;"
					set rs=connTemp.execute(query)
					if not rs.eof then
						pcArr=rs.getRows()
						intCount=ubound(pcArr,2)
						pcv_NextPrdID=0
						For i=0 to intCount
							if clng(pcArr(0,i))=clng(pIdProduct) then
								if i<intCount then
									pcv_NextPrdID=pcArr(0,i+1)
								else
									pcv_NextPrdID=pcArr(0,0)
								end if
								exit for
							end if
						Next
						if pcv_NextPrdID>0 then%>
								<li><a href="FindProductType.asp?id=<%=pcv_NextPrdID%>">Modify the next product in this category</a></li>
						<%
						end if
					end if
					set rs=nothing
				end if
				set rs=nothing
				%>
                <li><a href="editCategories.asp?nav=&lid=<%=pcv_ParentCatID%>">View other products assigned to this category</a></li>
                <li><a href="../pc/viewcategories.asp?idcategory=<%=pcv_ParentCatID%>" target="_blank">View the category in the storefront</a></li>
				<li><a href="manageCategories.asp">Manage categories</a></li>
				<% if pcv_ProductType="item" then %>
				<li><a href="AddRmvBTOItemsMulti1.asp">Assign to one or more configurable products</a></li>
				<% end if %>
				
				<li style="padding-top: 10px;"><a href="FindDupProductType.asp?idproduct=<%=pIdProduct%>">Clone this product</a></li>
				<li><a href="addProduct.asp?prdType=std">Add another product</a></li>
				<li><a href="LocateProducts.asp">Locate another product</a></li>
			</ul>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->