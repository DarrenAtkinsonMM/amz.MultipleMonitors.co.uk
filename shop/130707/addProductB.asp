<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="pcCalculateBTODefaultPrices.asp" -->
<%
dim pageTitle, section, f
pageTitle="Product Added"
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
	
		'// pconfigOnly = 1 ONLY when the product is a CONFIGURABLE ITEM
		if pcv_ProductType="item" then
			pconfigOnly="1"
		else
			pconfigOnly="0"
		end if


	pDescription=replace(request("description"),"""","&quot;")
	pDescription=pcf_ReplaceCharacters(pDescription)
	pDetails=pcf_ReplaceCharacters(request("details"))
	psDesc=pcf_ReplaceCharacters(request("sDesc"))
	pIdcategory1=request("idCategory1")

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
	if pDisplayLayout<>"c" and pDisplayLayout<>"l" and pDisplayLayout<>"o" and pDisplayLayout<>"t" then
		pDisplayLayout=""
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
	If pFormQuantity="" Then
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

	'//Retrieve Product Meta Tag related fields
	pcv_StrPrdMetaTitle=getUserInput(request.Form("PrdMetaTitle"), 0)
	pcv_StrPrdMetaDesc=getUserInput(request.Form("PrdMetaDesc"), 0)
	pcv_StrPrdMetaKeywords=getUserInput(request.Form("PrdMetaKeywords"), 0)
	
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

	pcv_showBtoCmMsg=request("showBtoCmMsg")
	if pcv_showBtoCmMsg="" then
		pcv_showBtoCmMsg="0"
	end if
	
	if not IsNumeric(pcv_showBtoCmMsg) then
		pcv_showBtoCmMsg=0
	end if

	ppTop=replace(request("ppTop"),"'","''")
	ppTopLeft=replace(request("ppTopLeft"),"'","''")
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


'// ADD PRODUCT INFORMATION - START


	'check if SKU already exists and flag
	dim DupSKU
	DupSKU=0
	query="SELECT sku FROM products WHERE sku='" &pSku& "';"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if NOT rs.eof then
		DupSKU=1
	end if
	set rs=nothing

	' Build main query
	dim dtTodaysDate
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
	end if

	query="INSERT INTO products (sku, description, details, configOnly, serviceSpec, price, listPrice, bToBPrice, cost, imageUrl, listhidden, weight, stock, active, idSupplier,hotDeal, emailText, deliveringTime, formQuantity, smallImageUrl,largeImageUrl, notax, noshipping, showInHome,iRewardPoints,IDBrand,OverSizeSpec,sdesc,downloadable,noStock,noshippingtext,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty,pcprod_QtyToPound, pcprod_EnteredOn,pcProd_BackOrder,pcProd_ShipNDays,pcProd_NotifyStock,pcProd_ReorderLevel,pcSupplier_ID,pcProd_IsDropShipped,pcDropShipper_ID,pcprod_GC,pcProd_SkipDetailsPage,pcprod_DisplayLayout, pcprod_MetaTitle, pcprod_MetaDesc, pcprod_MetaKeywords, pcProd_HideSKU,pcProd_showBtoCmMsg, pcProd_Apparel,pcProd_ParentPrd,pcProd_RelationShip,pcProd_ShowStockMsg,pcProd_StockMsg,pcProd_SizeLink,pcProd_SizeInfo,pcProd_SizeImg,pcProd_SizeURL,pcProd_AddPrice,pcProd_ApparelRadio, pcProd_Surcharge1, pcProd_Surcharge2, pcProd_multiQty, pcProd_MaxSelect, pcPrd_MojoZoom,pcProd_GoogleCat,pcProd_GoogleGender,pcProd_GoogleAge,pcProd_GoogleSize,pcProd_GoogleColor,pcProd_GooglePattern,pcProd_GoogleMaterial,pcProd_Top,pcProd_TopLeft,pcProd_TopRight,pcProd_Middle,pcProd_Bottom,pcProd_Tabs,pcProd_AvalaraTaxCode,pcProd_AltTagText) VALUES (N'" &pSku& "',N'" &pDescription& "',N'" & pDetails& "'," &pconfigOnly& "," &pserviceSpec& "," &pPrice& "," &pListPrice& "," &pBtoBPrice& "," &pCost& ",'" &pImageUrl& "'," &pListhidden& "," &pWeight& "," &pStock& "," &pActive& "," &pIdSupplier& "," &pHotDeal& ",N'" &pEmailText& "'," &pDeliveringTime& "," &pFormQuantity& ",'" &pSmallImageUrl& "','"&pLargeImageUrl&"',"&pnotax&","&pnoshipping&","&pshowInHome&","&iRewardPoints&"," & pIDBrand & ",'"&pOverSizeSpec&"',N'" & psdesc & "'," & pdownloadable & ","&pNoStock&","&pnoshippingtext&"," & pcv_intHideBTOPrice & "," & pcv_intQtyValidate & "," & pcv_lngMinimumQty & ","&pcv_QtyToPound&",'"&dtTodaysDate&"'," & pcbackorder & "," & pcShipNDays & "," & pcnotifystock & "," & pcreorderlevel & "," & pcIDSupplier & "," & pcIsdropshipped & "," & pcIDDropShipper & "," & pGC & "," & pcv_intSkipDetailsPage & ",'" & pDisplayLayout & "', N'"&pcv_StrPrdMetaTitle&"', N'"&pcv_StrPrdMetaDesc&"', N'"&pcv_StrPrdMetaKeywords&"'," & pcv_intHideSKU & "," & pcv_showBtoCmMsg & ", " & pcv_Apparel & ",0,''," & pcv_ShowStockMsg & ",N'" & pcv_StockMsg & "','" & pcv_SizeLink & "',N'" & pcv_SizeInfo & "','" & pcv_SizeImg & "','" & pcv_SizeURL & "',0," & pcv_ApparelRadio & ", " & pcv_Surcharge1 & ", " & pcv_Surcharge2 & "," & pcv_lngMultiQty & "," & pcv_MaxSelect & "," & pcv_IntMojoZoom & ",'" & pcv_GCat & "','" & pcv_GGen & "','" & pcv_GAge & "','" & pcv_GSize & "','" & pcv_GColor & "','" & pcv_GPat & "','" & pcv_GMat & "','" & ppTop & "','" & ppTopLeft & "','" & ppTopRight & "', '" & ppMiddle & "','" & ppBottom & "',N'" & ppTabs & "','" & pcv_AvalaraTaxCode & "','" & pAltTagText & "')"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)

	if err.number <> 0 then
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in main database query on addProudctB.asp") 
	end if

	' retrieved information on the product that was just inserted
	query="SELECT idProduct FROM products WHERE sku='" &pSku& "' AND price=" &pPrice& " ORDER by idProduct DESC"
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		set rs=nothing
		
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in query retrieving product from database on addProductB.asp") 
	end if

	if rs.eof then
		set rs=nothing
		
	   call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Can't find inserted product") 
	end if

	pIdProduct=rs("idProduct")
	set rs=nothing
	
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
	
	'Start SDBA
	'Insert a new record to know the Supplier is also a Drop-shipper or not
	query="INSERT INTO pcDropShippersSuppliers (idProduct,pcDS_IsDropShipper) VALUES (" & pIdProduct & "," & pcDropShipperSupplier & ")"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
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
	
	call pcs_hookProductModified(pIdProduct, "")
	
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
					'insert into database
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
				end if
			end if
			rs.moveNext
		loop
	end if
	SET rs=nothing

' Cloned product
IF request("AddType")="19" then 

	'get Custom fields
	query="SELECT IdXField,pcPXF_XReq FROM pcPrdXFields WHERE idProduct=" & request("PreProd") & ";"
	set rs=conntemp.execute(query)
	if not rs.eof then
		pcXArr=rs.getRows()
		set rs=nothing
		intX=ubound(pcXArr,2)
		For k=0 to intX
			strSQL="INSERT INTO pcPrdXFields (IDProduct,IdXfield,pcPXF_XReq) VALUES (" & pIDProduct & "," & pcXArr(0,k) & "," & pcXArr(1,k) & ");"
			set rs=conntemp.execute(strSQL)
			set rs=nothing
		Next
	end if
	set rs=nothing
	
	'//////////////////////////////////////////////////////////////
	'// START: Dup Product Options
	'//////////////////////////////////////////////////////////////	

	iddupProduct=request("PreProd")
	iddupAssignment=""
	if instr(iddupProduct,"|") then
		dupArray=Split(iddupProduct,"|")
		iddupProduct=dupArray(0)
		iddupAssignment=dupArray(1)
	end if
	
	idProduct=pIdProduct
	pricingdup=1 '// dup pricing
	updatedup=1 '// overwrite existing options
		
	strSQL="SELECT DISTINCT options_optionsGroups.idProduct, options_optionsGroups.idOptionGroup, options_optionsGroups.idOption, options_optionsGroups.price, options_optionsGroups.Wprice, options_optionsGroups.sortOrder, options_optionsGroups.InActive, pcProductsOptions.pcProdOpt_Required, pcProductsOptions.pcProdOpt_Order FROM options_optionsGroups INNER JOIN pcProductsOptions ON (options_optionsGroups.idOptionGroup = pcProductsOptions.idOptionGroup) AND (options_optionsGroups.idProduct = pcProductsOptions.idProduct) WHERE ( options_optionsGroups.idProduct="&iddupProduct&" "
	if len(iddupAssignment)>0 then
		strSQL = strSQL & "AND options_optionsGroups.idOptionGroup="&iddupAssignment&" " 
	end if
	strSQL = strSQL & ") ORDER BY options_optionsGroups.idOptionGroup;"

	
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(strSQL)
	
	repeatcnt = 0
	Drepeatcnt = 0
	contgo=0
	xCounter = 0
	cntG=0
	pcv_strMsg = ""
	cnt=0
	
	'//////////////////////////////////////////////////////////////
	'// START: LOOP Through every single Attribute
	'//////////////////////////////////////////////////////////////
	do until rs.eof
	
		'// Get all the Attribute Data
		intidOptionGroup=rs("idOptionGroup")
		intidOption=rs("idOption")
		intprice=rs("price")
		intWprice=rs("Wprice")
		intSortOrder=rs("sortOrder")
		intInActive=rs("InActive")
		intProdOptRequired=rs("pcProdOpt_Required")
		intProdOptOrder=rs("pcProdOpt_Order")
		if isNull(intProdOptRequired) OR intProdOptRequired="" then
			intProdOptRequired=0
		end if
		if isNull(intProdOptOrder) OR intProdOptOrder="" then
			intProdOptOrder=0
		end if
		if xCounter = 0 then
			OintidOptionGroup = intidOptionGroup
		end if
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'  START: Reporting Level Tasks
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		query="SELECT * FROM OptionsGroups WHERE idOptionGroup="&intidOptionGroup
		set rstemp=conntemp.execute(query)
		OptionGroupDesc=rstemp("OptionGroupDesc")
		set rstemp=nothing
		If Session("pcAdmin" & trim(OptionGroupDesc)) = "" AND xCounter>0 Then
			Session("pcAdmin" & trim(OptionGroupDesc)) = trim(OptionGroupDesc)
			if OintidOptionGroup<>intidOptionGroup then
				Drepeatcnt=0
				repeatcnt=0
				cnt=0
				pcv_strCleanUpSessions = pcv_strCleanUpSessions & Session("pcAdmin" & OptionGroupDesc) & ","
				pcv_strMsgMaster = pcv_strMsgMaster & pcv_strMsg
			end if	
		End if	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'  END: Reporting Level Tasks
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'  START: Attribute Level Tasks
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Check if this Attribute ALREADY exists in database before adding				
		strSQL="SELECT * FROM options_optionsGroups WHERE idproduct="& idProduct &" AND idoptionGroup="&intidOptionGroup&" AND idOption="&intidOption&";"
		set rstemp=conntemp.execute(strSQL)
		if rstemp.eof then	
			'// ADD (Replicate)
			if pricingdup="1" then
				strSQL="INSERT INTO options_optionsGroups (idProduct, idOptionGroup, idOption, price, Wprice,sortOrder,InActive) VALUES ("&idProduct&","&intidOptionGroup&","&intidOption&","&intprice&","&intWprice&"," & intSortOrder & "," & intInActive & ");"
			else
				strSQL="INSERT INTO options_optionsGroups (idProduct, idOptionGroup, idOption, price, Wprice,sortOrder,InActive) VALUES ("&idProduct&","&intidOptionGroup&","&intidOption&",0,0"&"," & intSortOrder & "," & intInActive & ");"
			end if	
			set rsSetDup=Server.CreateObject("ADODB.Recordset")
			set rsSetDup=conntemp.execute(strSQL)
			'// Set the update flag
			contgo=1
			cnt=cnt+1		
		else
			'// UPDATE (if over-write was selected)
			if updatedup="1" then
				strSQL="UPDATE options_optionsGroups SET price="&intprice&",Wprice="&intWprice&",sortOrder="&intSortOrder&",InActive="&intInActive&"  WHERE idproduct="& idProduct &" AND idoptionGroup="&intidOptionGroup&" AND idOption="&intidOption&";"
				set rsSetDup=Server.CreateObject("ADODB.Recordset")
				set rsSetDup=conntemp.execute(strSQL)
			end if						
		end if
		set rsSetDup=nothing
		set rstemp=nothing
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'  END: Attribute Level Tasks
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'  START: Product Level Tasks
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// If at least one Attribute was added for this product check that there is a relationship for the Group
		if contgo=1 then		
			'// If this is a new option group, then we need to add the relation
			strSQL="SELECT idOptionGroup, idproduct FROM pcProductsOptions WHERE idproduct="& idProduct &" AND idOptionGroup="& intidOptionGroup &" "
			set rsOptionCheck=conntemp.execute(strSQL)	
			if rsOptionCheck.eof then
				strSQL="INSERT INTO pcProductsOptions (idproduct, idOptionGroup, pcProdOpt_Required, pcProdOpt_Order) VALUES (" & idProduct &", " & intidOptionGroup & ", "&intProdOptRequired&", "&intProdOptOrder&")"
				set rstemp=conntemp.execute(strSQL)
				'// if the option group is new keep count
				cntG=cntG+1
			end if
			set rsOptionCheck = nothing		
		end if
		
		'// If and Attribute was NOT added for this product keep count
		if contgo=0 and updatedup="0" then
			repeatcnt=repeatcnt+1
		end if
		if contgo=0 and updatedup="1" then
			Drepeatcnt=Drepeatcnt+1
		end if
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'  END: Product Level Tasks
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		pcv_strMsg = ""
		
		If cnt>0 then
		pcv_strMsg = cnt &" attributes were copied from the Option Group: <b>"& OptionGroupDesc &"</b>.<br />"
		end if
				
	xCounter = xCounter + 1
	rs.movenext
	loop	
	set rs=nothing
	
	pcv_strMsgMaster = pcv_strMsgMaster & pcv_strMsg
	'response.write "<hr>" & pcv_strMsgMaster & "<hr>" & xCounter
	
	' Clean Up the Sessions
	pcv_strCleanUpSessions = split(pcv_strCleanUpSessions, ",")
	for x = lbound(pcv_strCleanUpSessions) to ubound(pcv_strCleanUpSessions)
		Session("pcAdmin" & pcv_strCleanUpSessions(x)) = ""
	next
	
	If cntG>0 then 
		'// If we have added a new group display that info
		strMsg = cntG &" Option Groups were copied. <br/>" & pcv_strMsgMaster
	else
		'// If we only updated existing groups
		strMsg = pcv_strMsgMaster
	end if	
	'//////////////////////////////////////////////////////////////
	'// END: Dup Product Options
	'//////////////////////////////////////////////////////////////		
	
	'Duplicate Additional Images
	query="SELECT pcProdImage_Url,pcProdImage_LargeUrl,pcProdImage_Order,pcProdImage_AltTagText FROM pcProductsImages WHERE idProduct=" & iddupProduct & ";"
    set rstemp=server.CreateObject("ADODB.RecordSet")
    set rstemp=conntemp.execute(query)
	
	if not rstemp.eof then
		pcArr=rstemp.getRows()
		set rstemp=nothing
		intCount=ubound(pcArr,2)
		For k=0 to intCount
			query="INSERT INTO pcProductsImages (idProduct,pcProdImage_Url,pcProdImage_LargeUrl,pcProdImage_Order,pcProdImage_AltTagText) VALUES("&pIDProduct&",'"&pcArr(0,k)&"','"&pcArr(1,k)&"',"&pcArr(2,k)&",'"&pcf_ReplaceCharacters(pcArr(3,k))&"')"
		    set rstemp=server.CreateObject("ADODB.RecordSet")
		    set rstemp=conntemp.execute(query)
			set rstemp=nothing
		Next
	end if
	
	query="SELECT pcProd_AdditionalImages FROM products WHERE idProduct=" & iddupProduct & ";"
	set rstemp=server.CreateObject("ADODB.RecordSet")
    set rstemp=conntemp.execute(query)
	
	if not rstemp.eof then
		pcv_IntAdditionalImages = rstemp("pcProd_AdditionalImages")
		
		query="UPDATE products SET pcProd_AdditionalImages = " & pcv_IntAdditionalImages & " WHERE idProduct = " & idProduct
		set rstemp2=server.CreateObject("ADODB.RecordSet")
		set rstemp2=conntemp.execute(query)
		set rstemp2=nothing
		call pcs_hookProductModified(idProduct, "")
	end if
	
	set rstemp=nothing

	'get product discounts
	query="select * from discountsPerQuantity where IDProduct=" & request("PreProd")
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=conntemp.execute(query)

	do while not rstemp.eof
		didcategory=rstemp("idcategory")
		ddiscountDesc=rstemp("discountDesc")
		if ddiscountDesc<>"" then
			ddiscountDesc=replace(ddiscountDesc,"'","''")
		end if
		ddiscountDesc= "N'" & ddiscountDesc & "'"
		dquantityFrom=rstemp("quantityFrom")
		dquantityUntil=rstemp("quantityUntil")
		ddiscountPerUnit=rstemp("discountPerUnit")
		dnum=rstemp("num")
		dpercentage=rstemp("percentage")
		ddiscountPerWUnit=rstemp("discountPerWUnit")
		dbaseproductonly=rstemp("baseproductonly")

		strSQL="insert into discountsPerQuantity (idproduct,idcategory,discountDesc,quantityFrom,quantityUntil,discountPerUnit,num,percentage,discountPerWUnit,baseproductonly) values (" & pIDProduct & "," & didcategory & "," & ddiscountDesc & "," & dquantityFrom & "," & dquantityUntil & "," & ddiscountPerUnit & "," & dnum & "," & dpercentage & "," & ddiscountPerWUnit & "," & dbaseproductonly & ")"
		set rs=connTemp.execute(strSQL)

		rstemp.MoveNext
	loop
	set rstemp=nothing
END IF 
'End Duplicate product

'// Duplicate Apparel Sub-Products
If statusAPP="1" OR scAPP=1 Then
	IF (pcv_Apparel="1") and (request("AddType")="19") and (request("PreProd")<>"") and (request("PreProd")<>"0") and (request("dupsubprd")="1") then
	
		query="select imageUrl,largeImageURL,stock,pcprod_Relationship,pcprod_Addprice,pcprod_AddWprice,nostock,pcProd_BackOrder,pcProd_ShipNDays,pcProd_NotifyStock,pcProd_ReorderLevel,pcProd_GoogleGender,pcProd_GoogleAge,pcProd_GoogleSize,pcProd_GoogleColor,pcProd_GooglePattern,pcProd_GoogleMaterial from products where pcprod_ParentPrd=" & request("PreProd") & " and removed=0 and active=0 ORDER BY IDProduct ASC"
		set rstemp=connTemp.execute(query)
		
		pcv_ParentPrd=pIDProduct
		pPrice1=pPrice
		pBtoBPrice1=pBtoBPrice
		if (pBtoBPrice1="") OR (cdbl(pBtoBPrice1)=0) then
			pBtoBPrice1=pPrice1
		end if
		
		psku1=psku
		pDescription1=pDescription
		pActive="0"
		
		do while not rstemp.eof
			pImageUrl=rstemp("imageUrl")
			pLargeImageUrl=rstemp("largeImageUrl")
			pstock=rstemp("stock")
			pcv_Relationship=rstemp("pcprod_Relationship")
			pAddPrice=rstemp("pcprod_addprice")
			if pAddPrice<>"" then
			else
				pAddPrice=0
			end if
			pAddWPrice=rstemp("pcprod_addWprice")
			if (pAddWPrice<>"") and (pAddWPrice<>"0") then
			else
				pAddWPrice="0"
			end if
			
			pPrice=cdbl(pPrice1)+cdbl(pAddPrice)
			if pPrice="" then
				pPrice="0"
			end if
			pListPrice="0"
			pBtoBPrice=cdbl(pBtoBPrice1)+cdbl(pAddWPrice)
			if (pBtoBPrice="") or (cdbl(pBtoBPrice)=0) then 
				pBtoBPrice=pPrice
			end If
			SP_NoStock=rstemp("nostock")
			if SP_NoStock="-1" then
				SP_NoStock=1
			end if
			SP_pcbackorder=rstemp("pcProd_BackOrder")
			SP_pcShipNDays=rstemp("pcProd_ShipNDays")
			SP_pcnotifystock=rstemp("pcProd_NotifyStock")
			SP_pcreorderlevel=rstemp("pcProd_ReorderLevel")
			
			pcv_GGen=rstemp("pcProd_GoogleGender")
			if pcv_GGen<>"" then
				pcv_GGen=replace(pcv_GGen,"'","''")
			end if
			pcv_GAge=rstemp("pcProd_GoogleAge")
			if pcv_GAge<>"" then
				pcv_GAge=replace(pcv_GAge,"'","''")
			end if
			pcv_GSize=rstemp("pcProd_GoogleSize")
			if pcv_GSize<>"" then
				pcv_GSize=replace(pcv_GSize,"'","''")
			end if
			pcv_GColor=rstemp("pcProd_GoogleColor")
			if pcv_GColor<>"" then
				pcv_GColor=replace(pcv_GColor,"'","''")
			end if
			pcv_GPat=rstemp("pcProd_GooglePattern")
			if pcv_GPat<>"" then
				pcv_GPat=replace(pcv_GPat,"'","''")
			end if
			pcv_GMat=rstemp("pcProd_GoogleMaterial")
			if pcv_GMat<>"" then
				pcv_GMat=replace(pcv_GMat,"'","''")
			end if
		
			pcv_TempArr=split(pcv_Relationship,"_")
			
			pSKU=pSku1
			pcv_Relationship=pIDProduct
			pDescription=" ("
			
			For i=1 to ubound(pcv_TempArr)
		
				pcv_Opt1=pcv_TempArr(i)
		
				query="select idOption from options_optionsGroups where idoptoptgrp=" & pcv_Opt1
				set rs=connTemp.execute(query)
		
				if not rs.eof then
					pcv_ROpt1=rs("idOption")
					query="select idoptoptgrp from options_optionsGroups where idOption=" & pcv_ROpt1 & " AND idproduct=" & pIDProduct
					set rs=connTemp.execute(query)
					pcv_Opt1=rs("idoptoptgrp")
				end if
		
				query="select optionDescrip,pcOpt_Code from Options where idOption=" & pcv_ROpt1
				set rs=connTemp.execute(query)
		
				if not rs.eof then
					pcv_OptName1=rs("optionDescrip")
					pcv_Code1=rs("pcOpt_Code")
				end if
			
				pSKU=pSKU & pcv_Code1
				pcv_Relationship=pcv_Relationship & "_" & pcv_Opt1
				pDescription=pDescription & pcv_OptName1
				if (i<ubound(pcv_TempArr)) then
					pDescription=pDescription & " - "
				end if
			Next
		
			pDescription=pDescription & ")"
		
			pDetails="Please see Product SKU #" & psku1 & " for more details"
			psDesc="Please see Product SKU #" & psku1 & " for more details"
			pDescription=replace(pDescription,"'","''")
			pDescription=pDescription1 & pDescription
			dtTodaysDate=Date()
			if SQL_Format="1" then
				dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
			else
				dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
			end if
			query="INSERT INTO products (sku, description, details, configOnly, serviceSpec, price, listPrice, bToBPrice, cost, imageUrl, listhidden, weight, stock, active, idSupplier, hotDeal, emailText, deliveringTime, formQuantity, smallImageUrl,largeImageUrl, notax, noshipping, showInHome,iRewardPoints,OverSizeSpec,sdesc,downloadable,noStock,noshippingtext,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty,pcprod_QtyToPound, pcprod_EnteredOn,pcProd_Apparel,pcProd_ParentPrd,pcProd_RelationShip,pcProd_ShowStockMsg,pcProd_StockMsg,pcProd_SizeLink,pcProd_SizeInfo,pcProd_SizeImg,pcProd_SizeURL,pcProd_AddPrice,pcProd_AddWPrice,pcProd_BackOrder,pcProd_ShipNDays,pcProd_NotifyStock,pcProd_ReorderLevel,pcProd_GoogleGender,pcProd_GoogleAge,pcProd_GoogleSize,pcProd_GoogleColor,pcProd_GooglePattern,pcProd_GoogleMaterial) VALUES (N'" &pSku& "',N'" &pDescription& "',N'" & pDetails& "'," &pconfigOnly& "," &pserviceSpec& "," &pPrice& "," &pListPrice& "," &pBtoBPrice& "," &pCost& ",'" &pImageUrl& "'," &pListhidden& "," &pWeight& "," &pStock& "," &pActive& "," &pIdSupplier& "," &pHotDeal& ",N'" &pEmailText& "'," &pDeliveringTime& "," &pFormQuantity& ",'" &pSmallImageUrl& "','"&pLargeImageUrl&"',"&pnotax&","&pnoshipping&","&pshowInHome&","&iRewardPoints&",'"&pOverSizeSpec&"',N'" & psdesc & "'," & pdownloadable & ","&SP_NoStock&","&pnoshippingtext&"," & pcv_intHideBTOPrice & "," & pcv_intQtyValidate & "," & pcv_lngMinimumQty & ","&pcv_QtyToPound&",'"&dtTodaysDate&"',0," & pcv_ParentPrd & ",'" & pcv_RelationShip & "'," & pcv_ShowStockMsg & ",N'" & pcv_StockMsg & "','" & pcv_SizeLink & "',N'" & pcv_SizeInfo & "','" & pcv_SizeImg & "','" & pcv_SizeURL & "'," & pAddPrice & "," & pAddWPrice & "," & SP_pcbackorder & "," & SP_pcShipNDays & "," & SP_pcnotifystock & "," & SP_pcreorderlevel & ",'" & pcv_GGen & "','" & pcv_GAge & "','" & pcv_GSize & "','" & pcv_GColor & "','" & pcv_GPat & "','" & pcv_GMat & "')"
			set rs=conntemp.execute(query)
			
			rstemp.MoveNext
			
		Loop
	
		'// Update parent products inventory levels if necessary 
		%>
		<!--#include file="../pc/app-updstock.asp"-->
		<%
	End If
END IF


'// Insert downloadable product to DProducts Table
if (pdownloadable<>"") and (pdownloadable="1") then
	query="Insert into DProducts (IdProduct,ProductURL,URLExpire,ExpireDays,License,LocalLG,RemoteLG,LicenseLabel1,LicenseLabel2,LicenseLabel3,LicenseLabel4,LicenseLabel5,AddToMail) values (" & pIdProduct & ",'" & pproducturl & "'," & pURLExpire & "," & pExpireDays & "," & pLicense & ",'" & pLocalLG & "','" & pRemoteLG & "',N'" & plicenselabel1 & "',N'" & plicenselabel2 & "',N'" & plicenselabel3 & "',N'" & plicenselabel4 & "',N'" & plicenselabel5 & "','" & pAddtoMail & "')"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
	if err.number <> 0 then
		
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error adding downloadable product information on addProductB.asp") 
	end if
end if


'GGG add-on start
	if (pGC<>"") and (pGC="1") then
		if SQL_Format="1" then
			pGCExpDate=(day(pGCExpDate)&"/"&month(pGCExpDate)&"/"&year(pGCExpDate))
		else
			pGCExpDate=(month(pGCExpDate)&"/"&day(pGCExpDate)&"/"&year(pGCExpDate))
		end if

		query="Insert into pcGC (pcGC_IdProduct,pcGC_Exp,pcGC_ExpDate,pcGC_ExpDays,pcGC_EOnly,pcGC_CodeGen,pcGC_GenFile) values (" & pIdProduct & "," & pGCExp & ",'" & pGCExpDate & "'," & pGCExpDay & "," & pGCEOnly & "," & pGCGen & ",'" & pGCGenFile & "')"

		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
		if err.number <> 0 then
			
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error adding gift certificate information on addProductB.asp") 
		end if
	end if
'GGG add-on end

' insert categories for the item
CATS=split(pIdcategory1,", ")
for dd=lbound(CATS) to ubound(CATS)
	if CATS(dd)<>"" then
		query="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pIdProduct& "," &CATS(dd)& ")"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
		if err.number <> 0 then
			
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error adding category information on addProductB.asp") 
		end if
	end if
next

Call RunCalBDPC()
%>

<!--#include file="AdminHeader.asp"-->

<table class="pcCPcontent">
	<tr>
		<td colspan="2">
			<div class="pcCPmessageSuccess">&quot;<%=removeSQ(pDescription)%>&quot; was successfully added 
<%if (request("dupsubprd")="1") then%>(included sub-products)<%end if%>. <a href="FindProductType.asp?id=<%=pIdProduct%>"><strong>Edit it</strong></a>.</div>
		</td>
	</tr>
	<tr>
		<td width="50%" valign="top">
		<% if DupSKU=1 then %>
				<p><span class="pcCPnotes">Warning: DUPLICATE SKU - The part number (SKU) that you have entered already exists in the database. It is up to you to decide whether to keep it 'as is' or to change it to a unique one. To edit it, select 'Modify this product' from the menu below.</span></p>
		<% end if %>
			<ul class="pcListIcon">
				<% if pcv_ProductType<>"item" then %>
				<li><a href="../pc/viewPrd.asp?idproduct=<%=pIdProduct%>&adminPreview=1" target="_blank">Preview in your storefront &gt;&gt;</a></li>
				<li style="padding-bottom: 10px;"><a href="FindProductType.asp?id=<%=pIdProduct%>">Modify this product</a></li>
				<% end if %>
				
				<% if pcv_ProductType="std" then %>
					<li><a href="modPrdOpta.asp?idproduct=<%=pIdProduct%>">Add/Modify product options</a> (e.g. sizes, colors, etc.)</li>
					<% if pcv_Apparel="1" then %>
						<li><a href="app-subPrdsMngAll.asp?idproduct=<%=pIdProduct%>">Create &amp; manage sub-products</a> </li>
					<% end if %>
				<% elseif pcv_ProductType="bto" then %>
					<li style="padding-top: 10px"><a href="modBTOconfiga.asp?idProduct=<%=pIdProduct%>"><strong>Setup</strong> this configurable product or service</a></li>
				<% end if %>
				
				<% if pcv_ProductType<>"item" then %>
				<li><a href="AdminCustom.asp?idproduct=<%=pIdProduct%>">Add/Modify custom search or input fields</a></li>
				<li><a href="crossSellAddb.asp?action=source&prdlist=<%=pIdProduct%>">Add a cross-selling relationship</a></li>
				<% end if %>
				
				<li><a href="FindProductQtyDisc.asp?idproduct=<%=pIdProduct%>">View/Add quantity discounts</a></li>
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