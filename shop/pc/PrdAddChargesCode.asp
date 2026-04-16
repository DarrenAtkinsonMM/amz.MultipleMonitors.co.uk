<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

Dim TempStr1, TempQDStr, QFrom, QTo, DUnit, QPercent, DWUnit, pcv_IncOpt, DIDProduct, TempDiscountStr
Dim MyGCodes,ReqTestStr
Dim pcv_CustomizedPrice,pcv_ItemDiscounts,pcv_tmpIDiscount,pcv_tmpIDiscount1,pcv_tmpCustomizedPrice
Dim pcv_ListForGenInfo
Dim pNostock,pcv_intBackOrder,pcv_intShipNDays,pMinPurchase

pcv_ListForGenInfo=""
pcv_CustomizedPrice=0
pcv_tmpIDiscount1=0
pcv_ItemDiscounts=0
pcv_tmpIDiscount=0
pcv_tmpCustomizedPrice=0

pcv_sffolder="../pc/"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Enhanced Views Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim pcv_strUseEnhancedViews, pcv_strHighSlide_Align, pcv_strHighSlide_Template
Dim pcv_strHighSlide_Eval, pcv_strHighSlide_Effects, pcv_strHighSlide_MinWidth, pcv_strHighSlide_MinHeight

pcv_strUseEnhancedViews = True '// Turn Enhanced Views ON or OFF
pcv_strHighSlide_Align = "center" '// Align Images from anchor or screen
pcv_strHighSlide_Template = "rounded-white" '// Template
pcv_strHighSlide_Eval = "this.thumb.alt"
pcv_strHighSlide_Effects = "'expand', 'fade'"
pcv_strHighSlide_MinWidth = 250
pcv_strHighSlide_MinHeight = 250
pcv_strHighSlide_Fade = "true"
pcv_strHighSlide_Dim = 0.3
pcv_strHighSlide_Interval = 3500
pcv_strHighSlide_Heading = "highslide-caption" '// "highslide-heading"
pcv_strHighSlide_Hide = "true"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Enhanced Views Configuration
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="pcCheckPricingCats.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Disallow purchasing. Quote Submission only 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_AddChargesSubmission
	%>
		<div class="row">
		    <div class="col-xs-12"> 
    <%
    IF (iBTOQuoteSubmitOnly=1) or (pnoprices>0) THEN
		if pQty="1" then ' hide the information if the quantity = 1
	%>

    		<input type="hidden" name="quantity" value="<%=pQty%>">
    <%
		else
	%>

            <div style="padding:5px;">
            <% response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_15")%> 
            <input type="text" name="quantity" size="7" maxlength="10" value="<%=pQty%>" readonly class="transparentField">
            </div>
	<%
		end if
	ELSE
		if pQty="1" then ' hide the information if the quantity = 1		
	%>

        <input type="hidden" name="quantity" value="<%=pQty%>">
        <button class="pcButton pcButtonAddToCart" id="add" name="add">
            <img src="<%=pcf_getImagePath("",rslayout("addtocart"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_addtocart")%>" />
            <span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_addtocart")%></span>
        </button>

	<%
		else
	%>
    		<hr />
			<div class="pcFormItem">
				<div class="pcFormLabel"><% response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_15")%></div>
				<div class="pcFormField"><input type="text" name="quantity" size="7" maxlength="10" value="<%=pQty%>" readonly class="transparentField"></div>
			</div>
			<button class="pcButton pcButtonAddToCart" id="add" name="add">
				<img src="<%=pcf_getImagePath("",rslayout("addtocart"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_addtocart")%>" />
				<span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_addtocart")%></span>
			</button>
	<% 
		end if
	END IF
    %>
  </div>
</div>   
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Disallow purchasing. Quote Submission only 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Totals
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_AddChargesTotals
%>
<div id="pcBTOfloatPrices">
<div class="pcTable">
	<div class="pcTableRowFull">
		<div class="pcTableColumn60">
			<b><% response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_4")%></b>
		</div>
		<div class="pcTableColumn1"></div>
		<div class="pcTableColumn39"> 
			<input name="curPrice" type="TEXT" style="text-align:right;" value="<%=scCurSign & money(pPriceDefault) %>"  readonly size="10" class="transparentField">
			<input type="hidden" name="TLPriceDefault" value="0">
			<input type="hidden" name="TLcurPrice" value="0">
			<input type="hidden" name="TLdefaultprice" value="0">
			<input type="hidden" name="TLtotal" value="0">
			<input type="hidden" name="Discounts" value="0">
			<input type="hidden" name="QDiscounts0" value="0">
			<input type="hidden" name="QDiscounts" value="0">
			<input type="hidden" name="TotalWithQD" value="0">
			<input type="hidden" name="TLGrandTotal2QD" value="0">
			<input type="hidden" name="CMDefault" value="0">
			<input type="hidden" name="CMWQD" value="0">
			<input type="hidden" name="TLGrandTotal" value="0">
			<input type="hidden" name="TLGrandTotal2" value="0">
			<input type="hidden" name="GrandTotal2" value="0">
		</div>
	</div> 
	
	<div class="pcTableRowFull">
		<div class="pcTableColumn60">
			<% if pcv_strAdminPrefix="1" then %>
				<%if request.QueryString("idquote")<>"" then%>
				<b><% response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_18A")%></b>
				<%else%>
				<b><% response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_18")%></b>
				<%end if%>
			<% else %>
				<b><% response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_16")%></b>
			<% end if %>
		</div>
		<div class="pcTableColumn1"></div>
		<div class="pcTableColumn39">
		<input name="CMPrice0" type="HIDDEN" value="<%=New_ConvertNum(pCMPrice)%>">
		<input name="CMWQD0" type="HIDDEN" value="<%=New_ConvertNum(pCMWQD)%>">		
		<input name="CMWQD" type="TEXT" style="text-align:right;" value="<%=scCurSign & money(pCMWQD)%>" readonly size="10" class="transparentField">
		<input name="CMPrice" type="hidden" value="<%=scCurSign & money(pCMPrice)%>">		
		</div>
	</div>
	
	<div class="pcTableRowFull">
		<div class="pcTableColumn60">
		<input name="currentValue0" type="HIDDEN" value="<%=New_ConvertNum(pPriceDefault)%>">
		<input name="jCnt" type="HIDDEN" value="<%=jCnt%>">
		<b><% response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_14")%></b> 	
		</div>
		<div class="pcTableColumn1"></div>
		<div class="pcTableColumn39"> 		
		<input name="total" type="TEXT" style="text-align:right;" value="None"  readonly size="10" class="transparentField">
		<input name="CHGTotal" type="hidden" value="0">		
		</div>
	</div>
	
	<div class="pcTableRowFull">
		<div class="pcTableColumn60">
			<b><% response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_6")%></b>
		</div>
		<div class="pcTableColumn1"></div>
		<div class="pcTableColumn39">
		<input name="GrandTotalQD" type="TEXT" style="text-align:right;" value="<%=scCurSign & money(pCMWQD+pPriceDefault)%>"  readonly size="10" class="transparentField">
		<input name="GrandTotal" type="hidden" value="<%=scCurSign & money(pCMPrice+pPriceDefault)%>">		
		</div>
	</div>	
</div>
</div>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Totals
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Disallow purchasing. Quote Submission only - Reconfig
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_SubmissionReconfig
%>
<div class="row">
    <div class="col-xs-12">
    
        <% 
        if pSavedQuantity="1" then ' hide if the quantity = 1
        %>
            <input type="hidden" name="quantity" value="<%=pSavedQuantity%>">
        <%
        else
        %>
            <div class="pcFormItem">
                <div class="pcFormLabel"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_15")%></div>
                <div class="pcFormField"><input type="text" name="quantity" size="7" maxlength="10" value="<%=pSavedQuantity%>" readonly class="transparentField"></div>
            </div>
        <%
        end if
        %>
                
        <% IF (pnoprices=0) AND (request("act")="placeOrder") THEN %>
            <button class="pcButton pcButtonAddToCart" id="add" name="add">
                <img src="<%=pcf_getImagePath("",rslayout("addtocart"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_addtocart")%>" />
                <span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_addtocart")%></span>
            </button>
        <% ELSE
            IF (pnoprices=0) THEN %>
                <button class="pcButton pcButtonSubmit" id="add" name="add">
                    <img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_submit")%>" />
                    <span class="pcButtonText"><%=dictLanguage.Item(Session("language")&"_css_submit")%></span>
                </button>
            <% END IF
        END IF %>
    
    </div>
</div>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Disallow purchasing. Quote Submission only - Reconfig
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Configuration Table
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_AddChargesTable
Dim query,rsSSObj,tmpquery
	tmpquery=""
	if (scOutOfStockPurchase="-1") AND (iBTOOutofStockPurchase="-1") then
		tmpquery=" AND ((products.stock>0) OR (products.nostock<>0) OR (products.pcProd_BackOrder<>0))"
	end if
	query="SELECT categories.idCategory, categories.categoryDesc, configSpec_Charges.multiSelect,products.pcprod_qtyvalidate,products.pcprod_minimumqty,products.idproduct, products.weight, products.description, configSpec_Charges.prdSort, configSpec_Charges.price, configSpec_Charges.Wprice, configSpec_Charges.showInfo, configSpec_Charges.cdefault, configSpec_Charges.requiredCategory, configSpec_Charges.displayQF,configSpec_Charges.pcConfCha_ShowDesc,configSpec_Charges.pcConfCha_ShowImg,configSpec_Charges.pcConfCha_ImgWidth,configSpec_Charges.pcConfCha_ShowSKU,products.sku,products.smallImageUrl,products.stock,products.noStock, products.pcProd_BackOrder, products.pcProd_ShipNDays,products.pcprod_minimumqty,configSpec_Charges.pcConfCha_UseRadio,products.details,products.sDesc,configSpec_Charges.Notes FROM categories INNER JOIN (products INNER JOIN configSpec_Charges ON (products.idproduct=configSpec_Charges.configProduct AND products.active<>0 AND products.removed=0" & tmpquery & ")) ON categories.idCategory = configSpec_Charges.configProductCategory WHERE configSpec_Charges.specProduct="&pIdProduct&" ORDER BY configSpec_Charges.catSort, categories.idCategory, configSpec_Charges.prdSort,products.description;"
	tmpquery=""
	set rsSSObj=conntemp.execute(query)
	displayQF="0"
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsSSObj=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
%>
<div class="">
	<% 
	CB_CatCnt = 0
	jcnt=0
	'*******************************************
	'******* START BTO Categories

	IF NOT rsSSobj.eof then  
		Dim strCol
		strCol = "class='pcBTOsecondRow row'"
		checkVar=0
		checkVarCat=0

		pcv_tmpArr=rsSSobj.GetRows()
		pcv_ArrCount=ubound(pcv_tmpArr,2)
		set rsSSobj=nothing

		'*********** LOOP CATs
						
		pcv_tmpN=0

		DO WHILE (pcv_tmpN<=pcv_ArrCount)
		
		tempVarCat = pcv_tmpArr(0,pcv_tmpN)
		VarMS=pcv_tmpArr(2,pcv_tmpN)
		
		If VarMS=False then 
			dim defaultPrice
			defaultPrice=Cdbl(0)
			dim cdVar
			cdVar="0"
			
			'**** IT IS NEW CAT
			If Clng(tempVarCat) <> Clng(checkVar) then
			    %>
                <div class="panel panel-default">
                    <%	
                    checkVar = tempVarCat
                    strCategoryDesc=pcv_tmpArr(1,pcv_tmpN)
                
                    pcv_ShowDesc="0"
                    pClngShowItemImg="0"
                    pClngSmImgWidth="0"
                    pClngShowSku="0"
                    
                    if pcv_tmpArr(15,pcv_tmpN)="1" then
                        pcv_ShowDesc="1"
                    end if
                    if pcv_tmpArr(16,pcv_tmpN)="1" then
                        pClngShowItemImg="1"
                    end if
                    if pcv_tmpArr(17,pcv_tmpN)>"0" then
                        pClngSmImgWidth=pcv_tmpArr(17,pcv_tmpN)
                    end if
                    if pcv_tmpArr(18,pcv_tmpN)="1" then
                        pClngShowSku="1"
                    end if
                    
                    '***** GET DEFAULT PRICE OF THE CAT
					pcv_minqty=1
					query="SELECT configSpec_Charges.configProduct,configSpec_Charges.price, configSpec_Charges.Wprice, configSpec_Charges.cdefault FROM configSpec_Charges WHERE configSpec_Charges.configProductCategory="&tempVarCat&" AND configSpec_Charges.specProduct="&pIdProduct&" AND configSpec_Charges.cdefault<>0;"
                    set rsTempObj=conntemp.execute(query)
                    if err.number<>0 then
                        call LogErrorToDatabase()
                        set rsTempObj=nothing
                        call closedb()
                        response.redirect "techErr.asp?err="&pcStrCustRefID
                    end if
        
                    If NOT rsTempObj.eof then
                        cdVar="1"
                        tmpintPrd=rsTempObj("configProduct")
                        dblprice=Cdbl(rsTempObj("price"))
                        dblWprice=Cdbl(rsTempObj("Wprice"))
                        if dblWprice=0 then
                            dblWprice=dblprice
                        end if
                        
                        query="SELECT products.pcprod_minimumqty FROM Products WHERE idproduct=" & tmpintPrd & ";"
                        set rsQ=connTemp.execute(query)
                        if not rsQ.eof then
                            pcv_minqty=rsQ("pcprod_minimumqty")
                            if IsNull(pcv_minqty) or pcv_minqty="" then
                                pcv_minqty=1
                            end if
                            if pcv_minqty="0" then
                                pcv_minqty=1
                            end if
                        else
                            pcv_minqty=1
                        end if
                        set rsQ=nothing
                        
                        intCC_BTO_Pricing=0
                        if session("customercategory")<>0 then
                            query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idcustomerCategory=" & session("customercategory") & " AND idBTOItem=" & tmpintPrd & " AND idBTOProduct=" & pIdProduct & ";" 
                            set rsCCObj=server.CreateObject("ADODB.RecordSet")
                            set rsCCObj=conntemp.execute(query)
                                                                
                            if err.number<>0 then
                                call LogErrorToDatabase()
                                set rsCCObj=nothing
                                call closedb()
                                response.redirect "techErr.asp?err="&pcStrCustRefID
                            end if
                                                                                                    
                            if NOT rsCCObj.eof then
                                intCC_BTO_Pricing=1
                                pcCC_BTO_Price=rsCCObj("pcCC_BTO_Price")
                            else
                                query="SELECT pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=" & session("customercategory") &" AND pcCC_Pricing.idProduct=" & tmpintPrd & ";"
                                set rsCCObj=server.CreateObject("ADODB.RecordSet")
                                set rsCCObj=conntemp.execute(query)
                                if NOT rsCCObj.eof then
                                    intCC_BTO_Pricing=1
                                    pcCC_BTO_Price=rsCCObj("pcCC_Price")
                                end if
                            end if
                            set rsCCObj=nothing
                        end if
																		
                        'customer logged in as ATB customer based on retail price
                        if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
                            dblprice=Cdbl(dblprice)-(pcf_Round(Cdbl(dblprice)*(cdbl(session("ATBPercentage"))/100),2))
                        end if
                        defaultPrice= Cdbl(dblprice)
                        
                        'customer logged in as ATB customer based on wholesale price
                        if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
                            dblWprice=Cdbl(dblWprice)-(pcf_Round(Cdbl(dblWprice)*(cdbl(session("ATBPercentage"))/100),2))
                            defaultPrice=Cdbl(dblWprice)
                        end if
                        
                        'customer logged in as a wholesale customer
                        if dblWprice>0 and session("customerType")=1 then
                            defaultPrice=Cdbl(dblWprice)
                        end if
                        
                        'customer logged in as a customer type with price different then the online price
                        if intCC_BTO_Pricing=1 then
                            if (pcCC_BTO_Price<>0) OR (pcCC_BTO_Price=0 AND intCC_BTO_Pricing=1) then
                                defaultPrice=Cdbl(pcCC_BTO_Price)
                            end if
                        end if
                        
                        defaultPrice=defaultPrice*pcv_minqty
                    end if
					Set rsTempObj=nothing
				    '***** END OF GET DEFAULT PRICE OF THE CAT
				
					jcnt=jCnt+1
					If strCol <> "class='pcBTOfirstRow row'" Then
						strCol = "class='pcBTOfirstRow row'"
					Else 
						strCol = "class='pcBTOsecondRow row'"
					End If
					%>
                    
					<div class="panel-heading"><%=pcv_tmpArr(1,pcv_tmpN)%>
					</div>
                    <div class="panel-body">
					
					<%
					' If there are configuration instructions for this category, show them here.
					CATNotes=pcv_tmpArr(29,pcv_tmpN)
					if CATNotes <> "" then
					%>
					<div <%=strCol%>>
						<div class="col-xs-12"><span class="catNotes"><%=CATNotes%></span></div>
					</div>
					<%
                    end if
                    
					pBTODisplayType=pcv_tmpArr(26,pcv_tmpN)
					if IsNull(pBTODisplayType) or pBTODisplayType="" then
						pBTODisplayType=1
					end if
					
					displayQF=pcv_tmpArr(14,pcv_tmpN)
					requiredCategory=pcv_tmpArr(13,pcv_tmpN)
					if pcv_tmpNewPath<>"" then
						pcv_tmpArr(11,pcv_tmpN)=0
					end if
					showInfo=pcv_tmpArr(11,pcv_tmpN)%>
					
                    <%
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    ' START: Show Dropdown
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    %>
                                 
                    <% 
                    if pBTODisplayType=1 then
						if (requiredCategory<>0) and (cdVar<>"1") then
                        	pcv_ListForGenInfo=pcv_ListForGenInfo & "GenDropInfo(document.additem.CAG" & tempVarCat & ");" & vbcrlf
						end if
                        %>
                        <div <%=strCol%>>
                            <% if (displayQF=True) then %>
                                <div class="col-xs-2">
									<input class="form-control quantity" type="text" size=2 name="CAG<%=tempVarCat%>QF" value="<%=pcv_minqty%>" onBlur="javascript:testdropqty(this,'document.additem.CAG<%=tempVarCat%>');">
                                </div>
                                <div class='col-xs-7'>
                            <%else%>
                                <div class="col-xs-9"> 
                            	<input type="hidden" name="CAG<%=tempVarCat%>QF" value="<%=pcv_minqty%>">
                            <%end if%>
                                <select class="form-control" name="CAG<%=tempVarCat%>" onChange="testdropdown('document.additem.CAG<%=tempVarCat%>'); calculate(this,0); showAvail<%=tempVarCat%>(this);">
                            </div>
                            <div class="col-xs-3">
                                <%
                                HiddenFields=""
                    else%>
                        <input type="hidden" name="CAG<%=tempVarCat%>QF" value="<%=pcv_minqty%>">
                        <%if Clng(requiredCategory)<>0 then
                            RTestStr="totalradio=document.additem.CAG" & tempVarCat & ".length;" & vbcrlf
                            RTestStr=RTestStr & "RadioChecked=0;" & vbcrlf
                            RTestStr=RTestStr & "if (totalradio>0) {" & vbcrlf
                            RTestStr=RTestStr & "for (var mk=0;mk<totalradio;mk++) {" & vbcrlf
                            RTestStr=RTestStr & "if (document.additem.CAG" & tempVarCat & "[mk].checked==true) { RadioChecked=1; break; } }" & vbcrlf
                            RTestStr=RTestStr & "} else { if (document.additem.CAG" & tempVarCat & ".checked==true) RadioChecked=1;}" & vbcrlf
                            RTestStr=RTestStr & "if (RadioChecked==0) {alert('"& dictLanguage.Item(Session("language")&"_alert_7") & replace(pcv_tmpArr(1,pcv_tmpN),"'","\'") & "'); return(false);}" & vbcrlf
                            ReqTestStr=ReqTestStr & RTestStr
                        end if%>
                    <%end if %>
                    
                    <%
					dim requiredVar, showInfoVar, ShowInfoArray
                    requiredVar="0"
                    showInfoVar="0"
                    ShowInfoArray = ""
																		if requiredCategory=False then
																			requiredVar = "1"
																		end if
																		if showInfo=True then
																			showInfoVar = "1"
																		end if
																		icount=0

																		pcv_tmpIDiscount=0
									
																		pcv_tmpTest=1
																		
																		pcv_FirstItem=1
																		pcv_tmpDefaultValue=0
                    intOpCnt = 0
                    StrBackOrd = "var availArr"&tempVarCat &" = new Array();" &vbcrlf
                    strselectvalue = "" 
                    DO WHILE ((pcv_tmpTest=1) AND (pcv_tmpN<=pcv_ArrCount))
                        if pBTODisplayType<>1 then
                        ShowInfoArray = ""%>
                            <div <%=strCol%>>
                        <%end if
                        icount=icount+1
																			pcv_prdDesc=pcv_tmpArr(27,pcv_tmpN)
																			pcv_prdSDesc=pcv_tmpArr(28,pcv_tmpN)
																			if IsNull(pcv_prdSDesc) or trim(pcv_prdSDesc)="" then
																				pcv_prdSDesc=pcv_prdDesc
																			end if
																			displayQF=pcv_tmpArr(14,pcv_tmpN)
                        pcv_qtyvalid=pcv_tmpArr(3,pcv_tmpN)
                        if isNULL(pcv_qtyvalid) OR pcv_qtyvalid="" then
                            pcv_qtyvalid="0"
                        end if
                        pcv_minQty=pcv_tmpArr(4,pcv_tmpN)
                        if isNULL(pcv_minQty) OR pcv_minQty="" then
                            pcv_minQty="1"
                        end if
                        if pcv_minQty<"1" then
                            pcv_minQty="1"
                        end if
																			intTempIdProduct=pcv_tmpArr(5,pcv_tmpN)
																			intTempIdCategory=pcv_tmpArr(0,pcv_tmpN)
																			cdefault=pcv_tmpArr(12,pcv_tmpN)
																			weight=pcv_tmpArr(6,pcv_tmpN)
                        prdBtoBPrice = Cdbl(pcv_tmpArr(10,pcv_tmpN))
                        prdPrice = Cdbl(pcv_tmpArr(9,pcv_tmpN))
                        if prdBtoBPrice=0 then
                            prdBtoBPrice=prdPrice
                        end if
                        strDescription=pcv_tmpArr(7,pcv_tmpN)
                        strSku=pcv_tmpArr(19,pcv_tmpN)
                        strSmallImage=pcv_tmpArr(20,pcv_tmpN)							
                        if strSmallImage = "" or strSmallImage = "no_image.gif" then
                            strSmallImage = "hide"
                        end if
                        pstock=pcv_tmpArr(21,pcv_tmpN)
                        pNostock=pcv_tmpArr(22,pcv_tmpN)	
                        if pNostock = "" or pNoStock = null then
                         pNostock = 0
                        end if						
                        pcv_intBackOrder = pcv_tmpArr(23,pcv_tmpN)							
                        pcv_intShipNDays = pcv_tmpArr(24,pcv_tmpN)
                        pMinPurchase = pcv_tmpArr(25,pcv_tmpN)
																			
                        intCC_BTO_Pricing=0																
                        if session("customercategory")<>0 then
                            query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idcustomerCategory=" & session("customercategory") & " AND idBTOItem=" & pcv_tmpArr(5,pcv_tmpN)& " AND idBTOProduct=" & pIdProduct & ";" 
                            set rsCCObj=server.CreateObject("ADODB.RecordSet")
                            set rsCCObj=conntemp.execute(query)
                                                                                                                            
                            if err.number<>0 then
                                call LogErrorToDatabase()
                                set rsCCObj=nothing
                                call closedb()
                                response.redirect "techErr.asp?err="&pcStrCustRefID
                            end if
                                                                            
                            if NOT rsCCObj.eof then
                                intCC_BTO_Pricing=1
                                pcCC_BTO_Price=rsCCObj("pcCC_BTO_Price")
                            else
                                query="SELECT pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=" & session("customercategory") &" AND pcCC_Pricing.idProduct=" & pcv_tmpArr(5,pcv_tmpN) & ";"
                                set rsCCObj=server.CreateObject("ADODB.RecordSet")
                                set rsCCObj=conntemp.execute(query)
                                if NOT rsCCObj.eof then
                                    intCC_BTO_Pricing=1
                                    pcCC_BTO_Price=rsCCObj("pcCC_Price")
                                end if
                            end if
                            SET rsCCObj=nothing
                        end if
                                                
                        'customer logged in as ATB customer based on retail price
                        if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
                            prdPrice=Cdbl(prdPrice)-(pcf_Round(Cdbl(prdPrice)*(cdbl(session("ATBPercentage"))/100),2))
                        end if
                                            
                        'customer logged in as ATB customer based on wholesale price
                        if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
                            prdBtoBPrice=Cdbl(prdBtoBPrice)-(pcf_Round(Cdbl(prdBtoBPrice)*(cdbl(session("ATBPercentage"))/100),2))
                            prdPrice=Cdbl(prdBtoBPrice)
                        end if
                        
                        'customer logged in as a wholesale customer
                        if prdBtoBPrice>0 and session("customerType")=1 then
                            prdPrice=Cdbl(prdBtoBPrice)
                        end if

                        'customer logged in as a customer type with price different then the online price
                        if intCC_BTO_Pricing=1 then
                            if (pcCC_BTO_Price<>0) OR (pcCC_BTO_Price=0 AND intCC_BTO_Pricing=1) then
                                prdPrice=Cdbl(pcCC_BTO_Price)
                            end if
                        end if
																			
																			tmp_qty=pcv_minQty*ProQuantity
																			
																			'if (requiredCategory<>0) and (cdVar<>"1") and (pcv_FirstItem=1) and (pBTODisplayType=1) then
																			if (cdVar="1") and (pcv_FirstItem=1) then
																				call CheckDiscount(pcv_tmpArr(5,pcv_tmpN),true,tmp_qty,prdPrice)
																				pcv_FirstItem=3
																				'pcv_tmpDefaultValue=prdPrice*tmp_qty
																				pcv_tmpDefaultValue=defaultPrice
																				pcv_CustomizedPrice=pcv_CustomizedPrice+cdbl(pcv_tmpDefaultValue)
																				pcv_tmpDefaultDiscount=pcv_tmpIDiscount
																			end if
																			call CheckDiscount(pcv_tmpArr(5,pcv_tmpN),pcv_tmpArr(12,pcv_tmpN),tmp_qty,prdPrice)
																			if cdefault=true then
																				pExt = " "
																				ShowInfoArray = ShowInfoArray & intTempIdProduct& ","
																				prdPrice1=prdPrice%>
																				<% if pBTODisplayType=1 then %>
                                    <div class="col-xs-9">  
                                        <option value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=weight%>_<%=prdPrice1%>" selected><%=strDescription%></option>
																					
                                        <% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
                                        strselectvalue = func_DisplayBOMsg																				
                                        HiddenFields=HiddenFields & "<input type=hidden name=""CAG" & tempVarCat & intTempIdProduct & "HF"" value=""" & pcv_qtyValid & "_" & pcv_minQty & """>" & vbcrlf
                                else %>
                                    <%if (displayQF=True) then%>
                                        <div class="col-xs-3">
										<input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=weight%>_<%=prdPrice1%>" checked onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; calculate(this,0);" class="clearBorder">
										<input class="form-control quantity" type="text" size="2" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="<%=pcv_minQty%>" onblur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>)) calculate(document.additem.CAG<%=tempVarCat%>,2);">
										<% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                                    <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                                <% end if %>
                                        </div>
                                        <div class="col-xs-6">
                                                <span><%=strDescription%></span>
												<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<% if pnoprices<2 then %>TEXT<% else %>Hidden<% end if %>" value="" readonly size="<%=len(pExt)%>" class="transparentField"><br>
                                                <% if not pClngShowSku = 0 then %>
                                                    <div class="pcSmallText"><%=strSku%></div>
                                                <% end if %>
                                    <%else%>
                                        <div class="col-xs-2">
												<input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=weight%>_<%=prdPrice1%>" checked onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; calculate(this,0);" class="clearBorder"><input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="<%=pcv_minQty%>">
                                                <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                                    <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                                <% end if %>
                                        </div>
                                        <div class="col-xs-7">
                                                <span><%=strDescription%></span>
												<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<% if pnoprices<2 then %>TEXT<% else %>Hidden<% end if %>" value="" readonly size="<%=len(pExt)%>" class="transparentField"><br>
																					
                                                <% if not pClngShowSku = 0 then %>
                                                    <div class="pcSmallText"><%=strSku%></div>
                                                <% end if %>
                                        <%end if%>
                                        <%=func_DisplayBOMsg1(tempVarCat)%>
                                    <%if pcv_ShowDesc="1" then%>
                                            <div class="row">
                                                <div class="col-xs-12"><span class="configDesc"><%=pcv_prdSDesc%></span></div>
                                            </div>
                                    <%end if%>
                                <% end if %>
                            <%'DEFAULT BUT NOT SELECTED
                            else %>
							<%
                            pExt = " "
                            prdPrice1=prdPrice
																				ShowInfoArray = ShowInfoArray & intTempIdProduct& ","
																				
                            If prdPrice=Cdbl(defaultPrice) then
                                prdPrice=0
                            Else
                                prdPrice=prdPrice-Cdbl(defaultPrice)
                            End if
                            
                            tmp_price=prdPrice+(tmp_qty-1)*prdPrice1-pcv_tmpIDiscount1
                            
                            if pnoprices<2 then
                                If tmp_price>0 then
                                    pExt = " - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(tmp_price)
                                Else
                                    If tmp_price<0 then
                                        pExt = " - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*tmp_price)
                                    End if
                                End if
                            End If
																				
                            If scDecSign="," then
                                prdPrice=replace(prdPrice,",",".")
                                prdPrice1=replace(prdPrice1,",",".")
                            End If 
                                if pBTODisplayType=1 then %>
                                    <div class="col-xs-9">
                                        <option value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=weight%>_<%=prdPrice1%>"><%=strDescription&pExt%></option>
                                        <%
                                        strBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
                                        HiddenFields=HiddenFields & "<input type=hidden name=""CAG" & tempVarCat & intTempIdProduct & "HF"" value=""" & pcv_qtyValid & "_" & pcv_minQty & """>" & vbcrlf
                                else %>
                                
                                    <% if (displayQF=True) then %>
                                    
                                        <div class="col-xs-3">
											<input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=weight%>_<%=prdPrice1%>" onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>';calculate(this,0);" class="clearBorder">
											<input class="form-control quantity" type="text" size="2" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="0" onblur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>)) calculate(document.additem.CAG<%=tempVarCat%>,2);">&nbsp;
                                            <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                                <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                            <% end if %>
                                        </div>
                                        <div class="col-xs-6">
                                            <span><%=strDescription%></span>
											<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%=pExt%>" readonly size="<%=len(pExt)%>" class="transparentField">
                                            <% if not pClngShowSku = 0 then %>
                                                <div class="pcSmallText"><%=strSku%></div>
                                            <% end if %>
																						
                                    <%else%>
                                        <div class="col-xs-2">
											<input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=weight%>_<%=prdPrice1%>" onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; calculate(this,0);" class="clearBorder"><input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="0">
                                            <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                                <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                            <% end if %>
                                        </div>
                                        <div class="col-xs-7">
                                            <span><%=strDescription%></span>
											<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%=pExt%>" readonly size="<%=len(pExt)%>" class="transparentField">
                                            <% if not pClngShowSku = 0 then %>
                                                <div class="pcSmallText"><%=strSku%></div>
                                            <% end if %>
																						
                                    <%end if%>
                                        <%'---- RAdios---%>											
                                        <%=func_DisplayBOMsg1(tempVarCat)%>
                                        <%if pcv_ShowDesc="1" then%>
                                            <div class="row">
                                                <div class="col-xs-12"><span class="configDesc"><%=pcv_prdSDesc%></span></div>
                                            </div>
																				
                                        <%end if%>	
                                    <% end if
                        end if %>
                        
                        <%IF pBTODisplayType<>1 THEN %>
                        </div>
                            <div class="col-xs-3">
                                <% if showInfoVar = "1" then %>
									
                                    <% if iBTODetLinkType=1 then%>
										<a class="" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')"><%=pcv_strBTODetTxt %></a>
                                    <%else%>
                                        <a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')">
                                            <span class="pcButtonText"><%=pcv_strBTODetTxt %></span>
                                            <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>">
                                        </a>
                                    <%end if
                                end if %>
                                <%
                                'Show Option Discounts icon
                                ProductArray = Split(ShowInfoArray,",")
                                for i = lbound(ProductArray) to (UBound(ProductArray)-1)
                                    if ProductArray(i)<>"" then
                                        MyTest=CheckOptDiscount(ProductArray(i))
                                        if MyTest=1 then%>
                                            <a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=ProductArray(i)%>')"><img alt="<%=dictLanguage.Item(Session("language")&"_viewPrd_16")%>" src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>"></a>
                                        <%end if
                                    end if
                                next
                                'End Show Option Discounts icon%>
                            </div>
                        </div>
                        <%END IF%>
                    
                        <%  
                        pcv_tmpN=pcv_tmpN+1
                        IF (pcv_tmpN<=pcv_ArrCount) THEN
                            if Clng(pcv_tmpArr(0,pcv_tmpN))<>Clng(checkVar) then
                                pcv_tmpTest=0
                            end if
                        end if
                        intOpCnt = intOpCnt + 1 
                    
                    LOOP '// DO WHILE ((pcv_tmpTest=1) AND (pcv_tmpN<=pcv_ArrCount))
      
                    IF (pcv_tmpTest=0) AND (pcv_tmpN<=pcv_ArrCount) THEN
                        pcv_tmpN=pcv_tmpN-1
                    END IF

                    Dim varTempDefaultPrice
                    varTempDefaultPrice=(defaultPrice-(defaultPrice*2))
                    If scDecSign="," then
                        varTempDefaultPrice=replace(varTempDefaultPrice,",",".")
                    End If
                    if requiredVar = "1" then
                        if pBTODisplayType<>1 then%>
                            <div <%=strCol%>><div class="col-xs-9">
                        <%end if
                        if Cdbl(varTempDefaultPrice)<0 then 
																				if pBTODisplayType=1 then
																					icount=icount+1 %>
																					<option value="0_<%=varTempDefaultPrice%>_0_0"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%> 
																					<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-(varTempDefaultPrice))%><%end if%></option>
																				 <%  StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf %>
																				<% else 
																					icount=icount+1%>
																					<input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0" onClick="calculate(this,0);" class="clearBorder"><input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                                                                    <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
                                                                                    <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*varTempDefaultPrice)%><%end if%>" readonly size="<%=len(" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*varTempDefaultPrice))%>" class="transparentField"><br>
																				<% end if %>
																			<% else if Cdbl(varTempDefaultPrice)<0 then	%>
																				<% if pBTODisplayType=1 then
																					icount=icount+1 %>
																					<option value="0_<%=varTempDefaultPrice%>_0_0"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%> 
																					<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%></option>
																				    <%  StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf %>
																				<% else
																					icount=icount+1 %>
																					<input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0" onClick="calculate(this,0);" class="clearBorder"><input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                                                                    <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
                                                                                    <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%>" readonly size="<%=len(" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice))%>" class="transparentField"><br>
																				<% end if %>
																			<% else if cdVar="0" then %>
																				<% if pBTODisplayType=1 then
																					icount=icount+1 %>
																					<option value="0_0.00_0_0" <% if cdVar="0" then Response.write "selected": strselectvalue = "" end if %>><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></option>
																				  <%  StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='' ;"  &vbcrlf %>
																				<% else
																					icount=icount+1 %>
																					<input type="radio" name="CAG<%=tempVarCat%>" value="0_0.00_0_0" <% if cdVar="0" then %> checked<% end if %> onClick="calculate(this,0);" class="clearBorder"><input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                                                                    <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
                                                                                    <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="" readonly size="1" class="transparentField"><br><% end if %>
																				<% end if
																			end if 
                    end if %>
                    <%if pBTODisplayType<>1 then%>
                    </div></div>
                    <%end if%>
                    <% end if %>
                    <% if pBTODisplayType=1 then %> 
                    </select>
                    <%=HiddenFields%>
                    <script type=text/javascript>
																			testdropdown('document.additem.CAG<%=tempVarCat%>');
																			</script>
																			<script type=text/javascript>
																			 <%=StrBackOrd %>
																			 function showAvail<%=tempVarCat%>(sel){
																			 document.getElementById("AV<%=tempVarCat%>").innerHTML = availArr<%=tempVarCat%>[sel.selectedIndex] + "&nbsp;";															 
																			 }
																			</script>
																		
																		<% if intOpCnt = 0 then %>
																			<span  id="AV<%=tempVarCat%>" ><%=func_DisplayBOMsg%>&nbsp;</span>
																		 <% else %>
																		 	<span  id="AV<%=tempVarCat%>" ><%=strselectvalue%></span>
																		 <% end if %>
																		<% end if %>
																		<%intOpCnt = intOpCnt + 1
																	
																	'// END DROP-DOWN%>
																	
                    <%IF pBTODisplayType<>1 THEN%>
                    <!--<div <%=strCol%>> -->
                    <%END IF%>
                    <input name="currentValue<%=jCnt%>" type="HIDDEN" value="<%if (pcv_FirstItem=3) then%><%=pcv_tmpDefaultValue%><%else%>0.00<%end if%>">
                    <input name="CAT<%=jCnt%>" type="HIDDEN" value="CAG<%=tempVarCat%>">
                    <input name="Discount<%=jCnt%>" type="HIDDEN" value="<%if (pcv_FirstItem=3) then%><%=pcv_tmpDefaultDiscount%><%else%><%=pcv_tmpIDiscount%><%end if%>">
																		
                    <%IF pBTODisplayType<>1 THEN%>
                    <!--</div> -->
                    <%END IF%>
                    <%IF pBTODisplayType=1 THEN
                    response.write "</div><div class=""col-xs-3"">"
																	end if %>
																	<%IF pBTODisplayType=1 THEN
                    if showInfoVar = "1" then%>
                        
						<% if iBTODetLinkType=1 then%>
							<a class="" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')"><%=pcv_strBTODetTxt %></a>
						<%else%>
							<a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')">
                                <span class="pcButtonText"><%=pcv_strBTODetTxt %></span>
							    <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>">
                            </a>
                        <%end if%>
						    
                    <% end if%>
                
                    <%
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    'Show Option Discounts icon
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    ProductArray = Split(ShowInfoArray,",")
                    MyTest=0
                    for i = lbound(ProductArray) to (UBound(ProductArray)-1)
                        if ProductArray(i)<>"" then
                            MyTest1=CheckOptDiscount(ProductArray(i))
                            if MyTest1=1 then
                                MyTest=1
                            end if
                        end if
                    next
                    if MyTest=1 then%>
					<a href="javascript:openbrowser('<%=pcv_sffolder%>OptpriceBreaks.asp?type=<%=Session("customerType")%>&amp;SIArray=<%=ShowInfoArray%>&amp;cd=<%=strCategoryDesc%>')">
					<img alt="<%=dictLanguage.Item(Session("language")&"_viewPrd_16")%>" src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>">
                    </a>
                    <%
                    end if
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    'End Show Option Discounts icon
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    response.write "</div></div>"
                    END IF
                    %>
                    </div>
                    </div>
                <% 
                end if
					
			Else   '// Else If VarMS=False then 
				
						tempCcat = pcv_tmpArr(0,pcv_tmpN)
						
						'************* IT IS NEW CAT
						If Clng(checkVarCat)<>Clng(tempCcat) Then
                            %>
                            <div class="panel panel-default">
                            <%	
						CB_CatCnt = CB_CatCnt + 1
						checkVarCat = Clng(tempCcat) %>
						<input type="hidden" name="CB_CatID<%=CB_CatCnt%>" value="<%=tempCcat%>">
						<%
						RTestStr=""
						RTestStr=RTestStr & vbcrlf & "RTest" & CB_CatCnt & "='';" & vbcrlf
						%>
								<%
								'=====================
								'LOOP THROUGH PRODUCTS
								'=====================
								If strCol <> "class='pcBTOfirstRow row'" Then
									strCol = "class='pcBTOfirstRow row'"
								Else 
									strCol = "class='pcBTOsecondRow row'"
								End If
								
								pcv_ShowDesc=pcv_tmpArr(15,pcv_tmpN)
								if IsNull(pcv_ShowDesc) or pcv_ShowDesc="" then
									pcv_ShowDesc="0"
								end if
								pClngShowItemImg=pcv_tmpArr(16,pcv_tmpN)
								if IsNull(pClngShowItemImg) or pClngShowItemImg="" then
									pClngShowItemImg="0"
								end if
								pClngSmImgWidth=pcv_tmpArr(17,pcv_tmpN)
								if IsNull(pClngSmImgWidth) or pClngSmImgWidth="" then
									pClngSmImgWidth="0"
								end if
								pClngShowSku=pcv_tmpArr(18,pcv_tmpN)
								if IsNull(pClngShowSku) or pClngShowSku="" then
									pClngShowSku="0"
								end if
								%>
								
								<div class="panel-heading"><%=pcv_tmpArr(1,pcv_tmpN)%>
									<%
									CATDesc=pcv_tmpArr(1,pcv_tmpN)
									requiredCategory=pcv_tmpArr(13,pcv_tmpN)
									

									if requiredCategory=-1 then
										ReqCAT=1
									else
										ReqCAT=0
									end if
									%>
								</div>
                                <div class="panel-body">
								<% 
									' If there are configuration instructions for this category, show them here.
									CATNotes=pcv_tmpArr(29,pcv_tmpN)
									if CATNotes<>"" then
									%>
									<div <%=strCol%>>  
										<div class="col-xs-12"><span class="catNotes"><%=CATNotes%></span></div>
									</div>
									<%
									end if
									%>
						<% PrdCnt = 0 %>
						<% 
						ShowInfoArray = ""
						showInfoVar="0"
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: SHOW CHECKBOXES WITH PRICE
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						pcv_tmpTest=1
							
						DO WHILE ((pcv_tmpTest=1) AND (pcv_tmpN<=pcv_ArrCount))
																	pcv_prdDesc=pcv_tmpArr(27,pcv_tmpN)
																	pcv_prdSDesc=pcv_tmpArr(28,pcv_tmpN)
								if IsNull(pcv_prdSDesc) or trim(pcv_prdSDesc)="" then
									pcv_prdSDesc=pcv_prdDesc
								end if
								pcv_qtyvalid=pcv_tmpArr(3,pcv_tmpN)
								if isNULL(pcv_qtyvalid) OR pcv_qtyvalid="" then
									pcv_qtyvalid="0"
								end if
								pcv_minQty=pcv_tmpArr(4,pcv_tmpN)
								if isNULL(pcv_minQty) OR pcv_minQty="" then
									pcv_minQty="1"
								end if
								if pcv_minQty<"1" then
									pcv_minQty="1"
								end if
								prdBtoBPrice = pcv_tmpArr(10,pcv_tmpN)
								prdPrice = pcv_tmpArr(9,pcv_tmpN)
								if prdBtoBPrice=0 then
									prdBtoBPrice=prdPrice
								end if
								displayQF=pcv_tmpArr(14,pcv_tmpN)
																	if pcv_tmpNewPath<>"" then
																		pcv_tmpArr(11,pcv_tmpN)=0
																	end if
																	If pcv_tmpArr(11,pcv_tmpN)=True then
																		showInfoVar="1"
																	End If
																	intTempIdProduct=pcv_tmpArr(5,pcv_tmpN)
								intTempIdCategory=pcv_tmpArr(0,pcv_tmpN)
								weight=pcv_tmpArr(6,pcv_tmpN)
								cdefault=pcv_tmpArr(12,pcv_tmpN)
								strDescription=pcv_tmpArr(7,pcv_tmpN)
								strSku=pcv_tmpArr(19,pcv_tmpN)
								strSmallImage=pcv_tmpArr(20,pcv_tmpN)							
								if strSmallImage = "" or strSmallImage = "no_image.gif" then
									strSmallImage = "hide"
								end if
								pstock=pcv_tmpArr(21,pcv_tmpN)
								pNostock=pcv_tmpArr(22,pcv_tmpN)	
								if pNostock = "" or pNoStock = null then
								 pNostock = 0
								end if
								pcv_intBackOrder = pcv_tmpArr(23,pcv_tmpN)
								pcv_intShipNDays = pcv_tmpArr(24,pcv_tmpN)
								pMinPurchase = pcv_tmpArr(25,pcv_tmpN)
																	strCategoryDesc=pcv_tmpArr(1,pcv_tmpN)
																	
																	ShowInfoArray = ShowInfoArray & intTempIdProduct& ","
																	ShowInfoArray = intTempIdProduct& "," 
								intCC_BTO_Pricing=0
								if session("customercategory")<>0 then
																		query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idcustomerCategory=" & session("customercategory") & " AND idBTOItem=" & pcv_tmpArr(5,pcv_tmpN)& " AND idBTOProduct=" & pIdProduct & ";" 
									set rsCCObj=server.CreateObject("ADODB.RecordSet")
									set rsCCObj=conntemp.execute(query)
																		
									if err.number<>0 then
										call LogErrorToDatabase()
										set rsCCObj=nothing
										call closedb()
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if

									if NOT rsCCObj.eof then
										intCC_BTO_Pricing=1
										pcCC_BTO_Price=rsCCObj("pcCC_BTO_Price")
									else
																			query="SELECT pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=" & session("customercategory") &" AND pcCC_Pricing.idProduct=" & pcv_tmpArr(5,pcv_tmpN) & ";"
										set rsCCObj=server.CreateObject("ADODB.RecordSet")
										set rsCCObj=conntemp.execute(query)
										if NOT rsCCObj.eof then
											intCC_BTO_Pricing=1
											pcCC_BTO_Price=rsCCObj("pcCC_Price")
										end if
									end if
									set rsCCObj=nothing
								end if
		
								'customer logged in as ATB customer based on retail price
								if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
									prdPrice=Cdbl(prdPrice)-(pcf_Round(Cdbl(prdPrice)*(cdbl(session("ATBPercentage"))/100),2))
								end if
	
								'customer logged in as ATB customer based on wholesale price
								if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
									prdBtoBPrice=Cdbl(prdBtoBPrice)-(pcf_Round(Cdbl(prdBtoBPrice)*(cdbl(session("ATBPercentage"))/100),2))
									prdPrice=Cdbl(prdBtoBPrice)
								end if
								
								'customer logged in as a wholesale customer
								if prdBtoBPrice>0 and session("customerType")=1 then
									prdPrice=Cdbl(prdBtoBPrice)
								end if
								'customer logged in as a customer type with price different then the online price
								if intCC_BTO_Pricing=1 then
									if (pcCC_BTO_Price<>0) OR (pcCC_BTO_Price=0 AND intCC_BTO_Pricing=1) then
										prdPrice=Cdbl(pcCC_BTO_Price)
									end if
								end if
																	
																	tmp_qty=pcv_minQty*ProQuantity

																	pcv_tmpIDiscount=0
																	call CheckDiscount(pcv_tmpArr(5,pcv_tmpN),pcv_tmpArr(12,pcv_tmpN),tmp_qty,prdPrice)
		
						PrdCnt = PrdCnt + 1
						jCnt = jCnt + 1%>
																
						<input name="MS<%=jCnt%>" type="HIDDEN" value="<%=VarMS%>">
						<input name="currentValue<%=jCnt%>" type="HIDDEN" value="<%if (cdefault<>"") and (cdefault<>0) then%><%=prdPrice%><%else%>0<%end if%>">
						<input name="Discount<%=jCnt%>" type="HIDDEN" value="<%=pcv_tmpIDiscount%>">
						<input name="CAT<%=jCnt%>" type="HIDDEN" value="CAG<%=tempCcat%>">
						<%if (cdefault<>"") and (cdefault<>0) then
							pcv_CustomizedPrice=pcv_CustomizedPrice+prdPrice
						end if%>
						<div <%=strCol%> style="vertical-align:top">  
							<div <%if (displayQF=True) then%>class="col-xs-3"<%else%>class="col-xs-2"<%end if%>>
                                <input type="hidden" name="Cat<%=intTempIdCategory%>_Prd<%=PrdCnt%>" value="<%=intTempIdProduct%>">
																		<input type="checkbox" name="CAG<%=intTempIdCategory&intTempIdProduct%>" value="<%=intTempIdProduct%>_<%=prdPrice%>_<%=weight%>_<%=prdPrice%>" onClick="javscript:document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>QF.value='<%=pcv_minQty%>'; calculate(this,0);" <%if (cdefault<>"") and (cdefault<>0) then%>checked<%end if%> class="clearBorder">
																		<%RTestStr=RTestStr & vbcrlf & "if (document.additem.CAG"& intTempIdCategory & intTempIdProduct & ".checked !=false) { RTest" & CB_CatCnt & "=" & "RTest" & CB_CatCnt & "+document.additem.CAG" & intTempIdCategory & intTempIdProduct & ".checked; }"& vbcrlf%>

																		<%if (displayQF=True) then%>
																			<input class="form-control quantity" type="text" size="2" name="CAG<%=intTempIdCategory&intTempIdProduct%>QF" value="<%if (cdefault<>"") and (cdefault<>0) then%><%=pcv_minQty%><%else%>0<%end if%>" onBlur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>)) calculate(document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>,0);">&nbsp;
																		<%else%>
																			<input type="hidden" name="CAG<%=intTempIdCategory&intTempIdProduct%>QF" value="<%if (cdefault<>"") and (cdefault<>0) then%><%=pcv_minQty%><%else%>0<%end if%>">
								<%end if%>
								<% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
									<img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
								<% end if %>
                            
							</div>
							
							<div <%if (displayQF=True) then%>class="col-xs-4"<%else%>class="col-xs-5"<%end if%>>
								<span><%=strDescription%></span>
								<% if not pClngShowSku = 0 then %>
									<div class="pcSmallText"><%=strSku%></div>
								<% end if %>
							</div>
							<div class="col-xs-2"> 
                                <div align="right">
                                    <%if pnoprices<2 then%><%=scCurSign & money(prdPrice)%><%end if%>
                                </div>
							</div>
							
							
							<div class="col-xs-3">
							<% if showInfoVar = "1" then %>
							
								<% if iBTODetLinkType=1 then %>
								    <a class="" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')"><%=pcv_strBTODetTxt %></a>
								<%else%>
								    <a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')">
                                        <span class="pcButtonText"><%=pcv_strBTODetTxt %></span>
                                        <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>" >
									</a>
							    <%end if%>
							
							<% end if %>
							<%
							'Show Option Discounts icon
							ProductArray = Split(ShowInfoArray,",")
							for i = lbound(ProductArray) to (UBound(ProductArray)-1)
								if ProductArray(i)<>"" then
									MyTest=CheckOptDiscount(ProductArray(i))
																			if MyTest=1 then%>
																			<a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=ProductArray(i)%>')"><img alt="<%=dictLanguage.Item(Session("language")&"_viewPrd_16")%>" src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>"></a>
																			<%end if
																		end if
																	next
																	'End Show Option Discounts icon%>																	
							</div>
						</div>
						<%'---- Check Boxes ---%>
						<% if func_DisplayBOMsg <> "" then %>
							<div <%=strCol%> style="vertical-align:top">
							<%=func_DisplayBOMsg1(tempVarCat)%>
							</div>
						<% end if %>
						<%if pcv_ShowDesc="1" then%>
						<div <%=strCol%> style="vertical-align:top">
							<div <%if (displayQF=True) then%>class="col-xs-3"<%else%>class="col-xs-2"<%end if%>></div>
                            <div class="col-xs-6">
								<span class="configDesc">
									<%=pcv_prdSDesc%>
								</span>
							</div>
						</div>
						<%end if%>
						
						<%pcv_tmpN=pcv_tmpN+1
						IF (pcv_tmpN<=pcv_ArrCount) THEN
						if Clng(pcv_tmpArr(0,pcv_tmpN))<>Clng(checkVarCat) then
							pcv_tmpTest=0
						end if
						END IF
						LOOP

						IF (pcv_tmpTest=0) AND (pcv_tmpN<=pcv_ArrCount) THEN
							pcv_tmpN=pcv_tmpN-1
						END IF
									
																ShowInfoArray = ""
																showInfoVar="0" 
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' END: SHOW CHECKBOXES WITH PRICE
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						%>
						<input type="hidden" name="PrdCnt<%=tempCcat%>" value="<%=PrdCnt%>">
							
							<%RTestStr=RTestStr & vbcrlf & "if (RTest" & CB_CatCnt & " == '') { alert('"& dictLanguage.Item(Session("language")&"_alert_7") & replace(CATDesc,"'","\'") & "'); return(false);}" & vbcrlf
							if ReqCAT=1 then
								ReqTestStr=ReqTestStr & RTestStr
								ReqCAT=0
							end if%>
						<% 
                        '=====================
						'End LOOP THROUGH PRODUCTS
						'===================== 
                        %>
                        </div></div>
                        <%
						end if
						'*****************************
						End If '**********************
						'*****************************
					pcv_tmpN=pcv_tmpN+1
		LOOP 'rsSSobj
	End if //'Have BTO Categories
	set rsSSobj=nothing
	'******* END BTO Categories
	'******************************************* 	
	
	response.write "<script type=text/javascript>" & VBCRlf
	response.write "function DisValue(IDPro,ProQ,ProP) {" & VBCRlf
	response.write "DisValue1=0;" & VBCRLf
	response.write "IDPro1=eval(IDPro);" & VBCRLf
	response.write "ProQ1=eval(ProQ);" & VBCRLf
	response.write "ProP1=eval(ProP);" & VBCRLf
	if TempDiscountStr<>"" then
	response.write TempDiscountStr & VBCRLf
	end if
	response.write "return(eval(DisValue1));" & VBCrlf
	response.write " } </script>" & VBCRlf
	
	response.write "<script type=text/javascript>" & VBCRlf
	response.write "function QDisValue(IDPro,ProQ,ProP) {" & VBCRlf
	response.write "DisValue1=0;" & VBCRLf
	response.write "IDPro1=eval(IDPro);" & VBCRLf
	response.write "ProQ1=ProQ.value;" & VBCRLf
	response.write "ProP1=eval(ProP);" & VBCRLf
	if TempQDStr<>"" then
	response.write TempQDStr & VBCRLf
	end if
	response.write "return(eval(DisValue1));" & VBCrlf
	response.write " } </script>" & VBCRlf
	%>
	<script type=text/javascript>
												function chkR()
												{
												<%if ReqTestStr<>"" then%>
												<%=ReqTestStr%>
												<%end if%>
												return(true);
												//return checkproqty(document.additem.quantity);
												}
	</script>
	<%Call pcs_GetDefaultBTOItemsMin%>
	<input type="hidden" name="FirstCnt" value="<%=jCnt%>">
	<input type="hidden" name="CB_CatCnt" value="<%=CB_CatCnt%>">
</div>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Configuration Table
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  javascript for calculations
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_AddChargesCalculations
%>
<script type=text/javascript>
		var scDecSign="<%=scDecSign%>";
		var scCurSign="<%=scCurSign%>";
		var tmpIDProduct="<%=pIDProduct%>";
		//Default Customized Total
		var Ctotal=0;
		//Default Item Discount Total
		var QD1=0;
		var optmsg1="<%=dictLanguage.Item(Session("language")&"_prodOpt_1")%>";
		var optmsg2="<%=dictLanguage.Item(Session("language")&"_prodOpt_2")%>";
		var showprices=<%=pnoprices%>;
	</script>

	<script type="text/javascript" src="<%=pcf_getJSPath("../includes/javascripts","calculate.js")%>"></script>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  javascript for calculations
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  IMAGES
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_AddChargesImages
%>  
	<div class="pcTable">  
	<%
	if len(pImageUrl) > 0 then
	'// A)  The image exists
	%>
		<div class="pcTableRowFull">
			<div class="pcShowMainImage">
				<%
                Dim pcv_strZoomLink, pcv_strZoomLocation  			
        
                if pcv_strUseEnhancedViews = True then
                    pcv_strZoomLink = "javascript:;"
                    pcv_strZoomLocation = "onclick=""pcf_initEnhancement(this,'"&pcf_getImagePath(pcv_tmpNewPath&"catalog",pLgimageURL)&"')"" class=""highslide"""
                else
                    pcv_strZoomLink="javascript:enlrge('"&pcf_getImagePath(pcv_tmpNewPath&"catalog",pLgimageURL)&"')"
                    pcv_strZoomLocation = ""
                end if 
                %>
                <% if len(pLgimageURL)>0 then %>
                    <a href="<%=pcv_strZoomLink%>" <%=pcv_strZoomLocation%>><img id='mainimg' name='mainimg' src='<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pImageUrl)%>' alt="<%=replace(pDescription,"""","&quot;")%>" /></a>
                <% else %>
                    <img id='mainimg' name='mainimg' src='<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pImageUrl)%>' alt="<%=replace(pDescription,"""","&quot;")%>" />
                <% end if %>
                <% if pcv_strUseEnhancedViews = True then %>
                	<div class="<%=pcv_strHighSlide_Heading%>"><%=replace(pDescription,"""","&quot;")%></div>
                <% end if %>
            </div>
		</div>
			   
		<% if len(pLgimageURL)>0 and pcv_strUseEnhancedViews = False then %>
        <div class="pcTableRowFull">
			<div class="pcShowMainImage">
                <a href="<%=pcv_strZoomLink%>" <%=pcv_strZoomLocation%>><img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("zoom"))%>" hspace="10" alt="<%=dictLanguage.Item(Session("language")&"_altTag_5")%>"></a>
                <% if pcv_strUseEnhancedViews = True then %>
                	<div class="<%=pcv_strHighSlide_Heading%>"><%=replace(pDescription,"""","&quot;")%></div>
                <% end if %>
            </div>
         </div>
        <% end if %>

		<% if pcv_strUseEnhancedViews = True then %>

			<script type=text/javascript>	
				$pc(document).ready(function() {
					hs.align = '<%=pcv_strHighSlide_Align%>';
					hs.transitions = [<%=pcv_strHighSlide_Effects%>];
					hs.outlineType = '<%=pcv_strHighSlide_Template%>';
					hs.fadeInOut = <%=pcv_strHighSlide_Fade%>;
					hs.dimmingOpacity = <%=pcv_strHighSlide_Dim%>;
					//hs.numberPosition = 'caption';
					<% if bCounter>0 then %>
						if (hs.addSlideshow) hs.addSlideshow({
							slideshowGroup: 'slides',
							interval: <%=pcv_strHighSlide_Interval%>,
							repeat: true,
							useControls: true,
							fixedControls: false,
							overlayOptions: {
								opacity: .75,
								position: 'top center',
								hideOnMouseOut: <%=pcv_strHighSlide_Hide%>
							}
						});	
					<% end if %>
				});

				function pcf_initEnhancement(ele,img) {
					if (document.getElementById('1')==null) {
						hs.expand(ele, { src: img, minWidth: <%=pcv_strHighSlide_MinWidth%>, minHeight: <%=pcv_strHighSlide_MinHeight%> }); 
					} else {
						document.getElementById('1').onclick();			
					}
				}
			</script>
                            
        <% end if %>     

	<%
	else
	'// B)  The image DOES NOT exist (show no_image.gif)
	%>		
		<div class="pcTableRowFull">
			<div class="pcShowMainImage">
				<img name='mainimg' src='<%=pcf_getImagePath(pcv_tmpNewPath & "catalog","no_image.gif")%>' alt="Product image not available" width="100" height="100">
			</div>
		</div>
	<% 
	end if
	%>
	</div>
	<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  IMAGES
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Prices
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_AddChargesPrices
%>
<div class="pcShowProductPrice" id="pcBTOhideTopPrices">
		<% if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
			pPrice=pPrice-(pcf_Round(pPrice*(cdbl(session("ATBPercentage"))/100),2))
		end if
		
		if pBtoBPrice>0 and session("customerType")=1 then
			if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
				pBtoBPrice=pBtoBPrice-(pcf_Round(pBtoBPrice*(cdbl(session("ATBPercentage"))/100),2))
			end if
			pPrice=pBtoBPrice
		End if
		
		if intCC_BTO_Pricing=1 then
			pPrice=pcCC_BTO_Price
		end if
		
			if (pPriceDefault>0) and (pnoprices<2) then
				if session("customerType")=1 then
					response.write "<b>" & dictLanguage.Item(Session("language")&"_viewPrd_15") & "</b>" & scCurSign & money(pPriceDefault)&"<br>" 
				else
					response.write "<b>" & bto_dictLanguage.Item(Session("language")&"_configurePrd_2") & "</b>" & scCurSign & money(pPriceDefault)&"<br>" 
				end if
			end if
								 
			if pnoprices=2 then%>
				<input name="GrandTotal2" type="hidden" value="<%=scCurSign%><%=money(pCMPrice+pPriceDefault)%>" >
				<input name="GrandTotal2QD" type="hidden" value="<%=scCurSign%><%=money(pCMWQD+pPriceDefault)%>" >
			<%else%>
				<b><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_3")%></b>
				<input name="GrandTotal2QD" type="text" value="<%=scCurSign%><%=money(pCMWQD+pPriceDefault)%>"  readonly size="14" class="transparentField">
				<input name="GrandTotal2" type="hidden" value="<%=scCurSign%><%=money(pCMPrice+pPriceDefault)%>" >
			<%end if%>
</div>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Prices
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Configuration Table - Reconfig
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_AddChargesTableReconfig
Dim query,rsSSObj,tmpquery
	tmpquery=""
	if (scOutOfStockPurchase="-1") AND (iBTOOutofStockPurchase="-1") then
		tmpquery=" AND ((products.stock>0) OR (products.nostock<>0) OR (products.pcProd_BackOrder<>0))"
	end if
	query="SELECT categories.idCategory, categories.categoryDesc, configSpec_Charges.multiSelect,products.pcprod_qtyvalidate,products.pcprod_minimumqty,products.idproduct, products.weight, products.description, configSpec_Charges.prdSort, configSpec_Charges.price, configSpec_Charges.Wprice, configSpec_Charges.showInfo, configSpec_Charges.cdefault, configSpec_Charges.requiredCategory, configSpec_Charges.displayQF,configSpec_Charges.pcConfCha_ShowDesc,configSpec_Charges.pcConfCha_ShowImg,configSpec_Charges.pcConfCha_ImgWidth,configSpec_Charges.pcConfCha_ShowSKU,products.sku,products.smallImageUrl,products.stock,products.noStock, products.pcProd_BackOrder, products.pcProd_ShipNDays,products.pcprod_minimumqty,configSpec_Charges.pcConfCha_UseRadio,products.details,products.sDesc,configSpec_Charges.Notes FROM categories INNER JOIN (products INNER JOIN configSpec_Charges ON (products.idproduct=configSpec_Charges.configProduct AND products.active<>0 AND products.removed=0" & tmpquery & ")) ON categories.idCategory = configSpec_Charges.configProductCategory WHERE configSpec_Charges.specProduct="&pIdProduct&" ORDER BY configSpec_Charges.catSort, categories.idCategory, configSpec_Charges.prdSort,products.description;"
	tmpquery=""
	set rsSSObj=conntemp.execute(query)
	displayQF="0"
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsSSObj=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if%>
	<div class="">
						<% CB_CatCnt = 0
						jcnt=0

						'*******************************************
						'******* START BTO Categories


						IF NOT rsSSobj.eof then  
							Dim strCol
							strCol = "class='pcBTOsecondRow row'"
							checkVar=0
							checkVarCat=0

							pcv_tmpArr=rsSSobj.GetRows()
							pcv_ArrCount=ubound(pcv_tmpArr,2)
							set rsSSobj=nothing

							'*********** LOOP CATs
						
							pcv_tmpN=0

							DO WHILE (pcv_tmpN<=pcv_ArrCount)

								tempVarCat = pcv_tmpArr(0,pcv_tmpN)
								VarMS=pcv_tmpArr(2,pcv_tmpN)
														
								If VarMS=False then 
									dim defaultPrice
									defaultPrice=Cdbl(0)
									dim cdVar
									cdVar="0"
									
									'**** IT IS NEW CAT
									If Clng(tempVarCat) <> Clng(checkVar) then
                                        %>
                                        <div class="panel panel-default">
                                        <%	
										checkVar = tempVarCat
										strCategoryDesc=pcv_tmpArr(1,pcv_tmpN)
										
										pcv_ShowDesc="0"
										pClngShowItemImg="0"
										pClngSmImgWidth="0"
										pClngShowSku="0"
										
										if pcv_tmpArr(15,pcv_tmpN)="1" then
											pcv_ShowDesc="1"
										end if
										if pcv_tmpArr(16,pcv_tmpN)="1" then
											pClngShowItemImg="1"
										end if
										if pcv_tmpArr(17,pcv_tmpN)>"0" then
											pClngSmImgWidth=pcv_tmpArr(17,pcv_tmpN)
										end if
										if pcv_tmpArr(18,pcv_tmpN)="1" then
											pClngShowSku="1"
										end if
										
										'***** GET DEFAULT PRICE OF THE CAT
										query="SELECT configSpec_Charges.configProduct,configSpec_Charges.price, configSpec_Charges.Wprice, configSpec_Charges.cdefault FROM configSpec_Charges WHERE configSpec_Charges.configProductCategory="&tempVarCat&" AND configSpec_Charges.specProduct="&pIdProduct&" AND configSpec_Charges.cdefault<>0;"
										set rsTempObj=conntemp.execute(query)
										if err.number<>0 then
											call LogErrorToDatabase()
											set rsTempObj=nothing
											call closedb()
											response.redirect "techErr.asp?err="&pcStrCustRefID
										end if
	
                                        If NOT rsTempObj.eof then
                                            cdVar="1"
                                            tmpintPrd=rsTempObj("configProduct")
                                            dblprice=Cdbl(rsTempObj("price"))
                                            dblWprice=Cdbl(rsTempObj("Wprice"))
                                            
                                            if dblWprice=0 then
                                                dblWprice=dblprice
                                            end if
                                            
                                            query="SELECT products.pcprod_minimumqty FROM Products WHERE idproduct=" & tmpintPrd & ";"
                                            set rsQ=connTemp.execute(query)
                                            if not rsQ.eof then
                                                pcv_minqty=rsQ("pcprod_minimumqty")
                                                if IsNull(pcv_minqty) or pcv_minqty="" then
                                                    pcv_minqty=1
                                                end if
                                                if pcv_minqty="0" then
                                                    pcv_minqty=1
                                                end if
                                            else
                                                pcv_minqty=1
                                            end if
                                            set rsQ=nothing
                                                                                        
                                            intCC_BTO_Pricing=0
                                            if session("customercategory")<>0 then
                                                query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idcustomerCategory=" & session("customercategory") & " AND idBTOItem=" & tmpintPrd & " AND idBTOProduct=" & pIdProduct & ";" 
                                                set rsCCObj=server.CreateObject("ADODB.RecordSet")
                                                set rsCCObj=conntemp.execute(query)
                                                
                                                if err.number<>0 then
                                                    call LogErrorToDatabase()
                                                    set rsCCObj=nothing
                                                    call closedb()
                                                    response.redirect "techErr.asp?err="&pcStrCustRefID
                                                end if
                                                                                    
                                                if NOT rsCCObj.eof then
                                                    intCC_BTO_Pricing=1
                                                    pcCC_BTO_Price=rsCCObj("pcCC_BTO_Price")
                                                else
                                                    query="SELECT pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=" & session("customercategory") &" AND pcCC_Pricing.idProduct=" & tmpintPrd & ";"
                                                    set rsCCObj=server.CreateObject("ADODB.RecordSet")
                                                    set rsCCObj=conntemp.execute(query)
                                                    if NOT rsCCObj.eof then
                                                        intCC_BTO_Pricing=1
                                                        pcCC_BTO_Price=rsCCObj("pcCC_Price")
                                                    end if
                                                end if
                                                set rsCCObj=nothing
                                            end if
                                                    
                                            'customer logged in as ATB customer based on retail price
                                            if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
                                                dblprice=Cdbl(dblprice)-(pcf_Round(Cdbl(dblprice)*(cdbl(session("ATBPercentage"))/100),2))
                                            end if
                                            defaultPrice= Cdbl(dblprice)
                                            
                                            'customer logged in as ATB customer based on wholesale price
                                            if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
                                                dblWprice=Cdbl(dblWprice)-(pcf_Round(Cdbl(dblWprice)*(cdbl(session("ATBPercentage"))/100),2))
                                                defaultPrice=Cdbl(dblWprice)
                                            end if
                                            
                                            'customer logged in as a wholesale customer
                                            if dblWprice>0 and session("customerType")=1 then
                                                defaultPrice=Cdbl(dblWprice)
                                            end if
                                            
                                            'customer logged in as a customer type with price different then the online price
                                            if intCC_BTO_Pricing=1 then
                                                if (pcCC_BTO_Price<>0) OR (pcCC_BTO_Price=0 AND intCC_BTO_Pricing=1) then
                                                    defaultPrice=Cdbl(pcCC_BTO_Price)
                                                end if
                                            end if
                                            
                                            defaultPrice=defaultPrice*pcv_minqty

                                        End if
                                        Set rsTempObj=nothing													
										'***** END OF GET DEFAULT PRICE OF THE CAT

                                        jcnt=jCnt+1
                                        If strCol <> "class='pcBTOfirstRow row'" Then
                                            strCol = "class='pcBTOfirstRow row'"
                                        Else 
                                            strCol = "class='pcBTOsecondRow row'"
                                        End If 
                                        %>
                                
                                        <div class="panel-heading"><%=pcv_tmpArr(1,pcv_tmpN)%>
                                        </div>
                                        <div class="panel-body">
                                        <%
                                        ' If there are configuration instructions for this category, show them here.
										CATNotes=pcv_tmpArr(29,pcv_tmpN)
										if CATNotes <> "" then
										%>
                                        <div <%=strCol%>>
                                            <div class="col-xs-12"><span class="catNotes"><%=CATNotes%></span></div>
                                        </div>
                                        <% end if
                                        
                                        pBTODisplayType=pcv_tmpArr(26,pcv_tmpN)
                                        if IsNull(pBTODisplayType) or pBTODisplayType="" then
                                            pBTODisplayType=1
                                        end if
                                                
                                        displayQF=pcv_tmpArr(14,pcv_tmpN)
                                        requiredCategory=pcv_tmpArr(13,pcv_tmpN)
                                        if pcv_tmpNewPath<>"" then
                                            pcv_tmpArr(11,pcv_tmpN)=0
                                        end if
										showInfo=pcv_tmpArr(11,pcv_tmpN)%>
					
										<%
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
										' START: Show Dropdown
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
										%>
						<% 'check to see what option was checked for this category
						dim tempPrd
						tempPrd=Clng(0)
						tempQ=clng(0)
						dim i
						for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
							if Clng(ArrCategory(i))=Clng(tempVarCat) then
								tempPrd = ArrProduct(i)
								'tempQ=ArrQuantity(i)
								tempQ=1
							end if
						next%>
																
                                        <% 
                                        '// START DROP-DOWN
                                        if pBTODisplayType=1 then
											pcv_ListForGenInfo=pcv_ListForGenInfo & "GenDropInfo(document.additem.CAG" & tempVarCat & ");" & vbcrlf
                                            response.write "<div " & strCol & ">"
											if pcv_tmpArr(14,pcv_tmpN)=true then %>
												<div class='col-xs-2'>
												<input class="form-control quantity" type="text" size="2" name="CAG<%=tempVarCat%>QF" value="<%=tempQ%>" onblur="javascript:testdropqty(this,'document.additem.CAG<%=tempVarCat%>');">&nbsp;
                                                </div>
                                                <div class='col-xs-7'>
											<%else%>
												<div class='col-xs-9'>
												<input type="hidden" name="CAG<%=tempVarCat%>QF" value="<%=tempQ%>">
                                            <%end if%>
                                            <select class="form-control" name="CAG<%=tempVarCat%>" onChange="testdropdown('document.additem.CAG<%=tempVarCat%>'); calculate(this,0); showAvail<%=tempVarCat%>(this);">
                                            <% HiddenFields=""
                                        else %>
											<input type="hidden" name="CAG<%=tempVarCat%>QF" value="<%=tempQ%>">
											<%pcv_ListForGenInfo=pcv_ListForGenInfo & "GenRadioExtInfo(document.additem.CAG" & tempVarCat & ");" & vbcrlf
											if Clng(requiredCategory)<>0 then
                                                RTestStr="totalradio=document.additem.CAG" & tempVarCat & ".length;" & vbcrlf
                                                RTestStr=RTestStr & "RadioChecked=0;" & vbcrlf
                                                RTestStr=RTestStr & "if (totalradio>0) {" & vbcrlf
                                                RTestStr=RTestStr & "for (var mk=0;mk<totalradio;mk++) {" & vbcrlf
                                                RTestStr=RTestStr & "if (document.additem.CAG" & tempVarCat & "[mk].checked==true) { RadioChecked=1; break; } }" & vbcrlf
                                                RTestStr=RTestStr & "} else { if (document.additem.CAG" & tempVarCat & ".checked==true) RadioChecked=1;}" & vbcrlf
                                                RTestStr=RTestStr & "if (RadioChecked==0) {alert('"& dictLanguage.Item(Session("language")&"_alert_7") & replace(pcv_tmpArr(1,pcv_tmpN),"'","\'") & "'); return(false);}" & vbcrlf
                                                ReqTestStr=ReqTestStr & RTestStr
                                            end if%>
                                        <% end if '// if pBTODisplayType=1 then %>
                                
                                        <% 
										Dim requiredVar, showInfoVar, ShowInfoArray, SelectedVar
                                        requiredVar="0"
                                        showInfoVar="0"
                                        ShowInfoArray = ""
										SelectedVar = "0"
						
										if pcv_tmpArr(13,pcv_tmpN)=False then
											requiredVar = "1"
										end if
										if pcv_tmpNewPath<>"" then
											pcv_tmpArr(11,pcv_tmpN)=0
										end if

										if pcv_tmpArr(11,pcv_tmpN)=True then
                                            showInfoVar = "1"
                                        end if
                                        icount=0

                                        pcv_tmpIDiscount=0
						pcv_tmpCustomizedPrice=0
									
						pcv_tmpTest=1
						intOpCnt = 0
						StrBackOrd = "var availArr"&tempVarCat &" = new Array();" &vbcrlf
						strselectvalue = "" 
						DO WHILE ((pcv_tmpTest=1) AND (pcv_tmpN<=pcv_ArrCount))
							if pBTODisplayType<>1 then
							ShowInfoArray = ""%>
								<div <%=strCol%>>
							<%end if
							icount=icount+1
							intTempIdProduct=pcv_tmpArr(5,pcv_tmpN)
							intTempIdCategory=pcv_tmpArr(0,pcv_tmpN)
							
							
							pcv_qtyvalid=pcv_tmpArr(3,pcv_tmpN)
							if isNULL(pcv_qtyvalid) OR pcv_qtyvalid="" then
								pcv_qtyvalid="0"
							end if
							pcv_minQty=pcv_tmpArr(4,pcv_tmpN)
							if isNULL(pcv_minQty) OR pcv_minQty="" then
								pcv_minQty="1"
							end if
							if pcv_minQty<"1" then
								pcv_minQty="1"
							end if
							displayQF=pcv_tmpArr(14,pcv_tmpN)
							prdBtoBPrice = Cdbl(pcv_tmpArr(10,pcv_tmpN))
							prdPrice = Cdbl(pcv_tmpArr(9,pcv_tmpN))
							if prdBtoBPrice=0 then
								prdBtoBPrice=prdPrice
							end if
							strDescription=pcv_tmpArr(7,pcv_tmpN)
							strSku=pcv_tmpArr(19,pcv_tmpN)
							strSmallImage=pcv_tmpArr(20,pcv_tmpN)							
							if strSmallImage = "" or strSmallImage = "no_image.gif" then
								strSmallImage = "hide"
							end if
							pstock=pcv_tmpArr(21,pcv_tmpN)
							pNostock=pcv_tmpArr(22,pcv_tmpN)	
							if pNostock = "" or pNoStock = null then
							 pNostock = 0
							end if						
							pcv_intBackOrder = pcv_tmpArr(23,pcv_tmpN)							
							pcv_intShipNDays = pcv_tmpArr(24,pcv_tmpN)
							pMinPurchase = pcv_tmpArr(25,pcv_tmpN)
							pcv_prdDesc=pcv_tmpArr(27,pcv_tmpN)
							pcv_prdSDesc=pcv_tmpArr(28,pcv_tmpN)
							if IsNull(pcv_prdSDesc) or trim(pcv_prdSDesc)="" then
								pcv_prdSDesc=pcv_prdDesc
							end if
							intCC_BTO_Pricing=0																
							if session("customercategory")<>0 then
								query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idcustomerCategory=" & session("customercategory") & " AND idBTOItem=" & intTempIdProduct& " AND idBTOProduct=" & pIdProduct & ";" 
								set rsCCObj=server.CreateObject("ADODB.RecordSet")
								set rsCCObj=conntemp.execute(query)
																																
								if err.number<>0 then
									call LogErrorToDatabase()
									set rsCCObj=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
																				
								if NOT rsCCObj.eof then
									intCC_BTO_Pricing=1
									pcCC_BTO_Price=rsCCObj("pcCC_BTO_Price")
								else
									query="SELECT pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=" & session("customercategory") &" AND pcCC_Pricing.idProduct=" & intTempIdProduct & ";"
									set rsCCObj=server.CreateObject("ADODB.RecordSet")
									set rsCCObj=conntemp.execute(query)
									if NOT rsCCObj.eof then
										intCC_BTO_Pricing=1
										pcCC_BTO_Price=rsCCObj("pcCC_Price")
									end if
								end if
								SET rsCCObj=nothing
							end if
													
							'customer logged in as ATB customer based on retail price
							if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
								prdPrice=Cdbl(prdPrice)-(pcf_Round(Cdbl(prdPrice)*(cdbl(session("ATBPercentage"))/100),2))
							end if
												
							'customer logged in as ATB customer based on wholesale price
							if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
								prdBtoBPrice=Cdbl(prdBtoBPrice)-(pcf_Round(Cdbl(prdBtoBPrice)*(cdbl(session("ATBPercentage"))/100),2))
								prdPrice=Cdbl(prdBtoBPrice)
							end if
							
							'customer logged in as a wholesale customer
							if prdBtoBPrice>0 and session("customerType")=1 then
								prdPrice=Cdbl(prdBtoBPrice)
							end if

							'customer logged in as a customer type with price different then the online price
							if intCC_BTO_Pricing=1 then
								if (pcCC_BTO_Price<>0) OR (pcCC_BTO_Price=0 AND intCC_BTO_Pricing=1) then
									prdPrice=Cdbl(pcCC_BTO_Price)
								end if
							end if
							
							
							if Clng(tempPrd) = Clng(intTempIdProduct) then
								tmp_selected=true
							else
								tmp_selected=false
							end if
							
							if tmp_selected then
								tmp_qty=tempQ*ProQuantity
							else
								tmp_qty=pcv_minQty*ProQuantity
							end if
							
							call CheckDiscount(intTempIdProduct,tmp_selected,tmp_qty,prdPrice)
							if tmp_selected then
								'pcv_tmpCustomizedPrice=Cdbl(prdPrice)*tempQ-Cdbl(defaultPrice)
								pcv_tmpCustomizedPrice=Cdbl(prdPrice)*tempQ
								pExt = " "
							end if
							
							'DEFAULT ITEM
							if pcv_tmpArr(12,pcv_tmpN)=true then
								ShowInfoArray = ShowInfoArray & intTempIdProduct& ","
								prdPrice1=prdPrice
								'ALSO SELECTED ITEM
								If Clng(tempPrd) = Clng(intTempIdProduct) then 
									if pBTODisplayType=1 then %>
                                    <div class="col-xs-9">
										<option value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>" selected><%=strDescription & pExt%></option>
										<% HiddenFields=HiddenFields & "<input type=hidden name=""CAG" & tempVarCat & intTempIdProduct & "HF"" value=""" & pcv_qtyValid & "_" & pcv_minQty & """>" %>
                                                        <% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
                                                        strselectvalue = func_DisplayBOMsg
                                                        %>
                                                <%else %>
                                                    <% if (displayQF=True) then %>
                                                        <div class="col-xs-3">
															<input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>" checked onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; calculate(this,0);" class="clearBorder">
															<input class="form-control quantity" type="text" size="2" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="<%=tempQ%>" onblur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>)) calculate(document.additem.CAG<%=tempVarCat%>,2);">&nbsp;
                                                            <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                                                <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                                            <% end if %>
                                                        </div>
                                                        <div class="col-xs-6">
                                                            <span><%=strDescription%></span>
															<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%=pExt%>" readonly class="transparentField" size="<%=len(pExt)%>">
                                                            <% if not pClngShowSku = 0 then %>
                                                                <div class="pcSmallText"><%=strSku%></div>
                                                            <% end if %>
                                                    <%else%>
                                                        <div class="col-xs-2">
															<input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>" checked onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; calculate(this,0);" class="clearBorder"><input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="<%=tempQ%>">
                                                            <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                                                <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                                            <% end if %>
                                                        </div>
                                                        <div class="col-xs-7">
                                                            <span><%=strDescription%></span>
															<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%=pExt%>" readonly class="transparentField">
                                                            <% if not pClngShowSku = 0 then %>
                                                                <div class="pcSmallText"><%=strSku%></div>
                                                            <% end if %>
                                                    <%end if%>
                                                    <%=func_DisplayBOMsg1(tempVarCat)%>
                                                    <%if pcv_ShowDesc="1" then%>
                                                        <div class="row">
                                                            <div class="col-xs-12"><span class="configDesc"><%=pcv_prdSDesc%></span></div>
                                                        </div>
                                                    <%end if%>
									<% end if %>
								<%'DEFAULT BUT NOT SELECTED
								else %>
									<% if pBTODisplayType=1 then %>
									    <div class="col-xs-9">
                                            <option value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>"><%=strDescription%></option>
											<%
											StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
											HiddenFields=HiddenFields & "<input type=hidden name=""CAG" & tempVarCat & intTempIdProduct & "HF"" value=""" & pcv_qtyValid & "_" & pcv_minQty & """>" & vbcrlf
									else %>
											<%if (displayQF=True) then%>
                                                <div class="col-xs-3">
											        <input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>" onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; calculate(this,0);" class="clearBorder">&nbsp;
											        <input type="text" size="2" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="0" onblur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>)) calculate(document.additem.CAG<%=tempVarCat%>,2);">&nbsp;
											        <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
												        <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
											        <% end if %>
                                                </div>
                                                <div class="col-xs-6">
											        <span><%=strDescription%></span>
											        <% if not pClngShowSku = 0 then %>
												        <div class="pcSmallText"><%=strSku%></div>
											        <% end if %>
											        <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="" readonly class="transparentField"><br>
											<%else%>
                                                <div class="col-xs-2">
											        <input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>" onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; calculate(this,0);" class="clearBorder">
											        <input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="0">
											        <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
												        <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
											        <% end if %>
                                                </div>
                                                <div class="col-xs-7">   
											        <span><%=strDescription%></span>
											        <% if not pClngShowSku = 0 then %>
												        <span class="pcSmallText"><%=strSku%></span>
											        <% end if %>
											        <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="" readonly class="transparentField"><br>
											<%end if%>
											<%=func_DisplayBOMsg1(tempVarCat)%>
											<%if pcv_ShowDesc="1" then%>
												<div class="row">
													<div class="col-xs-12"><span class="configDesc"><%=pcv_prdSDesc%></span></div>
												
												</div>
											<%end if%>
											<% end if%>
								<% end if %>
							<%'NOT DEFAULT ITEM
							else
								ShowInfoArray = ShowInfoArray & intTempIdProduct& "," 
								dim pExt
								pExt = ""
								prdPrice1=prdPrice
                                                If prdPrice=Cdbl(defaultPrice) then
                                                    prdPrice=0
                                                Else
                                                    prdPrice=prdPrice-Cdbl(defaultPrice)
                                                End if
                                                
                                                tmp_price=prdPrice+(tmp_qty-1)*prdPrice1-pcv_tmpIDiscount1
                                                
                                                if pnoprices<2 then
                                                    If tmp_price>0 then
                                                        pExt = " - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(tmp_price)
                                                    Else
                                                        If tmp_price<0 then
                                                            pExt = " - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*tmp_price)
                                                        End If
                                                    End If
                                                End If
                                                
                                                If scDecSign="," then
                                                    prdPrice=replace(prdPrice,",",".")
                                                    prdPrice1=replace(prdPrice1,",",".")
                                                End If
								If Clng(tempPrd) = Clng(intTempIdProduct) then
									if pBTODisplayType=1 then %>
                                         <div class="col-xs-9">
												<option value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>" selected><%=strDescription%></option>
												<% 
														StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
														strselectvalue = func_DisplayBOMsg
														HiddenFields=HiddenFields & "<input type=hidden name=""CAG" & tempVarCat & intTempIdProduct & "HF"" value=""" & pcv_qtyValid & "_" & pcv_minQty & """>" & vbcrlf
                                                else %>
                                                    <% if (displayQF=True) then%>
                                                        <div class="col-xs-3">
															<input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>" onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; calculate(this,0);" checked class="clearBorder">
															<input class="form-control quantity" type="text" size="2" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="<%=tempQ%>" onblur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>)) calculate(document.additem.CAG<%=tempVarCat%>,2);">&nbsp;
                                                            <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                                                <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                                            <% end if %>
                                                        </div>
                                                        <div class="col-xs-6">
                                                            <span><%=strDescription%></span>
															<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="" readonly class="transparentField" size="<%=len(pExt)%>">
                                                            <% if not pClngShowSku = 0 then %>
                                                                <div class="pcSmallText"><%=strSku%></div>
                                                            <% end if %>
                                                    <%else%>
                                                        <div class="col-xs-2">
															<input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>" onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; calculate(this,0);" checked class="clearBorder"><input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="<%=tempQ%>">
                                                            <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                                                <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                                            <% end if %>

                                                        </div>
                                                        <div class="col-xs-7">


															<span><%=strDescription%></span>
															<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="" readonly class="transparentField" size="<%=len(pExt)%>"><br>
											
                                                            <% if not pClngShowSku = 0 then %>
                                                                <div class="pcSmallText"><%=strSku%></div>
                                                            <% end if %>
                                                    <%end if%>
													
                                                    <%=func_DisplayBOMsg1(tempVarCat)%>
                                                    <%if pcv_ShowDesc="1" then%>
                                                        <div class="row">
                                                            <div class="col-xs-12"><span class="configDesc"><%=pcv_prdSDesc%></span></div>
                                                        </div>
                                                    <%end if%>
                                                <% end if
								else
								' Other
									    if pBTODisplayType=1 then %>
											    <div class="col-xs-9">
												    <option value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>"><%=strDescription&pExt%></option>
											        <%
											        StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
											        HiddenFields=HiddenFields & "<input type=hidden name=""CAG" & tempVarCat & intTempIdProduct & "HF"" value=""" & pcv_qtyValid & "_" & pcv_minQty & """>" & vbcrlf

											else %>
											<%
											if (displayQF=True) then%>
											    <div class="col-xs-3">
                                                    <input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>" onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; calculate(this,0);" class="clearBorder">
											        <input type="text" size="2" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="0" onblur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>)) calculate(document.additem.CAG<%=tempVarCat%>,2);">&nbsp;
											        <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
												        <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
											        <% end if %>
                                                </div>
                                                <div class="col-xs-6">
											        <span><%=strDescription%></span>
											        <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%=pExt%>" readonly class="transparentField" size="<%=len(pExt)%>"><br>
											        <% if not pClngShowSku = 0 then %>
												        <div class="pcSmallText"><%=strSku%></div>
											        <% end if %>
											<%else%>
											    <div class="col-xs-2">
											        <input type="radio" name="CAG<%=tempVarCat%>" value="<%=intTempIdProduct%>_<%=prdPrice1%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice1%>" onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; calculate(this,0);" class="clearBorder"><input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="0">
											        <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
												        <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
											        <% end if %>
                                                </div>
                                                <div class="col-xs-7">
											        <span><%=strDescription%></span>
											        <% if not pClngShowSku = 0 then %>
												        <div class="pcSmallText"><%=strSku%></div>
											        <% end if %>
											        <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%=pExt%>" readonly class="transparentField" size="<%=len(pExt)%>"><br>
											<%end if%>
                                            
													<%'---- RAdios---%>											
                                                    <%=func_DisplayBOMsg1(tempVarCat)%>
                                                    <%if pcv_ShowDesc="1" then%>
                                                        <div class="row">
                                                            <div class="col-xs-12"><span class="configDesc"><%=pcv_prdSDesc%></span></div>
                                                        </div>
                                                    <%end if%>
                                                    
                                                <% end if
                                            end if
                                            end if
                                            
                                            IF pBTODisplayType<>1 THEN%>
                                            </div><div class="col-xs-3">
                                                <% if showInfoVar="1" then %>
													
													<% if iBTODetLinkType=1 then%>	
														<a class="" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')"><%=pcv_strBTODetTxt %></a>
													<%else%>
														<a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')">
                                                            <span class="pcButtonText"><%=pcv_strBTODetTxt %></span>
                                                            <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>">
                                                        </a>
                                                    <%end if
                                                end if %>
                                                <%
                                                'Show Option Discounts icon
                                                ProductArray = Split(ShowInfoArray,",")
                                                for i = lbound(ProductArray) to (UBound(ProductArray)-1)
                                                    if ProductArray(i)<>"" then
                                                        MyTest=CheckOptDiscount(ProductArray(i))
                                                        if MyTest=1 then%>
                                                            <a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=ProductArray(i)%>')"><img alt="<%=dictLanguage.Item(Session("language")&"_viewPrd_16")%>" src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>"></a>
                                                        <%end if
                                                    end if
                                                next
                                                'End Show Option Discounts icon%>
                                            </div>
                                            </div>
                                            <%END IF

                                            pcv_tmpN=pcv_tmpN+1
                                            IF (pcv_tmpN<=pcv_ArrCount) THEN
                                                if Clng(pcv_tmpArr(0,pcv_tmpN))<>Clng(checkVar) then
                                                    pcv_tmpTest=0
                                                end if
                                            end if
                                        intOpCnt = intOpCnt + 1
                                        
                                        LOOP '// DO WHILE ((pcv_tmpTest=1) AND (pcv_tmpN<=pcv_ArrCount))
								
						IF (pcv_tmpTest=0) AND (pcv_tmpN<=pcv_ArrCount) THEN
							pcv_tmpN=pcv_tmpN-1
						END IF

						Dim varTempDefaultPrice
						varTempDefaultPrice=(defaultPrice-(defaultPrice*2))
						If scDecSign="," then
							varTempDefaultPrice=replace(varTempDefaultPrice,",",".")
						End If
						if requiredVar = "1" then
							if pBTODisplayType<>1 then%>
								<div <%=strCol%>><div class="col-xs-9">
							<%end if
							if Cdbl(varTempDefaultPrice)<0 then 
								if tempPrd=0 then %>
									<%pcv_tmpCustomizedPrice=varTempDefaultPrice
									if pBTODisplayType=1 then
									icount=icount+1%>
										<option value="0_<%=varTempDefaultPrice%>_0_0" selected><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%><%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(varTempDefaultPrice)%><%end if%></option>
									    <% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
										   strselectvalue = func_DisplayBOMsg
										%>
									<% else
									icount=icount+1 %>
										<input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0" checked onClick="calculate(this,0);" class="clearBorder">
										
										<input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
										<span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
										<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*varTempDefaultPrice)%><%end if%>" readonly class="transparentField" size="<%=len(" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*varTempDefaultPrice))%>">
									
									<% end if %>
								<% else %>
									<% if pBTODisplayType=1 then
									icount=icount+1%>
										<option value="0_<%=varTempDefaultPrice%>_0_0"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%><%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(varTempDefaultPrice)%><%end if%></option>
									   <% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
									 else
									icount=icount+1 %>
										<input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0"  onClick="calculate(this,0);" class="clearBorder">
										
										<input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
										<span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
										<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*varTempDefaultPrice)%><%end if%>" readonly class="transparentField" size="<%=len(" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*varTempDefaultPrice))%>">
									
									<% end if %>
								<% end if %>
							<% else if Cdbl(varTempDefaultPrice)<0 then	
							if tempPrd=0 then %>
								<%pcv_tmpCustomizedPrice=varTempDefaultPrice
								if pBTODisplayType=1 then
								icount=icount+1%>
									<option value="0_<%=varTempDefaultPrice%>_0_0" selected><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%><%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%></option>
								   <% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
									  strselectvalue = func_DisplayBOMsg
								else
								icount=icount+1 %>
									<input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0" checked onClick="calculate(this,0);" class="clearBorder">
									<input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
									<span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
									<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%>" readonly class="transparentField" size="<%=len(" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice))%>">
								
								<% end if %>
							<% else %>
								<% if pBTODisplayType=1 then
								icount=icount+1%>
									<option value="0_<%=varTempDefaultPrice%>_0_0"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%><%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%></option>
									<% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='' ;"  &vbcrlf %>
								<% else
								icount=icount+1 %>
									<input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0" onClick="calculate(this,0);" class="clearBorder">
									
									<input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
									<span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
									<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%>" readonly class="transparentField" size="<%=len(" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice))%>">
								
								<% end if %>

						<% end if %>
						<% else if cdVar="0" then 
							if tempPrd=0 then %>
								<% if pBTODisplayType=1 then
								icount=icount+1 %>
									<option value="0_0_0_0" selected><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></option>
									  <%  StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='' ;"  &vbcrlf 
									  %>
								<% else
								icount=icount+1 %>
								<input type="radio" name="CAG<%=tempVarCat%>" value="0_0_0_0" checked onClick="calculate(this,0);" class="clearBorder">
								
								<input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
								<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="" readonly class="transparentField" size="1">
								
								<% end if %>
							<% else %>
								<% if pBTODisplayType=1 then
								icount=icount+1 %>
									<option value="0_0_0_0"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></option>
									      <%StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='' ;"  &vbcrlf 
								   else
										icount=icount+1 %>
										<input type="radio" name="CAG<%=tempVarCat%>" value="0_0_0_0" onClick="calculate(this,0);" class="clearBorder">
										
										<input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                        <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
										<input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="" readonly class="transparentField" size="1">
									
										<% end if %>
						<% end if %>
						<% end if %>
						<% end if
						end if
                                            if pBTODisplayType<>1 then%>
                                            </div></div>
                                            <%end if
                                        end if %>
                                        
                                        <% if pBTODisplayType=1 then %> 
                                            </select>
                                            <%=HiddenFields%>
                                            <script type=text/javascript>
                                             <%=StrBackOrd %>
                                             function showAvail<%=tempVarCat%>(sel){
                                             document.getElementById("AV<%=tempVarCat%>").innerHTML = availArr<%=tempVarCat%>[sel.selectedIndex] + "";															 
                                             }
                                            </script>
					
                                            <span  id="AV<%=tempVarCat%>" ><%=strselectvalue%></span>
                                         <% end if %>
						
										<%pcv_CustomizedPrice=pcv_CustomizedPrice+pcv_tmpCustomizedPrice%>
                                        <%intOpCnt = intOpCnt + 1

                                        '// END DROP-DOWN
                                        IF pBTODisplayType<>1 THEN%>
                                             <!--<div <%=strCol%>> --> 
                                        <%END IF%>
										<input name="currentValue<%=jCnt%>" type="HIDDEN" value="<%=pcv_tmpCustomizedPrice%>">
										<input name="Discount<%=jCnt%>" type="HIDDEN" value="<%=pcv_tmpIDiscount%>">
                                        <input name="CAT<%=jCnt%>" type="HIDDEN" value="CAG<%=tempVarCat%>">
                                        <%IF pBTODisplayType<>1 THEN%>
                                            <!--</div>  -->
                                        <%END IF%>
                                        <% if pBTODisplayType=1 then 
                                            response.write "</div><div class=""col-xs-3"">"
                                        end if %>
                                        <%IF pBTODisplayType=1 THEN
                                        if showInfoVar="1" then%>
											
											<% if iBTODetLinkType=1 then%>
                                                <a class="" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')"><%=pcv_strBTODetTxt %></a>
											<%else%>
                                                <a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')">
                                                    <span class="pcButtonText"><%=pcv_strBTODetTxt %></span>
												    <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>">
                                                </a>
											<%end if%>
                                            
                                        <%end if%>
										<%
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
										'Show Option Discounts icon
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                        ProductArray = Split(ShowInfoArray,",")
                                        MyTest=0
                                        for i = lbound(ProductArray) to (UBound(ProductArray)-1)
                                            if ProductArray(i)<>"" then
                                                MyTest1=CheckOptDiscount(ProductArray(i))
                                                if MyTest1=1 then
                                                    MyTest=1
                                                end if
                                            end if
                                        next
                                        if MyTest=1 then%>
											<a href="javascript:openbrowser('<%=pcv_sffolder%>OptpriceBreaks.asp?type=<%=Session("customerType")%>&SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')">
												<img alt="<%=dictLanguage.Item(Session("language")&"_viewPrd_16")%>" src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>">
											</a>
                                        <%end if
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                        'End Show Option Discounts icon
										'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                        response.write "</div></div>"
                                        END IF%>
                                            
									</div></div>
                                    <% 
                                    End If '// If Clng(tempVarCat) <> Clng(checkVar) then
									
                                Else

                                    tempCcat = pcv_tmpArr(0,pcv_tmpN)
															
                                    '************* IT IS NEW CAT
                                    If Clng(checkVarCat)<>Clng(tempCcat) Then
                                        %>
                                        <div class="panel panel-default">
                                        <%	
                                        CB_CatCnt = CB_CatCnt + 1
                                        checkVarCat = Clng(tempCcat) %>
                                        <input type="hidden" name="CB_CatID<%=CB_CatCnt%>" value="<%=tempCcat%>">
                                        <%
										RTestStr=""
                                        RTestStr=RTestStr & vbcrlf & "RTest" & CB_CatCnt & "='';" & vbcrlf
										%>
                                        <%
                                        '=====================
                                        'LOOP THROUGH PRODUCTS
                                        '=====================
                                        pcv_ShowDesc=pcv_tmpArr(15,pcv_tmpN)
                                        if IsNull(pcv_ShowDesc) or pcv_ShowDesc="" then
                                            pcv_ShowDesc="0"
                                        end if
                                        pClngShowItemImg=pcv_tmpArr(16,pcv_tmpN)
                                        if IsNull(pClngShowItemImg) or pClngShowItemImg="" then
                                            pClngShowItemImg="0"
                                        end if
                                        pClngSmImgWidth=pcv_tmpArr(17,pcv_tmpN)
                                        if IsNull(pClngSmImgWidth) or pClngSmImgWidth="" then
                                            pClngSmImgWidth="0"
                                        end if
                                        pClngShowSku=pcv_tmpArr(18,pcv_tmpN)
                                        if IsNull(pClngShowSku) or pClngShowSku="" then
                                            pClngShowSku="0"
                                        end if
                                        CATDesc=pcv_tmpArr(1,pcv_tmpN)
										requiredCategory=pcv_tmpArr(13,pcv_tmpN)
										CATNotes=pcv_tmpArr(29,pcv_tmpN)


                                        If strCol <> "class='pcBTOfirstRow row'" Then
                                            strCol = "class='pcBTOfirstRow row'"
                                        Else 
                                            strCol = "class='pcBTOsecondRow row'"
                                        End If %>
                                                
                                                <div class="panel-heading"><%=CATDesc%>
                                                    <% if requiredCategory=-1 then
                                                        ReqCAT=1
                                                    else
                                                        ReqCAT=0
                                                    end if%>																	
                                                </div>
                                                <div class="panel-body">
                                                <% ' If there are configuration instructions for this category, show them here.
                                                if CATNotes<>"" then%>
                                                    <div <%=strCol%>>
                                                        <div class="col-xs-12"><span class="catNotes"><%=CATNotes%></span></div>
                                                    </div>
                                                <%end if%>
                                                <% PrdCnt = 0 %>
                                                <% ShowInfoArray = ""
                                                showInfoVar="0"
												'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
												' START: SHOW CHECKBOXES WITH PRICE
												'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                                                pcv_tmpTest=1
            
                                                DO WHILE ((pcv_tmpTest=1) AND (pcv_tmpN<=pcv_ArrCount))
						
								intTempIdProduct=pcv_tmpArr(5,pcv_tmpN)

								'check to see if this option was checked
								SelectVar="0"
								tempQ=0
								for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
									if Clng(ArrProduct(i))=Clng(intTempIdProduct) then
										if Clng(ArrCategory(i))=Clng(checkVarCat) then
											SelectVar="1"
											'tempQ=ArrQuantity(i)
											tempQ=1
										end if
									end if
								next
													
								pcv_prdDesc=pcv_tmpArr(27,pcv_tmpN)
								pcv_prdSDesc=pcv_tmpArr(28,pcv_tmpN)
								if IsNull(pcv_prdSDesc) or trim(pcv_prdSDesc)="" then
									pcv_prdSDesc=pcv_prdDesc
								end if
								pcv_qtyvalid=pcv_tmpArr(3,pcv_tmpN)
								if isNULL(pcv_qtyvalid) OR pcv_qtyvalid="" then
									pcv_qtyvalid="0"
								end if
								pcv_minQty=pcv_tmpArr(4,pcv_tmpN)
								if isNULL(pcv_minQty) OR pcv_minQty="" then
									pcv_minQty="1"
								end if
								if pcv_minQty<"1" then
									pcv_minQty="1"
								end if
								prdBtoBPrice = pcv_tmpArr(10,pcv_tmpN)
								prdPrice = pcv_tmpArr(9,pcv_tmpN)
								if prdBtoBPrice=0 then
									prdBtoBPrice=prdPrice
								end if
								displayQF=pcv_tmpArr(14,pcv_tmpN)
								intTempIdCategory=pcv_tmpArr(0,pcv_tmpN)
								weight=pcv_tmpArr(6,pcv_tmpN)
								cdefault=pcv_tmpArr(12,pcv_tmpN)
								strDescription=pcv_tmpArr(7,pcv_tmpN)
								strSku=pcv_tmpArr(19,pcv_tmpN)
								strSmallImage=pcv_tmpArr(20,pcv_tmpN)							
								if strSmallImage = "" or strSmallImage = "no_image.gif" then
									strSmallImage = "hide"
								end if
								pstock=pcv_tmpArr(21,pcv_tmpN)
								pNostock=pcv_tmpArr(22,pcv_tmpN)	
								if pNostock = "" or pNoStock = null then
								 pNostock = 0
								end if
								pcv_intBackOrder = pcv_tmpArr(23,pcv_tmpN)
								pcv_intShipNDays = pcv_tmpArr(24,pcv_tmpN)
								pMinPurchase = pcv_tmpArr(25,pcv_tmpN)
																	
								strCategoryDesc=pcv_tmpArr(1,pcv_tmpN) 
								if pcv_tmpNewPath<>"" then
									pcv_tmpArr(11,pcv_tmpN)=0
								end if
								If pcv_tmpArr(11,pcv_tmpN)=True then
									showInfoVar="1"
								End If
						
								intCC_BTO_Pricing=0
								if session("customercategory")<>0 then
									query="SELECT pcCC_BTO_Price FROM pcCC_BTO_Pricing WHERE idcustomerCategory=" & session("customercategory") & " AND idBTOItem=" & intTempIdProduct& " AND idBTOProduct=" & pIdProduct & ";" 
									set rsCCObj=server.CreateObject("ADODB.RecordSet")
									set rsCCObj=conntemp.execute(query)
																		
									if err.number<>0 then
										call LogErrorToDatabase()
										set rsCCObj=nothing
										call closedb()
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if

									if NOT rsCCObj.eof then
										intCC_BTO_Pricing=1
										pcCC_BTO_Price=rsCCObj("pcCC_BTO_Price")
									else
										query="SELECT pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=" & session("customercategory") &" AND pcCC_Pricing.idProduct=" & intTempIdProduct & ";"
										set rsCCObj=server.CreateObject("ADODB.RecordSet")
										set rsCCObj=conntemp.execute(query)
										if NOT rsCCObj.eof then
											intCC_BTO_Pricing=1
											pcCC_BTO_Price=rsCCObj("pcCC_Price")
										end if
									end if
									set rsCCObj=nothing
								end if
		
								'customer logged in as ATB customer based on retail price
								if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
									prdPrice=Cdbl(prdPrice)-(pcf_Round(Cdbl(prdPrice)*(cdbl(session("ATBPercentage"))/100),2))
								end if
	
								'customer logged in as ATB customer based on wholesale price
								if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
									prdBtoBPrice=Cdbl(prdBtoBPrice)-(pcf_Round(Cdbl(prdBtoBPrice)*(cdbl(session("ATBPercentage"))/100),2))
									prdPrice=Cdbl(prdBtoBPrice)
								end if
								
								'customer logged in as a wholesale customer
								if prdBtoBPrice>0 and session("customerType")=1 then
									prdPrice=Cdbl(prdBtoBPrice)
								end if
								'customer logged in as a customer type with price different then the online price
								if intCC_BTO_Pricing=1 then
									if (pcCC_BTO_Price<>0) OR (pcCC_BTO_Price=0 AND intCC_BTO_Pricing=1) then
										prdPrice=Cdbl(pcCC_BTO_Price)
									end if
								end if
								
							if SelectVar="1" then
								tmp_selected=true
							else
								tmp_selected=false
							end if
							
							if tmp_selected then
								tmp_qty=tempQ*ProQuantity
							else
								tmp_qty=pcv_minQty*ProQuantity
							end if
							pcv_tmpIDiscount=0
							call CheckDiscount(intTempIdProduct,tmp_selected,tmp_qty,prdPrice)
							
							pcv_tmpCustomizedPrice=0
							if tmp_selected then
								'if cdefault<>"" and cdefault<>0 then
								'	pcv_tmpCustomizedPrice=(tempQ-pcv_minqty)*prdPrice
								'else
									pcv_tmpCustomizedPrice=tempQ*prdPrice
								'end if
								pcv_CustomizedPrice=pcv_CustomizedPrice+pcv_tmpCustomizedPrice
							else
								'if cdefault<>"" and cdefault<>0 then
								'	pcv_tmpCustomizedPrice=cdbl(-pcv_minqty*prdPrice)
								'	pcv_CustomizedPrice=pcv_CustomizedPrice+pcv_tmpCustomizedPrice
								'end if
							end if

						ShowInfoArray = ShowInfoArray & intTempIdProduct& ","
						ShowInfoArray = intTempIdProduct& ","   
						PrdCnt = PrdCnt + 1
						jCnt = jCnt + 1%>
						<input name="MS<%=jCnt%>" type="HIDDEN" value="<%=VarMS%>">
						<input name="currentValue<%=jCnt%>" type="HIDDEN" value="<%=pcv_tmpCustomizedPrice%>">
						<input name="Discount<%=jCnt%>" type="HIDDEN" value="<%=pcv_tmpIDiscount%>">
						<input name="CAT<%=jCnt%>" type="HIDDEN" value="CAG<%=tempCcat%>">
						
						<div <%=strCol%> style="vertical-align:top">  
							<div <%if (displayQF=True) then%>class="col-xs-3"<%else%>class="col-xs-2"<%end if%>> 
							<input type="hidden" name="Cat<%=intTempIdCategory%>_Prd<%=PrdCnt%>" value="<%=intTempIdProduct%>">
							<% If SelectVar="1" then %>
							<input type="checkbox" name="CAG<%=intTempIdCategory&intTempIdProduct%>" value="<%=intTempIdProduct%>_<%=prdPrice%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice%>" onClick="javscript:document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>QF.value='<%=pcv_minQty%>'; calculate(this,0);" checked class="clearBorder">
							<% else %>
							<input type="checkbox" name="CAG<%=intTempIdCategory&intTempIdProduct%>" value="<%=intTempIdProduct%>_<%=prdPrice%>_<%=pcv_tmpArr(6,pcv_tmpN)%>_<%=prdPrice%>" onClick="javscript:document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>QF.value='<%=pcv_minQty%>'; calculate(this,0);" class="clearBorder">
							<% end if
							
							SelectVar="0"
							RTestStr=RTestStr & vbcrlf & "if (document.additem.CAG"& intTempIdCategory & intTempIdProduct & ".checked !=false) { RTest" & CB_CatCnt & "=" & "RTest" & CB_CatCnt & "+document.additem.CAG" & intTempIdCategory & intTempIdProduct & ".checked; }"& vbcrlf
							%>
					
								<%if (displayQF=True) then%>
									<input class="form-control quantity" type="text" size="2" name="CAG<%=intTempIdCategory&intTempIdProduct%>QF" value="<%=tempQ%>" onblur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>)) calculate(document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>,0);">
								<%else%>
									<input type="hidden" name="CAG<%=intTempIdCategory&intTempIdProduct%>QF" value="<%=tempQ%>">
								<%end if%>
								<% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
									<img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
								<% end if %>
                            </div>
                            <div <%if (displayQF=True) then%>class="col-xs-4"<%else%>class="col-xs-5"<%end if%>>
								<span><%=strDescription%></span>
								<% if not pClngShowSku = 0 then %>
									<div class="pcSmallText"><%=strSku%></div>
								<% end if %>
							</div>
							<div class="col-xs-2"> 
								<div align="right">
								<%if pnoprices<2 then%><%=scCurSign & money(prdPrice)%><%end if%>
								</div>
							</div>
							

							<div class="col-xs-3">
							<% if showInfoVar = "1" then %>
							
								<% if iBTODetLinkType=1 then %>
								    <a class="" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')"><%=pcv_strBTODetTxt %></a>
								<%else%>
									<a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('<%=pcv_sffolder%>ShowChargesInfo.asp?SIArray=<%=ShowInfoArray%>&cd=<%=strCategoryDesc%>')">
                                        <span class="pcButtonText"><%=pcv_strBTODetTxt %></span>
									    <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>">
                                    </a>							
								<% end if 
								
							 end if %>
							 
							<%
							'Show Option Discounts icon
							ProductArray = Split(ShowInfoArray,",")
							for i = lbound(ProductArray) to (UBound(ProductArray)-1)
								if ProductArray(i)<>"" then
									MyTest=CheckOptDiscount(ProductArray(i))
									if MyTest=1 then
										if pnoprices<2 then%>
											<a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=ProductArray(i)%>')"><img alt="<%=dictLanguage.Item(Session("language")&"_viewPrd_16")%>" src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>"></a>
										<%end if
									end if
								end if
							next
							'End Show Option Discounts icon%> 
							</div>
						</div>
					
						<% if func_DisplayBOMsg <> "" then %>
							<div <%=strCol%>>
								<%=func_DisplayBOMsg1(tempVarCat)%>
							</div>
						<% end if %>
						<%if pcv_ShowDesc="1" then%>
						<div <%=strCol%>>
							<div <%if (displayQF=True) then%>class="col-xs-3"<%else%>class="col-xs-2"<%end if%>></div>
								<div class="col-xs-6">
									<span class="configDesc">
										<%=pcv_prdSDesc%>
									</span>
							</div>
						</div>
						<%end if%>
						
						<%pcv_tmpN=pcv_tmpN+1
						IF (pcv_tmpN<=pcv_ArrCount) THEN
						if Clng(pcv_tmpArr(0,pcv_tmpN))<>Clng(checkVarCat) then
							pcv_tmpTest=0
						end if
						END IF
						LOOP

						IF (pcv_tmpTest=0) AND (pcv_tmpN<=pcv_ArrCount) THEN
							pcv_tmpN=pcv_tmpN-1
						END IF
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' END: SHOW CHECKBOXES WITH PRICE
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						%>
						<input type="hidden" name="PrdCnt<%=tempCcat%>" value="<%=PrdCnt%>">
													
							<%RTestStr=RTestStr & vbcrlf & "if (RTest" & CB_CatCnt & " == '') { alert('"& dictLanguage.Item(Session("language")&"_alert_7") & replace(CATDesc,"'","\'") & "'); return(false);}" & vbcrlf
							if ReqCAT=1 then
								ReqTestStr=ReqTestStr & RTestStr
								ReqCAT=0
							end if%>
						<% 
						'=====================
						'End LOOP THROUGH PRODUCTS
						'===================== 
                        %>
                        </div></div>
                        <%
						end if
						'*****************************
						End If '**********************
						'*****************************
					pcv_tmpN=pcv_tmpN+1
		LOOP 'rsSSobj
	end if //'Have BTO Categories
	set rsSSobj=nothing
	'******* END BTO Categories
	'******************************************* 	
    %>
	
	<%
	
	response.write "<script type=text/javascript>" & VBCRlf
	response.write "function DisValue(IDPro,ProQ,ProP) {" & VBCRlf
	response.write "DisValue1=0;" & VBCRLf
	response.write "IDPro1=eval(IDPro);" & VBCRLf
	response.write "ProQ1=eval(ProQ);" & VBCRLf
	response.write "ProP1=eval(ProP);" & VBCRLf
	if TempDiscountStr<>"" then
	response.write TempDiscountStr & VBCRLf
	end if
	response.write "return(eval(DisValue1));" & VBCrlf
	response.write " } </script>" & VBCRlf
	
	response.write "<script type=text/javascript>" & VBCRlf
	response.write "function QDisValue(IDPro,ProQ,ProP) {" & VBCRlf
	response.write "DisValue1=0;" & VBCRLf
	response.write "IDPro1=eval(IDPro);" & VBCRLf
	response.write "ProQ1=ProQ.value;" & VBCRLf
	response.write "ProP1=eval(ProP);" & VBCRLf
	if TempQDStr<>"" then
	response.write TempQDStr & VBCRLf
	end if
	response.write "return(eval(DisValue1));" & VBCrlf
	response.write " } </script>" & VBCRlf
	%>
	<script type=text/javascript>
	function chkR()
	{
	<%if ReqTestStr<>"" then%>
	<%=ReqTestStr%>
	<%end if%>
	return(true);
	//return checkproqty(document.additem.quantity);
	}
	</script>
	<%Call pcs_GetDefaultBTOItemsMin%>
	<input type="hidden" name="FirstCnt" value="<%=jCnt%>">
	<input type="hidden" name="CB_CatCnt" value="<%=CB_CatCnt%>">
</div>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Configuration Table - Reconfig
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Prices - Reconfig
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_AddChargesPricesReconfig
	if (pPrice>0) and (pnoprices<>1) then
			if session("customerType")=1 then
			response.write "<b>" & dictLanguage.Item(Session("language")&"_viewPrd_15") & "</b>" & scCurSign & money(pPriceDefault)&"<br>" 
			else
			response.write "<b>" & bto_dictLanguage.Item(Session("language")&"_configurePrd_2") & "</b>" & scCurSign & money(pPriceDefault)&"<br>" 
			end if
	end if 
	%>
	<% if pnoprices=1 then %>
		<input name="GrandTotal2" type="hidden" value="<%=scCurSign%><%=money(pCMPrice+pPriceDefault)%>" >
		<input name="GrandTotal2QD" type="hidden" value="<%=scCurSign%><%=money(pCMWQD+pPriceDefault)%>" >
	<% else %>
		<b><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_3")%></b>
		<input name="GrandTotal2QD" type="TEXT" value="<%=scCurSign%><%=money(pCMWQD+pPriceDefault)%>"  readonly size="14" class="transparentField">
		<input name="GrandTotal2" type="hidden" value="<%=scCurSign%><%=money(pCMPrice+pPriceDefault)%>" >
	<%
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Prices - Reconfig
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Get Minimum Quantity of Default BTO Items
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_GetDefaultBTOItemsMin
Dim query,rs,pcArr,i,intCount,dCount

query="SELECT products.idproduct, products.pcprod_minimumqty,configSpec_products.cdefault FROM categories, products, categories_products INNER JOIN configSpec_products ON categories_products.idCategory=configSpec_products.configProductCategory WHERE (((configSpec_products.specProduct)="&pIdProduct&") AND ((configSpec_products.configProduct)=[categories_products].[idproduct]) AND ((categories_products.idCategory)=[categories].[idcategory]) AND ((categories_products.idProduct)=[products].[idproduct])) ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort;"
set rs=connTemp.execute(query)%>
<script type=text/javascript>
	var defitems = new Array();
	var defmin = new Array();
	var defset = new Array();
<%if not rs.eof then
	pcArr=rs.getRows()
	intCount=ubound(pcArr,2)
	set rs=nothing
	For i=0 to intCount%>
	defitems[<%=i%>]=<%=pcArr(0,i)%>;
	<%if IsNull(pcArr(1,i)) or pcArr(1,i)="" then
	pcArr(1,i)=0
	end if%>
	defmin[<%=i%>]=<%=pcArr(1,i)%>;
	<%if pcArr(2,i)<>0 then%>
		defset[<%=i%>]=1;
	<%else%>
		defset[<%=i%>]=0;
	<%end if%>
	<%Next%>
	defitemscount=<%=intCount%>;
<%else%>
	defitemscount=-1;
<%end if
set rs=nothing%>
</script>
<%End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Get Minimum Quantity of Default BTO Items
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_AddChargesDiscounts
	if pnoprices<2 then ' Does Not Hide Prices
	if pDiscountPerQuantity=-1 then 
	%>
	<div class="pcTable pcShowList">
	<% If session("customerType")="1" then %>
		<div class="pcTableRowFull">
			<a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=pidProduct%>&type=1')"><img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>" border="0"></a>&nbsp;
			<a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?idproduct=<%=pidProduct%>&type=1')"><%response.write dictLanguage.Item(Session("language")&"_viewPrd_16")%></a>
		</div>
	<% else %>
		<div class="pcTableRowFull">
			<a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=pidProduct%>')"><img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>" border="0"></a>&nbsp;
			<a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?idproduct=<%=pidProduct%>')"><%response.write dictLanguage.Item(Session("language")&"_viewPrd_16")%></a>
		</div>
	<% end if %>
	</div>
	<%
	end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Check Quantity Discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub CheckDiscount(DIDProduct,IsDefault,ItemQty,ItemPrice)
	dim rs,query,pcArr,intCount,i
	query="SELECT quantityFrom,quantityUntil,discountperUnit,percentage,discountperWUnit FROM discountsPerQuantity WHERE IDProduct=" & DIDProduct & ";"
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	pcv_tmpIDiscount1=0
	if not rs.eof then
		pcArr=rs.GetRows()
		intCount=ubound(pcArr,2)
		set rs=nothing
		TempStr1=""
		For i=0 to intCount
			QFrom=pcArr(0,i)
			QTo=pcArr(1,i)
			DUnit=pcArr(2,i)
			QPercent=pcArr(3,i)
			DWUnit=pcArr(4,i)
			if (DWUnit=0) and (DUnit>0) then
				DWUnit=DUnit
			end if
			

			
				if (clng(ItemQty)>=clng(QFrom)) AND (clng(ItemQty)<=clng(QTo)) then
					if QPercent="-1" then
						if session("customerType")=1 then
							pcv_tmpIDiscount1=ItemQty*ItemPrice*0.01*DWUnit
						else
							pcv_tmpIDiscount1=ItemQty*ItemPrice*0.01*DUnit
						end if
					else
						if session("customerType")=1 then
							pcv_tmpIDiscount1=ItemQty*DWUnit
						else
							pcv_tmpIDiscount1=ItemQty*DUnit
						end if
					end if
					IF IsDefault=true THEN
						pcv_tmpIDiscount=pcv_tmpIDiscount1
						pcv_ItemDiscounts=pcv_ItemDiscounts+pcv_tmpIDiscount
					END IF
				end if
			
			TempStr1="if ((IDPro1 == " & DIDProduct & ") && (ProQ1 >= " & QFrom & ") && (ProQ1 <= " & QTo & ")) {" & Vbcrlf
			if QPercent="-1" then
				if session("customerType")=1 then
					TempStr1=TempStr1 & "DisValue1=ProQ1*ProP1*0.01*" & DWUnit & ";" & vbcrlf
				else
					TempStr1=TempStr1 & "DisValue1=ProQ1*ProP1*0.01*" & DUnit & ";" & vbcrlf
				end if
			else
				if session("customerType")=1 then
					TempStr1=TempStr1 & "DisValue1=ProQ1*" & DWUnit & ";" & vbcrlf
				else
					TempStr1=TempStr1 & "DisValue1=ProQ1*" & DUnit & ";" & vbcrlf
				end if
			end if
			TempStr1=TempStr1 & "}" & vbcrlf
			TempDiscountStr=TempDiscountStr & TempStr1
		Next
	end if
	set rs=nothing
End Sub				
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Check Quantity Discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Check Option Discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~				
Public Function CheckOptDiscount(DIDProduct)
	dim query,rs
	if session("customerType")=1 then
		query="select discountPerUnit,discountPerWUnit from discountsPerQuantity where IDProduct=" & DIDProduct & " AND discountPerWUnit<>0;"
	else
		query="select discountPerUnit,discountPerWUnit from discountsPerQuantity where IDProduct=" & DIDProduct & " AND discountPerUnit<>0;"
	end if
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	CheckOptDiscount=0
	if not rs.eof then
		CheckOptDiscount=1
	end if
	set rs=nothing
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Check Option Discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Convert Numbers for Special Server
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function New_ConvertNum(tmpValue)
Dim tmp1
	tmp1=tmpValue
	if Instr(CStr(10/3),",")>0 then
		if Instr(tmp1,",")>Instr(tmp1,".") then
			tmp1=replace(tmp1,",",".")
		end if
	end if
	New_ConvertNum=tmp1
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Convert Numbers for Special Server
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Start SDBA
' START:  Display Back-Order Message
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Function func_DisplayBOMsg
	if isNULL(pMinPurchase) or pMinPurchase="" then
		pMinPurchase=0
	end if
	If (scOutofStockPurchase=-1) AND ((CLng(pStock)<1) OR (clng(pStock)<clng(pMinPurchase))) AND (Clng(pNoStock)=0) AND (Clng(pcv_intBackOrder)=1) Then
		If clng(pcv_intShipNDays)>0 then		  
			func_DisplayBOMsg = dictLanguage.Item(Session("language")&"_viewPrd_60")&dictLanguage.Item(Session("language")&"_sds_viewprd_1") & pcv_intShipNDays & dictLanguage.Item(Session("language")&"_sds_viewprd_1b")
		else
		 	func_DisplayBOMsg = "" 
		End if
	End If
End Function

Function func_DisplayBOMsg1(tempCat)
	if isNULL(pMinPurchase) or pMinPurchase="" then
		pMinPurchase=0
	end if
	If (scOutofStockPurchase=-1) AND ((CLng(pStock)<1) OR (clng(pStock)<clng(pMinPurchase))) AND (Clng(pNoStock)=0) AND (Clng(pcv_intBackOrder)=1) Then
		If clng(pcv_intShipNDays)>0 then		  
			func_DisplayBOMsg1 = "<span style=""padding-left: 22px;"" id=""AV" & tempCat & """>" & dictLanguage.Item(Session("language")&"_viewPrd_60")&dictLanguage.Item(Session("language")&"_sds_viewprd_1") & pcv_intShipNDays & dictLanguage.Item(Session("language")&"_sds_viewprd_1b") & "</span>"
		else
		 	func_DisplayBOMsg1 = "" 
		End if
	End If
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Display Back-Order Message
'End SDBA
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Cache Errors
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// This routine takes a proactive approach to handling History and Cache issue between browsers.
'// If we encounter an issue with History or Cache we redirect the page to itself. 
'// This is the only way to be 100% sure of the page's stability in all browsers.	
Public Sub pcs_BTOPageReLoader 
Dim tmpquery
	'// Calculate the Page URL
	'//Check URL Variables
	tmpquery=""
	For Each Item In Request.QueryString
		fieldName = getUserInput(Item,0)
		fieldValue = getUserInput(Request(Item),0)
		if fieldValue<>"" then
			fieldValue=replace(fieldValue,"''","'")
		end if
		if tmpquery<>"" then
			tmpquery=tmpquery & "&"
		end if
		tmpquery=tmpquery & fieldName &"=" &fieldValue
	Next
	If (Request.ServerVariables("HTTPS") = "off") Then
		pcv_srtRestoreFreshURL = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO") & "?" & tmpquery
	Else
		pcv_srtRestoreFreshURL = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO") & "?" & tmpquery
	End If
	'// Set a hidden field value. We will use this value as an environmental indicator. Its our "Coal Mine Canary".
	response.write "<form name=""pct"" id=""pct""><input name=""pctf"" value=""1"" style=""visibility:hidden""></form>"
	'// This block of javascript will check the hidden field value each time the page loads.
	response.write "<script type=""text/javascript"">"& chr(10)
	response.write "var intro=""1"";"& chr(10)
	response.write "var outro;"& chr(10)
	response.write "function pcf_isRefreshNeeded() {"& chr(10)
	response.write "   outro = (intro!=document.pct.pctf.value);"& chr(10)
	response.write "   document.pct.pctf.value=2;"& chr(10)
	response.write "   document.pct.pctf.defaultValue=2;" & chr(10)
	response.write "   pcf_doRefreshNeeded();"& chr(10)
	response.write "}"& chr(10)
	response.write "function pcf_doRefreshNeeded() {"& chr(10)
	response.write "	if (outro==true) {"& chr(10)
	response.write " 			// re-load the page"& chr(10)
	response.write "			 window.location = '"&pcv_srtRestoreFreshURL&"';"& chr(10)
	response.write "		} else {"& chr(10)
	response.write "	}"& chr(10)
	response.write "}"& chr(10)
	response.write "</script>"& chr(10)	
	response.write "<script type=""text/javascript"">"& chr(10)
	response.write "function pcf_WaitOnBody(pcv_OnloadFunction) {"& chr(10)
	response.write "var introOnload=window.onload;"& chr(10)
	response.write "if (typeof window.onload!='function') {"& chr(10)
	response.write "window.onload=pcv_OnloadFunction;"& chr(10)
	response.write "} else {"& chr(10)
	response.write "window.onload=function() {"& chr(10)
	response.write "if (introOnload) {"& chr(10)
	response.write "introOnload();"& chr(10)
	response.write "}"& chr(10)
	response.write "pcv_OnloadFunction();"& chr(10)
	response.write "}}}"& chr(10)
	response.write "pcf_WaitOnBody(pcf_isRefreshNeeded);"& chr(10)	
	response.write "</script>"& chr(10)
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Cache Errors
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	

'SB S
Public Sub pcs_SubscriptionProduct

	If pSubscriptionID <> 0  then
		
	  	If pIsLinked="1" Then
			%> <!--#include file="inc_sb_widget.asp"--> <%
		End If	  

	 	response.write "<input type=""hidden"" name=""pSubscriptionID"" id=""pcSubId"" value="""&pSubscriptionID&""">"
		
	End If
	
End Sub
'SB S
%>
