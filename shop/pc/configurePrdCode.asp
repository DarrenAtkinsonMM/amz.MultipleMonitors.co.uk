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
Dim strQtyCheck

strQtyCheck=""

Dim CheckAPPStr
CheckAPPStr=""

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
' START:  Conflict Management Module
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<script type="text/javascript">
BTOHaveRules=0;
</script>
<%
Dim pcv_HaveRules
pcv_HaveRules=0

Public Sub ConflictManager()
Dim query,tmpquery
tmpquery=" AND products.removed=0 AND products.active<>0 "
if (scOutOfStockPurchase="-1") AND (iBTOOutofStockPurchase="-1") then
	tmpquery=tmpquery & " AND ((products.stock>0) OR (products.nostock<>0) OR (products.pcProd_BackOrder<>0))"
end if

pcv_HaveRules=0

query="SELECT pcBTORules.pcBR_IDSourcePrd,pcBTORules.pcBR_isCAT,pcBTORules.pcBR_isCAT,pcBTORules.pcBR_isCAT,pcBTORules.pcBR_isCAT,pcBTORules.pcBR_isCAT,pcBTORules.pcBR_isCAT,pcBTORules.pcBR_Must_Exists,pcBTORules.pcBR_CanNot_Exists,pcBTORules.pcBR_CatMust_Exists,pcBTORules.pcBR_CatCanNot_Exists,pcBTORules.pcBR_ID,products.pcProd_showBTOCMMsg FROM Products INNER JOIN pcBTORules ON Products.idproduct=pcBTORules.pcBR_IDBTOPrd WHERE pcBTORules.pcBR_IDBTOPrd=" & pIdProduct & tmpquery & ";"
set rs=connTemp.execute(query)

intCount=0
RulesStr=""
ListPrd=""
IF not rs.eof then
	pcv_HaveRules=1
	pcv_A=rs.getRows()
	intCount=ubound(pcv_A,2)
	
	For i=0 to intCount
		if (pcv_A(3,i)<>"") and (pcv_A(3,i)="1") then
			query="select categoryDesc from Categories where idcategory=" & pcv_A(0,i)
			set rs=connTemp.execute(query)
			pcv_A(3,i)=ClearHTMLTags2(rs("categoryDesc"),0)
			set rs=nothing
		else
			query="select description from Products where idproduct=" & pcv_A(0,i)
			set rs=connTemp.execute(query)
			pcv_A(3,i)=ClearHTMLTags2(rs("description"),0)
			set rs=nothing
		end if
		
		pcv_A(1,i)=""
		pcv_A(2,i)=""
		pcv_A(4,i)=""
		pcv_A(5,i)=""
		
		'Get Rules Strings
		if (pcv_A(7,i)<>"") and (pcv_A(7,i)="1") then
			query="SELECT pcBRMust.pcBRMust_Item FROM Products INNER JOIN pcBRMust ON Products.idproduct=pcBRMust.pcBRMust_Item WHERE pcBRMust.pcBR_ID=" & pcv_A(11,i) & tmpquery & ";"
			set rsT=connTemp.execute(query)
			if not rsT.eof then
				tmpA=rsT.getRows()
				intCount3=ubound(tmpA,2)
				For j=0 to intCount3
					pcv_A(1,i)=pcv_A(1,i) & tmpA(0,j) & ","
				Next
			end if
			set rsT=nothing
		end if
		
		if (pcv_A(8,i)<>"") and (pcv_A(8,i)="1") then
			query="SELECT pcBRCanNot.pcBRCanNot_Item FROM Products INNER JOIN pcBRCanNot ON Products.idproduct=pcBRCanNot.pcBRCanNot_Item WHERE pcBRCanNot.pcBR_ID=" & pcv_A(11,i) & tmpquery & ";"
			set rsT=connTemp.execute(query)
			if not rsT.eof then
				tmpA=rsT.getRows()
				intCount3=ubound(tmpA,2)
				For j=0 to intCount3
					pcv_A(2,i)=pcv_A(2,i) & tmpA(0,j) & ","
				Next
			end if
			set rsT=nothing
		end if
		
		if (pcv_A(9,i)<>"") and (pcv_A(9,i)="1") then
			query="SELECT pcBRCatMust_Item FROM pcBRCatMust WHERE pcBR_ID=" & pcv_A(11,i) & ";"
			set rsT=connTemp.execute(query)
			if not rsT.eof then
				tmpA=rsT.getRows()
				intCount3=ubound(tmpA,2)
				For j=0 to intCount3
					pcv_A(4,i)=pcv_A(4,i) & tmpA(0,j) & ","
				Next
			end if
			set rsT=nothing
		end if
		
		if (pcv_A(10,i)<>"") and (pcv_A(10,i)="1") then
			query="SELECT pcBRCatCanNot_Item FROM pcBRCatCanNot WHERE pcBR_ID=" & pcv_A(11,i) & ";"
			set rsT=connTemp.execute(query)
			if not rsT.eof then
				tmpA=rsT.getRows()
				intCount3=ubound(tmpA,2)
				For j=0 to intCount3
					pcv_A(5,i)=pcv_A(5,i) & tmpA(0,j) & ","
				Next
			end if
			set rsT=nothing
		end if
		
	Next
	
	'ID Source Prd
	RulesStr="var Rule1=new Array();" & vbcrlf
	'Must List
	RulesStr=RulesStr & "var Rule2=new Array();" & vbcrlf
	'CanNot List
	RulesStr=RulesStr & "var Rule3=new Array();" & vbcrlf
	'Source Product Name
	RulesStr=RulesStr & "var Rule4=new Array();" & vbcrlf
	'Product was selected 1/0
	RulesStr=RulesStr & "var Rule5=new Array();" & vbcrlf
	'Must Categories List
	RulesStr=RulesStr & "var Rule6=new Array();" & vbcrlf
	'CanNot Categories List
	RulesStr=RulesStr & "var Rule7=new Array();" & vbcrlf
	'isCAT 1/0
	RulesStr=RulesStr & "var Rule8=new Array();" & vbcrlf
	'Numbers of Configurator Plus Rules
	RulesStr=RulesStr & "var RuleCount=" & intCount & ";" & vbcrlf
	'Show Configurator Plus Messages
	if IsNull(pcv_A(12,0)) or pcv_A(12,0)="" then
		pcv_A(12,0)=0
	end if
	RulesStr=RulesStr & "var ShowBTOCMMsg=" & pcv_A(12,0) & ";" & vbcrlf
	
	For k=0 to intCount
		RulesStr=RulesStr & "Rule1[" & k & "]='" & pcv_A(0,k) & "';" & vbcrlf
		RulesStr=RulesStr & "Rule2[" & k & "]='" & pcv_A(1,k) & "';" & vbcrlf
		RulesStr=RulesStr & "Rule3[" & k & "]='" & pcv_A(2,k) & "';" & vbcrlf
		RulesStr=RulesStr & "Rule4[" & k & "]='" & ClearHTMLTags2(replace(replace(pcv_A(3,k),"'","\'"),"&amp;","&"),0) & "';" & vbcrlf
		RulesStr=RulesStr & "Rule5[" & k & "]=0;" & vbcrlf
		RulesStr=RulesStr & "Rule6[" & k & "]='" & pcv_A(4,k) & "';" & vbcrlf
		RulesStr=RulesStr & "Rule7[" & k & "]='" & pcv_A(5,k) & "';" & vbcrlf
		RulesStr=RulesStr & "Rule8[" & k & "]=" & pcv_A(6,k) & ";" & vbcrlf
	Next
END IF
set rs=nothing

if RulesStr<>"" then
query="SELECT configSpec_products.configProduct,configSpec_products.configProductCategory,categories.categoryDesc FROM categories INNER JOIN (products INNER JOIN configSpec_products ON (products.idproduct=configSpec_products.configProduct" & tmpquery & ")) ON categories.idCategory = configSpec_products.configProductCategory WHERE configSpec_products.specProduct="&pIdProduct&" ORDER BY configSpec_products.configProductCategory ASC;"
set rs=connTemp.execute(query)
pcv_B=rs.getRows()
intCount1=ubound(pcv_B,2)
pcv_idCat=0
CatCount=-1
strPrds=""
RulesStr=RulesStr & "var CatID=new Array();" & vbcrlf
RulesStr=RulesStr & "var CatPrds=new Array();" & vbcrlf
RulesStr=RulesStr & "var CatName=new Array();" & vbcrlf
For k=0 to intCount1
	if clng(pcv_idCat)<>clng(pcv_B(1,k)) then
		if strPrds<>"" then
		RulesStr=RulesStr & "CatPrds[" & CatCount & "]='" & strPrds & "';" & vbcrlf
		strPrds=""
		end if
		CatCount=CatCount+1
		pcv_idCat=clng(pcv_B(1,k))
		RulesStr=RulesStr & "CatID[" & CatCount & "]='" & pcv_B(1,k) & "';" & vbcrlf
		RulesStr=RulesStr & "CatName[" & CatCount & "]='" & ClearHTMLTags2(replace(replace(pcv_B(2,k),"'","\'"),"&amp;","&"),0) & "';" & vbcrlf
		strPrds=strPrds & pcv_B(0,k) & ","
	else
		strPrds=strPrds & pcv_B(0,k) & ","
	end if
Next

if strPrds<>"" then
	RulesStr=RulesStr & "CatPrds[" & CatCount & "]='" & strPrds & "';" & vbcrlf
	strPrds=""
end if
RulesStr=RulesStr & "var CatCount=" & CatCount & ";" & vbcrlf
  %>
<script type="text/javascript">
BTOHaveRules=1;
<%=RulesStr%>
var pcv_dicProdOpt1="<%=dictLanguage.Item(Session("language")&"_prodOpt_1")%>";
var pcv_dicProdOpt2="<%=dictLanguage.Item(Session("language")&"_prodOpt_2")%>";
var pcv_msg_btocm_1="<%=bto_dictLanguage.Item(Session("language")&"_btocm_1")%>";
var pcv_msg_btocm_2="<%=bto_dictLanguage.Item(Session("language")&"_btocm_2")%>";
var pcv_msg_btocm_3="<%=bto_dictLanguage.Item(Session("language")&"_btocm_3")%>";
var pcv_msg_btocm_9="<%=bto_dictLanguage.Item(Session("language")&"_btocm_9")%>";
var pcv_msg_btocm_10="<%=bto_dictLanguage.Item(Session("language")&"_btocm_10")%>";
var pcv_msg_btocm_5="<%=bto_dictLanguage.Item(Session("language")&"_btocm_5")%>";
var pcv_msg_btocm_6="<%=bto_dictLanguage.Item(Session("language")&"_btocm_6")%>";
var pcv_msg_btocm_7="<%=bto_dictLanguage.Item(Session("language")&"_btocm_7")%>";
var pcv_msg_btocm_5a="<%=bto_dictLanguage.Item(Session("language")&"_btocm_5a")%>";
var pcv_msg_btocm_5b="<%=bto_dictLanguage.Item(Session("language")&"_btocm_5b")%>";
var pcv_msg_btocm_8a="<%=bto_dictLanguage.Item(Session("language")&"_btocm_8a")%>";
var pcv_msg_btocm_8b="<%=bto_dictLanguage.Item(Session("language")&"_btocm_8b")%>";
var pcv_msg_btocm_8c="<%=bto_dictLanguage.Item(Session("language")&"_btocm_8c")%>";
var pcv_msg_btocm_8d="<%=bto_dictLanguage.Item(Session("language")&"_btocm_8d")%>";
var pcv_msg_btocm_8e="<%=bto_dictLanguage.Item(Session("language")&"_btocm_8e")%>";
var pcv_msg_btocm_8f="<%=bto_dictLanguage.Item(Session("language")&"_btocm_8f")%>";
var pcv_msg_btocm_8g="<%=bto_dictLanguage.Item(Session("language")&"_btocm_8g")%>";
var pcv_msg_btocm_11="<%=bto_dictLanguage.Item(Session("language")&"_btocm_11")%>";
var pcv_msg_btocm_11a="<%=bto_dictLanguage.Item(Session("language")&"_btocm_11a")%>";
var pcv_msg_btocm_4="<%=bto_dictLanguage.Item(Session("language")&"_btocm_4")%>";
var pcv_msg_btocm_4a="<%=bto_dictLanguage.Item(Session("language")&"_btocm_4a")%>";
var pcv_msg_btocm_4b="<%=bto_dictLanguage.Item(Session("language")&"_btocm_4b")%>";
var pcv_msg_btocm_4c="<%=bto_dictLanguage.Item(Session("language")&"_btocm_4c")%>";
var pcv_msg_btocm_4d="<%=bto_dictLanguage.Item(Session("language")&"_btocm_4d")%>";
var pcv_msg_btocm_10="<%=bto_dictLanguage.Item(Session("language")&"_btocm_10")%>";
</script>
<script language="javascript" type="text/javascript" src="<%=pcf_getJSPath("../includes/javascripts","checkRules.js")%>"></script>
<%
else
pcv_HaveRules=0%>
<script type="text/javascript">
BTOHaveRules=0;
</script>
<%end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Conflict Management Module
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
			
			tmpSubList = ""
			If statusAPP="1" Then

				query="SELECT idproduct FROM Products WHERE pcProd_ParentPrd=" & DIDProduct & " AND removed=0 AND active=0 AND pcProd_SPInActive=0;"
				set rs=connTemp.execute(query)
				if not rs.eof then
					tmpArr=rs.getRows()
					tmpCount=ubound(tmpArr,2)
					For j=0 to tmpCount
						tmpSubList=tmpSubList & " || (IDPro1 == " & tmpArr(0,j) & ")"
					Next
				end if
				set rs=nothing

			End If
			
			TempStr1="if (((IDPro1 == " & DIDProduct & ")" & tmpSubList & ") && (ProQ1 >= " & QFrom & ") && (ProQ1 <= " & QTo & ")) {" & Vbcrlf

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
' START:  Check Quantity Discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_CheckQTYDiscount

	TempDiscountStr=""
	TempQDStr=""

	query="SELECT quantityFrom,quantityUntil,discountperUnit,percentage,discountperWUnit,baseproductonly FROM discountsPerQuantity WHERE IDProduct=" & pIDProduct
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
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
			pcv_IncOpt=pcArr(5,i)

			tmpSubList=""
			If statusAPP="1" Then

				query="SELECT idproduct FROM Products WHERE pcProd_ParentPrd=" & pIDProduct & " AND removed=0 AND active=0 AND pcProd_SPInActive=0;"
				set rs=connTemp.execute(query)			
				if not rs.eof then
					tmpArr=rs.getRows()
					tmpCount=ubound(tmpArr,2)
					For j=0 to tmpCount
						tmpSubList=tmpSubList & " || (IDPro1 == " & tmpArr(0,j) & ")"
					Next
				end if
				set rs=nothing
			
			End If

			TempStr1="if (((IDPro1 == " & pIDProduct & ")" & tmpSubList & ") && (ProQ1 >= " & QFrom & ") && (ProQ1 <= " & QTo & ")) {" & Vbcrlf

			if QPercent="-1" then
				if session("customerType")=1 then
					TempStr1=TempStr1 & "DisValue1=ProP1*0.01*" & DWUnit & ";" & vbcrlf
				else
					TempStr1=TempStr1 & "DisValue1=ProP1*0.01*" & DUnit & ";" & vbcrlf
				end if
			else
				if session("customerType")=1 then
					TempStr1=TempStr1 & "DisValue1=ProQ1*" & DWUnit & ";" & vbcrlf
				else
					TempStr1=TempStr1 & "DisValue1=ProQ1*" & DUnit & ";" & vbcrlf
				end if
			end if
			TempStr1=TempStr1 & "}" & vbcrlf
			TempQDStr=TempQDStr & TempStr1
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
' START:  create javascript for calculations
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_BTOJavaCalculations
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
' END:  create javascript for calculations
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Disallow purchasing. Quote Submission only
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_BTOPurchasing
	if (ProQuantity="" or ProQuantity="0") and pSavedQuantity<>"" then
		ProQuantity = pSavedQuantity
	end if	
	

	xOptionsCnt = 0 '// no options on this page
	
	pcv_strFuntionCall = "cdDynamic"
	
	If BTOCharges=0 then
		myImg=rslayout("addtocart")
		btnText=dictLanguage.Item(Session("language")&"_css_addtocart")
		btnStyle="pcButtonAddToCart"
	Else
		myImg=rslayout("submit")
		btnText=dictLanguage.Item(Session("language")&"_css_submit")
		btnStyle="pcButtonSubmit"
	End if
	
	if (iBTOQuoteSubmitOnly=1) or (pnoprices>0) then%>
	<div class="row">
		<div class="col-xs-12">
			<input class="form-control quantity" type="text" id="quantity" name="quantity" onBlur="if (checkproqty(this)) New_AutoUpdateQtyPrice();" size="4" value="<%=ProQuantity%>">
	
		<% If pserviceSpec <> 0 then %>

			<%if BTOCharges=1 then%>
				<button class="pcButton <%=btnStyle%>" id="addtocart" name="add" disabled="disabled" style="vertical-align: bottom;">
					<img src="<%=pcf_getImagePath("",myImg)%>" alt="<%=btnText%>" />
					<span class="pcButtonText"><%=btnText%></span>
				</button>
			<%end if%>
		<% End If %>
		</div>
	</div>
	<%
	else
	%>
	<div class="row">
		<%
		if xrequired="1" then %>
			<div class="col-xs-12">
				<input class="form-control quantity" type="text" id="quantity" name="quantity" onBlur="if (checkproqty(this)) New_AutoUpdateQtyPrice();" size="5" maxlength="10" value="<%=ProQuantity%>">
				<button class="pcButton <%=btnStyle%>" id="addtocart" name="add" onClick="javascript: if (checkproqty(document.additem.quantity)) {if (chkR()) {<%=pcv_strFuntionCall%>(<%=reqstring%>,0);}} return false">
					<img src="<%=pcf_getImagePath("",myImg)%>" alt="<%=btnText%>" />
					<span class="pcButtonText"><%=btnText%></span>
				</button>
			</div>
		<% 
		else 
		%>		
			<div class="col-xs-12">
				<input class="form-control quantity" type="text" id="quantity" name="quantity" onBlur="if (checkproqty(this)) New_AutoUpdateQtyPrice();" size="5" maxlength="10" value="<%=ProQuantity%>">
				<button class="pcButton <%=btnStyle%>" id="addtocart" name="add" disabled="disabled" style="vertical-align: bottom;">
					<img src="<%=pcf_getImagePath("",myImg)%>" alt="<%=btnText%>" />
					<span class="pcButtonText"><%=btnText%></span>
				</button>
			</div>
		<% 
		end if
		%>
	</div>
	<%
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Disallow purchasing. Quote Submission only
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show SKU
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ShowSKU
IF pHideSKU<>"1" THEN%>
	<div class="pcShowProductSku">
		<%=dictLanguage.Item(Session("language")&"_viewCat_P_8")%>: <%=pSku%>
	</div>
<%END IF
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show SKU
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Custom Search Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_CustomSearchFields
Dim query,rs,pcArr,intCount,i
	query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & pIdProduct & " AND pcSearchFieldShow=1 ORDER BY pcSearchFields.pcSearchFieldOrder ASC,pcSearchFields.pcSearchFieldName ASC;"
	set rs=connTemp.execute(query)
	IF not rs.eof THEN
		pcArr=rs.getRows()
		set rs=nothing
		intCount=ubound(pcArr,2)
		response.Write("<div style='padding-top: 5px;'></div>")
		For i=0 to intCount
				response.write "<div class='pcShowProductCustSearch'>"&pcArr(1,i)&": <a href='showsearchresults.asp?customfield="&pcArr(0,i)&"&SearchValues="&Server.URLEncode(pcArr(2,i))&"'>"&pcArr(3,i)&"</a></div>"
		Next
	END IF
	set rs=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Custom Search Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Brand (If assigned)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ShowBrand
	if sBrandPro="1" then
		if (pIDBrand&""<>"") and (pIDBrand&""<>"0") then
		response.write "<div class='pcShowProductBrand'>"
		response.write dictLanguage.Item(Session("language")&"_viewPrd_brand")
		%>
			<a href="viewBrands.asp?idBrand=<%=pIDBrand%>">
				<%=BrandName%>
			</a>
		<% 
		response.write "</div>"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Brand
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Units in Stock (if on, show the stock level here)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_UnitsStock
	if scdisplayStock=-1 AND pNoStock=0 then
		if pstock > 0 then
			response.write "<div class='pcShowProductStock'>"
			response.write dictLanguage.Item(Session("language")&"_viewPrd_19") & " " & pStock
			response.write "</div>"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Units in Stock (if on, show the stock level here)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Free Shipping Text
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_NoShippingText
	if scorderlevel <> "0" then
	else
		' Check to see if the product is set as a Non-Shipping Item and display message if product is for sale
		if pnoshipping="-1" and (pFormQuantity <> "-1" or NotForSaleOverride(session("customerCategory"))=1) and pnoshippingtext="-1" then 
			response.write "<div class='pcShowProductNoShipping'>"
			response.write dictLanguage.Item(Session("language")&"_viewPrd_8")
			response.write "</div>"
		end if
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Free Shipping Text
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Out of Stock Message
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_OutStockMessage
	' if out of stock and show message is enabled (-1) then show message unless stock is ignored
	if (scShowStockLmt=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=0) OR (pserviceSpec<>0 AND scShowStockLmt=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=0) then
		response.write "<div>"&dictLanguage.Item(Session("language")&"_viewPrd_60")&dictLanguage.Item(Session("language")&"_viewPrd_7")& "</div>"
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Out of Stock Message
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Start SDBA
' START:  Display Back-Order Message
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_DisplayBOMsg
	If (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_intBackOrder=1) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_intBackOrder=1) Then
		If clng(pcv_intShipNDays)>0 then
			response.write "<div>"&dictLanguage.Item(Session("language")&"_viewPrd_60")&dictLanguage.Item(Session("language")&"_sds_viewprd_1") & pcv_intShipNDays & dictLanguage.Item(Session("language")&"_sds_viewprd_1b") & "</div>"
		End if
	End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Display Back-Order Message
'End SDBA
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Get Additional Images Array
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
function pcf_GetAdditionalImages
	' // SELECT DATA SET
	' TABLES: pcProductsImages
	' COLUMNS ORDER: pcProductsImages.pcProdImage_Url, pcProductsImages.pcProdImage_LargeUrl, pcProductsImages.pcProdImage_Order
	
	query = 		"SELECT pcProductsImages.pcProdImage_Url, pcProductsImages.pcProdImage_LargeUrl, pcProductsImages.pcProdImage_Order "
	query = query & "FROM pcProductsImages "
	query = query & "WHERE pcProductsImages.idProduct=" & pidProduct &" "
	query = query & "ORDER BY pcProductsImages.pcProdImage_Order;"	
	set rs=server.createobject("adodb.recordset")
	set rs=conntemp.execute(query)	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	If rs.EOF Then
		pcf_GetAdditionalImages = ""
	Else
		pcf_GetAdditionalImages = ""
		Dim xCounter '// declare a temporary counter
		xCounter = 0
		do while NOT rs.EOF
		
		pcv_strProdImage_Url = ""
		pcv_strProdImage_LargeUrl = ""
		pcv_strProdImage_Url = rs("pcProdImage_Url")
		pcv_strProdImage_LargeUrl = rs("pcProdImage_LargeUrl")
			
			if len(pcv_strProdImage_Url)>0 then
			xCounter = xCounter + 1
				if xCounter > 1 then
					pcf_GetAdditionalImages = pcf_GetAdditionalImages & ","
				end if
				'// Add a sorted item onto the end of the string
				pcf_GetAdditionalImages = pcf_GetAdditionalImages & pcv_strProdImage_Url & "," & pcv_strProdImage_LargeUrl
			end if

		rs.movenext 
		loop		
	End If
	set rs=nothing
end function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Get Additional Images Array
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_MakeAdditionalImage

	'// Make the popup link, but dont set large image preference if the large image doesnt exist
	If len(pcv_strShowImage_LargeUrl)>0 Then		
		pcv_strLargeUrlPopUp= "javascript:pcAdditionalImages('"&pcf_getImagePath("catalog",pcv_strShowImage_LargeUrl)&"','"&pidProduct&"')" 
	Else
		pcv_strShowImage_LargeUrl = pcv_strShowImage_Url '// we dont have one, show the regular size
		pcv_strLargeUrlPopUp= "javascript:pcAdditionalImages('"&pcf_getImagePath("catalog",pcv_strShowImage_Url)&"','"&pidProduct&"')" 
	End If
	
	if pcv_strPopWindowOpen = 1 then
		%>
		<a href="#">	
			<img onmouseover='javascript:window.document.mainimg.src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pcv_strShowImage_LargeUrl)%>";' src='<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pcv_strShowImage_Url)%>' alt="<%=replace(pDescription,"""","&quot;")%>" />		
		</a> 
	<% else %>
			<%	
			'// Use Enhanced Views
			If pcv_strUseEnhancedViews = True Then 
			%>
				<a href="<%=pcf_getImagePath("catalog",pcv_strShowImage_LargeUrl)%>" class="highslide" onclick="return hs.expand(this, { slideshowGroup: 'slides' })" id="<%=bcounter%>"><img onmouseover='javascript:window.document.mainimg.src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pcv_strShowImage_Url)%>";' src='<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pcv_strShowImage_Url)%>' alt="<%=replace(pDescription,"""","&quot;")%>" /></a>
                <% if pcv_strUseEnhancedViews = True then %>
                	<div class="<%=pcv_strHighSlide_Heading%>"><%=replace(pDescription,"""","&quot;")%></div>
                <% end if %>
        	<%
			'// Use Pop Window 
			Else 
				%>	
                <a href="<%=pcv_strLargeUrlPopUp%>"><img onmouseover='javascript:window.document.mainimg.src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pcv_strShowImage_Url)%>";' src='<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pcv_strShowImage_Url)%>' alt="<%=replace(pDescription,"""","&quot;")%>" /></a> 
        <% End If		
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Additional Product Images (If there are any)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_AdditionalImages

if len(pImageUrl) > 0 then ' // only if there is a main image can there be additional images.
	pcv_strAdditionalImages = pcf_GetAdditionalImages '// set variable to array of images, if there are any
	if len(pcv_strAdditionalImages)>0 then '// there is a main, are there additionals?
	%>
	<div class="pcShowAdditional">
		<%
		'// the main image to the first place in the image set
		pcv_strAdditionalImages = pImageUrl & "," & pLgimageURL & "," & pcv_strAdditionalImages
		
		Dim pcArray_AdditionalImages '// declare a temporary array
		pcArray_AdditionalImages = Split(pcv_strAdditionalImages,",")	
		
		bCounter = 1
		
		'// When the product has additional images, this variable defines how many thumbnails are shown per row, below the main product image
		if pcv_intProdImage_Columns="" then
			pcv_intProdImage_Columns = 3
		end if
		
		modnum = pcv_intProdImage_Columns '// Get this from the db
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START Loop
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		For cCounter = LBound(pcArray_AdditionalImages) TO UBound(pcArray_AdditionalImages)
		
		'// Check if we have a normal image
		Dim pcv_strTempAssignment	
		pcv_strTempAssignment = ""
		pcv_strTempAssignment = pcArray_AdditionalImages(cCounter)
		pcv_strShowImage_Url = pcv_strTempAssignment '// we have one, set it
		
		cCounter = cCounter + 1 '// now get the large image
			
		'// Do Not generate an additional image if there is not one
		If len(pcv_strShowImage_Url)>0 Then
		
				'// Check if we have a large image
				pcv_strTempAssignment = ""	
				pcv_strTempAssignment = pcArray_AdditionalImages(cCounter)
				pcv_strShowImage_LargeUrl = pcv_strTempAssignment '// we have one
				
				if not bCounter mod modnum = 0 then%>
					<%pcs_MakeAdditionalImage%>
				<% Else %>
					<%pcs_MakeAdditionalImage%>
				<% end if		
				bCounter = bCounter + 1
		End If	
		
		Next
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END Loop
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		%>
	</div>
	<% if pcv_strPopWindowOpen <> 1 then %>
		<div style="padding-bottom: 10px;">(<i><%=dictLanguage.Item(Session("language")&"_viewPrd_28")%></i>)</div>
	<% end if %>
	<%	end if '// end if len(pcf_GetAdditionalImages)>0 then
end if	'// end if len(pImageUrl) > 0 then
%>
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
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Additional Product Images (If there are any)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Add button in the Control Panel
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_AdmBTOPurchasing

	if (ProQuantity="" or ProQuantity="0") and pSavedQuantity<>"" then
		ProQuantity = pSavedQuantity
	end if

	xOptionsCnt = 0 '// no options on this page
	
	pcv_strFuntionCall = "cdDynamic"
	
	myImg=rslayout("submit")
	btnText=dictLanguage.Item(Session("language")&"_css_submit")
	btnStyle="pcButtonSubmit"
	
	if (iBTOQuoteSubmitOnly=1) or (pnoprices>0) then%>
	<div class="pcFormItem">
		<input type="text" id="quantity" name="quantity" onBlur="if (checkproqty(this)) New_AutoUpdateQtyPrice();" size="4" value="<%=ProQuantity%>">
		<% If pserviceSpec <> 0 then %>
			<%if BTOCharges=1 then%>
				<%if request("idquote")<>"" then%>
					<button class="pcButton <%=btnStyle%>" id="addtocart" name="add">
						<img src="<%=pcf_getImagePath("",myImg)%>" alt="<%=btnText%>" />
						<span class="pcButtonText"><%=btnText%></span>
					</button>
				<%end if%>
			<%end if%>
		<% End If %>
	</div>
	<%
	else
	%>
	<div class="pcFormItem">
		<%if xrequired="1" then %>
			<input type="text" id="quantity" name="quantity" onBlur="if (checkproqty(this)) New_AutoUpdateQtyPrice();" size="5" maxlength="10" value="<%=ProQuantity%>">
			<button class="pcButton <%=btnStyle%>" id="addtocart" name="add" onClick="javascript: if (checkproqty(document.additem.quantity)) {if (chkR()) {<%=pcv_strFuntionCall%>(<%=reqstring%>,0);}} return false">
				<img src="<%=pcf_getImagePath("",myImg)%>" alt="<%=btnText%>" />
				<span class="pcButtonText"><%=btnText%></span>
			</button>
		<%else%>		
			<input  type="text" id="quantity" name="quantity" onBlur="if (checkproqty(this)) New_AutoUpdateQtyPrice();" size="5" maxlength="10" value="<%=ProQuantity%>">
			<%if request("idquote")<>"" OR pcv_strAdminPrefix="1" then%>
			<button class="pcButton <%=btnStyle%>" id="addtocart" name="add">
				<img src="<%=pcf_getImagePath("",myImg)%>" alt="<%=btnText%>" />
				<span class="pcButtonText"><%=btnText%></span>
			</button>
			<%end if%>
		<%end if%>
	</div>
	<%
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Add button in the Control Panel
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  product configuration table - Reconfigure
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_BTOReconfigTable
    Dim query,rsSSObj,tmpquery

    tmpquery=""
    if (scOutOfStockPurchase="-1") AND (iBTOOutofStockPurchase="-1") then
        tmpquery=" AND ((products.stock>0) OR (products.nostock<>0) OR (products.pcProd_BackOrder<>0))"
    end if
    
    call CreateAppPopUp()
    query="SELECT categories.idCategory, categories.categoryDesc, configSpec_products.multiSelect,products.pcprod_qtyvalidate,products.pcprod_minimumqty,products.idproduct, products.weight, products.description, configSpec_products.prdSort, configSpec_products.price, configSpec_products.Wprice, configSpec_products.showInfo, configSpec_products.cdefault, configSpec_products.requiredCategory, configSpec_products.displayQF,configSpec_products.pcConfPro_ShowDesc,configSpec_products.pcConfPro_ShowImg,configSpec_products.pcConfPro_ImgWidth,configSpec_products.pcConfPro_ShowSKU,products.sku,products.smallImageUrl,products.stock,products.noStock, products.pcProd_BackOrder, products.pcProd_ShipNDays,products.pcprod_minimumqty,configSpec_Products.pcConfPro_UseRadio,products.pcProd_multiQty,products.pcProd_Apparel,products.details,products.sDesc,configSpec_products.Notes FROM categories INNER JOIN (products INNER JOIN configSpec_products ON (products.idproduct=configSpec_products.configProduct AND products.active<>0 AND products.removed=0" & tmpquery & ")) ON categories.idCategory = configSpec_products.configProductCategory WHERE configSpec_products.specProduct="&pIdProduct&" ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort,products.description;"
    
    set rsSSObj=conntemp.execute(query)
    displayQF="0"
    funcTestCat=""
    %>
    <div class="">
        <%
        dim jCnt
        CB_CatCnt = 0
        jCnt=0
        
        '*******************************************
        '******* START Configurator Categories
    
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
                
    
                        pcv_HaveApparel=0
    
                
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
                        query="SELECT configSpec_products.configProduct,configSpec_products.price, configSpec_products.Wprice, configSpec_products.cdefault FROM configSpec_products WHERE configSpec_products.configProductCategory="&tempVarCat&" AND configSpec_products.specProduct="&pIdProduct&" AND configSpec_products.cdefault<>0;"
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
                        CATNotes=pcv_tmpArr(31,pcv_tmpN)
                        if CATNotes <> "" then
                        %>
                        <div <%=strCol%>>
                            <div class="col-xs-12"><span class="catNotes"><%=CATNotes%></span></div>
                        </div>
                        <%
                        end if
                        %>
    
                        <%'BTOCM-S%>
                        <div class="row">
                            <span name="CMMsg<%=pcv_tmpArr(0,pcv_tmpN)%>" id="CMMsg<%=pcv_tmpArr(0,pcv_tmpN)%>"></span>
                        </div>
                        <%'BTOCM-E%>
    
                        <%
                        pBTODisplayType=pcv_tmpArr(26,pcv_tmpN)
                        if IsNull(pBTODisplayType) or pBTODisplayType="" then
                            pBTODisplayType=1
                        end if
                        
                        displayQF=pcv_tmpArr(14,pcv_tmpN)
                        requiredCategory=pcv_tmpArr(13,pcv_tmpN)
                        if pcv_tmpNewPath<>"" then
                            pcv_tmpArr(11,pcv_tmpN)=0
                        end if
                        showInfo=pcv_tmpArr(11,pcv_tmpN)
                        %>
                        
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
                                tempQ=ArrQuantity(i)
                            end if
                        next%>
                                     
                        <% 
                        myApparel=""

                        if pBTODisplayType=1 then
                        
                            pcv_ListForGenInfo=pcv_ListForGenInfo & "GenDropInfo(document.additem.CAG" & tempVarCat & ");" & vbcrlf %>
                            
                            <div <%=strCol%>>
                            
                                <% 
                                '// CLASS CONFIGURATION: SELECT W. NO IMAGES (RC)
                                If (pcv_tmpArr(14,pcv_tmpN)=true) Then 
                                    pcv_strColumn1 = "col-xxs-12 col-xs-12 col-sm-1"
                                    pcv_strColumn2 = "col-xxs-12 col-xs-12 col-sm-9"
                                Else 
                                    pcv_strColumn1 = "col-xxs-12 col-xs-12 col-sm-1"
                                    pcv_strColumn2 = "col-xxs-12 col-xs-12 col-sm-9"
                                End If 
                                %>    
                                <div class="<%=pcv_strColumn1%>">
                                    
                                    <% '// Row 1: Quantity %>
                                    <% if pcv_tmpArr(14,pcv_tmpN)=true then %>
                                        <input class="form-control quantity" type="text" size="2" id="CAG<%=tempVarCat%>QF" name="CAG<%=tempVarCat%>QF" value="<%=tempQ%>" onblur="javascript:testdropqty(this,'document.additem.CAG<%=tempVarCat%>');">
                                    <% else %>
                                        <input type="hidden" id="CAG<%=tempVarCat%>QF" name="CAG<%=tempVarCat%>QF" value="<%=tempQ%>">
                                    <% end if %>
                                </div>
                                <div class='<%=pcv_strColumn2%>'>

                                    <input type=hidden name="app_IDProduct_CAG<%=tempVarCat%>" value="0">
                                    <input type=hidden name="app_Price_CAG<%=tempVarCat%>" value="0">
                                    <input type=hidden name="app_AddPrice_CAG<%=tempVarCat%>" value="0">
                                    <input type=hidden name="app_VIndex_CAG<%=tempVarCat%>" value="0">
                                
                                    <% '// Row 2: Select %>
                                    <select class="form-control" name="CAG<%=tempVarCat%>" onChange="testCAG<%=tempVarCat%>(); testdropdown('document.additem.CAG<%=tempVarCat%>'); CheckPreValue(this,1,0); showAvail<%=tempVarCat%>(this);">
    
                            </div>
                            <div class="col-xxs-12 col-xs-12 col-sm-2">
                                <%
                                HiddenFields=""
                                %>
                                
                        <% else '// if pBTODisplayType=1 then %>
                        
                            <input type="hidden" name="CAG<%=tempVarCat%>QF" value="<%=tempQ%>">
                            <%
                            pcv_ListForGenInfo=pcv_ListForGenInfo & "GenRadioExtInfo(document.additem.CAG" & tempVarCat & ");" & vbcrlf
                                if Clng(requiredCategory)<>0 then
                                    RTestStr="totalradio=document.additem.CAG" & tempVarCat & ".length;" & vbcrlf
                                    RTestStr=RTestStr & "RadioChecked=0;" & vbcrlf
                                    RTestStr=RTestStr & "if (totalradio>0) {" & vbcrlf
                                    RTestStr=RTestStr & "for (var mk=0;mk<totalradio;mk++) {" & vbcrlf
                                    RTestStr=RTestStr & "if (document.additem.CAG" & tempVarCat & "[mk].checked==true) { RadioChecked=1; break; } }" & vbcrlf
                                    RTestStr=RTestStr & "} else { if (document.additem.CAG" & tempVarCat & ".checked==true) RadioChecked=1;}" & vbcrlf
                                    RTestStr=RTestStr & "if (RadioChecked==0) {alert('"& dictLanguage.Item(Session("language")&"_alert_7") & replace(pcv_tmpArr(1,pcv_tmpN),"'","\'") & "'); return(false);}" & vbcrlf
                                    ReqTestStr=ReqTestStr & RTestStr
                                end if
                                %>
                                
                        <% end if '// if pBTODisplayType=1 then %>
                            
                        <%
                        dim requiredVar, showInfoVar, ShowInfoArray, SelectedVar
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
                        
                            cdefault=pcv_tmpArr(12,pcv_tmpN)
                            weight=pcv_tmpArr(6,pcv_tmpN)
                            pcv_Apparel=pcv_tmpArr(28,pcv_tmpN)
                            if IsNull(pcv_Apparel) or pcv_Apparel="" then
                                pcv_Apparel=0
                            end if
                            if pcv_Apparel="1" then
                                pcv_HaveApparel=1
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
                            pcv_multiQty=pcv_tmpArr(27,pcv_tmpN)
                            if isNULL(pcv_multiQty) OR pcv_multiQty="" then
                                pcv_multiQty="0"
                            end if
                            pcv_prdDesc=pcv_tmpArr(29,pcv_tmpN)
                            pcv_prdSDesc=pcv_tmpArr(30,pcv_tmpN)
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
                            
                            prdPrice1=prdPrice							
                        
                            '// START: Apparel
                            If statusAPP="1" Then
                                sp_IDProduct=0
                                sp_PrdName=""
                                if pcv_Apparel="1" then
                                    HaveDiffPrice=0
                                    queryQ="SELECT price,btoBPrice FROM Products WHERE idProduct=" & intTempIdProduct & ";"
                                    set rsQ=connTemp.execute(queryQ)
                                    if not rsQ.eof then
                                        pcParentPrice=rsQ("price")
                                        pcParentWPrice=rsQ("btoBPrice")
                                        set rsQ=nothing
                                        queryQ="SELECT price,Wprice FROM configSpec_Products WHERE specProduct=" & pIDProduct & " AND configProduct=" & intTempIdProduct  & " AND configProductCategory=" & tempVarCat & ";"
                                        set rsQ=connTemp.execute(queryQ)
                                        if not rsQ.eof then
                                            pcPriceConf=rsQ("price")
                                            pcWPriceConf=rsQ("Wprice")
                                            set rsQ=nothing
                                            
                                            if ccur(pcPriceConf)<>ccur(pcParentPrice) then
                                                HaveDiffPrice=1
                                            else
                                                if (ccur(pcWPriceConf)<>ccur(pcParentWPrice)) AND (ccur(pcWPriceConf)<>ccur(pcParentPrice)) then
                                                    HaveDiffPrice=1
                                                end if
                                            end if
                                        end if
                                        set rsQ=nothing
                                    end if
                                    set rsQ=nothing
                                    
                                    query="SELECT idproduct,description,price,btoBPrice,stock,noStock,pcProd_BackOrder,pcProd_ShipNDays FROM products WHERE pcProd_ParentPrd=" & intTempIdProduct & " AND removed=0 AND pcProd_SPInActive=0 AND idproduct=" & tempPrd 
                                    set rsA=connTemp.execute(query)
                                    if not rsA.eof then
                                        sp_IDProduct=rsA("idproduct")
                                        sp_PrdName=rsA("description")
                                        sp_PrdName=replace(sp_PrdName,"""","&quot;")
                                        prdPrice = Cdbl(rsA("price"))
                                        prdWPrice=rsA("btoBPrice")
                                        if IsNull(prdWPrice) or prdWPrice="" then
                                            prdWPrice=0
                                        end if
                                        if Cdbl(prdWPrice)=0 then
                                            prdWPrice=prdPrice
                                        end if
                                        pstock=rsA("stock")
                                        pNostock=rsA("nostock")	
                                        if pNostock = "" or pNoStock = null then
                                            pNostock = 0
                                        end if						
                                        pcv_intBackOrder = rsA("pcProd_BackOrder")							
                                        pcv_intShipNDays = rsA("pcProd_ShipNDays")
                                                                                    
                                        pcv_SPPrice1=CheckParentPrices(sp_IDProduct,prdPrice,prdWPrice,0)
                                        pcv_SPWPrice1=CheckParentPrices(sp_IDProduct,prdPrice,prdWPrice,1)
                                                                                    
                                        if Cdbl(prdWPrice)>0 and session("customerType")=1 then
                                            prdPrice = prdWPrice
                                        end if
                        
                                        if session("customerCategory")<>0 then
                                            prdPrice=pcv_SPPrice1
                                        else
                                            if (pcv_SPWPrice1>"0") and (session("customerType")=1) then
                                                prdPrice=pcv_SPWPrice1
                                            end if
                                        end if
                                        
                                        IF HaveDiffPrice=1 THEN
                                            prdPrice=pcPriceConf
                                            if Cdbl(pcWPriceConf)>0 and session("customerType")=1 then
                                                prdPrice = pcWPriceConf
                                            end if
                                        END IF
                                    end if
                                    set rsA=nothing
                                    query="SELECT idproduct,description,price,btoBPrice FROM products WHERE pcProd_ParentPrd=" & intTempIdProduct & " AND removed=0 AND ((stock>0) or (nostock>0) or (pcProd_BackOrder>0)) AND pcProd_SPInActive=0;"
                                    set rsA=connTemp.execute(query)
                                    do while not rsA.eof
                                        call CheckDiscount(rsA("idproduct"),0,tmp_qty,prdPrice)
                                        rsA.MoveNext
                                    loop
                                    set rsA=nothing
                                end if
                            End If
                            '// END: Apparel
                            
                            if (Clng(tempPrd) = Clng(intTempIdProduct)) or ((Clng(sp_IDProduct)>0) AND (Clng(tempPrd)=Clng(sp_IDProduct)) AND (pcv_Apparel="1")) then
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
                                pcv_tmpCustomizedPrice=Cdbl(prdPrice)*tempQ-Cdbl(defaultPrice)
                                pExt = " "
                            end if
                            
                            '// START: Apparel
                            if pcv_Apparel="1" then
                                if pBTODisplayType=1 then
                                    CheckAPPStr=CheckAPPStr & "tmp1=document.additem.CAG" & tempVarCat & ".value; tmp2=tmp1.split('_');" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "if (eval(tmp2[0])=='" & intTempIdProduct & "') { alert('" & dictLanguage.Item(Session("language")&"_configPrd_spmsg3") & """" & replace(strDescription,"'","\'") & """'); return(false); }" & vbcrlf
                                else
                                    CheckAPPStr=CheckAPPStr & "totalradio=document.additem.CAG" & tempVarCat & ".length;" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "if (totalradio>0) {" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "for (i=0;i<document.additem.CAG" & tempVarCat & ".length;i++){" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "	if (document.additem.CAG" & tempVarCat & "[i].checked==true) {" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "tmp1=document.additem.CAG" & tempVarCat & "[i].value; tmp2=tmp1.split('_');" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "if (eval(tmp2[0])=='" & intTempIdProduct & "') { alert('" & dictLanguage.Item(Session("language")&"_configPrd_spmsg3") & """" & replace(strDescription,"'","\'") & """'); return(false); }" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "} }" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "} else {" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "	if (document.additem.CAG" & tempVarCat & ".checked==true) {" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "tmp1=document.additem.CAG" & tempVarCat & ".value; tmp2=tmp1.split('_');" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "if (eval(tmp2[0])=='" & intTempIdProduct & "') { alert('" & dictLanguage.Item(Session("language")&"_configPrd_spmsg3") & """" & replace(strDescription,"'","\'") & """'); return(false); }" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "} }" & vbcrlf
                                end if
                            end if
                            '// END: Apparel
                        
                            'Configurator Plus - S
                            if pBTODisplayType<>1 then%>
                                <input type="hidden" name="TXT<%=intTempIdProduct%>" value="<%=ClearHTMLTags2(pcv_tmpArr(7,pcv_tmpN),0)%>">
                            <%end if
                            'Configurator Plus - E
                            
                            
                            '// HAS DEFAULT?
                            if pcv_tmpArr(12,pcv_tmpN)=true then
                                
                                '// DEFAULT ITEM
                                ShowInfoArray = ShowInfoArray & intTempIdProduct& ","
                                
                                '// LOAD DEFAULT SETTINGS
                                If (Clng(tempPrd) = Clng(intTempIdProduct)) or ((Clng(sp_IDProduct)>0) AND (Clng(tempPrd)=Clng(sp_IDProduct)) AND (pcv_Apparel="1")) then                                 
                                    '// DEFAULT ALSO SELECTED ITEM
                                    pcv_strPricingInfo = pExt
                                    strselectvalue = func_DisplayBOMsg
                                    pcv_stringThing = round(prdPrice-prdPrice1,2)
                                    pcv_strIsSelected = "selected"
                                    pcv_strIsChecked = "checked"
                                    pcv_intQtyFieldValue = tempQ
                                else                                 
                                    '// DEFAULT BUT NOT SELECTED
                                    pcv_strPricingInfo = ""
                                    strselectvalue = ""
                                    pcv_stringThing = "0.00"
                                    pcv_strIsSelected = ""
                                    pcv_strIsChecked = ""
                                    pcv_intQtyFieldValue = 0                                    
                                end if 
                                pcv_stringThing2 = prdPrice
                            
                            
                            Else '// if pcv_tmpArr(12,pcv_tmpN)=true then
                             
                                '// NOT DEFAULT ITEM 
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
                                        End if
                                    End if
                                End If
                                                                
                                If scDecSign="," then
                                    prdPrice=replace(prdPrice,",",".")
                                    prdPrice1=replace(prdPrice1,",",".")
                                End If
                                
                                '// LOAD DEFAULT SETTINGS
                                If (Clng(tempPrd) = Clng(intTempIdProduct)) or ((Clng(sp_IDProduct)>0) AND (Clng(tempPrd)=Clng(sp_IDProduct)) AND (pcv_Apparel="1")) then                                 
                                    '// DEFAULT ALSO SELECTED ITEM
                                    pcv_strPricingInfo = ""
                                    strselectvalue = func_DisplayBOMsg
                                    pcv_stringThing = prdPrice
                                    pcv_strIsSelected = "selected"
                                    pcv_strIsChecked = "checked"
                                    pcv_intQtyFieldValue = tempQ
                                else                                 
                                    '// DEFAULT BUT NOT SELECTED
                                    pcv_strPricingInfo = pExt
                                    strselectvalue = ""
                                    pcv_stringThing = prdPrice
                                    pcv_strIsSelected = ""
                                    pcv_strIsChecked = ""
                                    pcv_intQtyFieldValue = 0                                    
                                end if 
                                pcv_stringThing2 = prdPrice1
                        
                            end if '// if pcv_tmpArr(12,pcv_tmpN)=true then 
                            
                            '// CLASS CONFIGURATION: RADIO (RC)
                            If (displayQF=True) Then 
                                pcv_strColumn1 = "col-xxs-12 col-xs-6 col-sm-2"
                                pcv_strColumn2 = "col-xxs-12 col-xs-6 col-sm-2"
                                pcv_strColumn3 = "col-xxs-12 col-xs-12 col-sm-6"
                            Else 
                                pcv_strColumn1 = "col-xxs-12 col-xs-6 col-sm-1"
                                pcv_strColumn2 = "col-xxs-12 col-xs-6 col-sm-3"
                                pcv_strColumn3 = "col-xxs-12 col-xs-12 col-sm-6"
                            End If 
                            %>                            
                        
                            <% if pBTODisplayType=1 then %>
                            
                                <div class="col-xxs-12 col-xs-12 col-sm-9"> 
                                    <option value="<%if sp_IDProduct>"0" then%><%=sp_IDProduct%><%else%><%=intTempIdProduct%><%end if%>_<%=pcv_stringThing%>_<%=weight%>_<%=pcv_stringThing2%>_<%=intTempIdProduct%>" <%=pcv_strIsSelected%>><%if sp_IDProduct>"0" then%><%=sp_PrdName & pcv_strPricingInfo%><%else%><%=strDescription & pcv_strPricingInfo%><%end if%></option>
                                    <% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
                        
                                    HiddenFields=HiddenFields & "<input type=hidden name=""CAG" & tempVarCat & intTempIdProduct & "HF"" value=""" & pcv_qtyValid & "_" & pcv_minQty & """>" & vbcrlf %>
                                
                            <% else '// if pBTODisplayType=1 then %>
                        
                                <div class="<%=pcv_strColumn1%>">
                        
                                    <% '// Row 1: Radio %>
                                    <input type="radio" name="CAG<%=tempVarCat%>" value="<%if sp_IDProduct>"0" then%><%=sp_IDProduct%><%else%><%=intTempIdProduct%><%end if%>_<%=pcv_stringThing%>_<%=weight%>_<%=pcv_stringThing2%>_<%=intTempIdProduct%>" onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; CheckPreValue(this, 2, 0);" <%=pcv_strIsChecked%> class="clearBorder">
                                    
                                    <% '// Row 1: Quantity %>
                                    <% if (displayQF=True) then %>
                                        <input class="form-control quantity" type="text" size="2" id="CAG<%=tempVarCat%>QF<%=icount-1%>" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="<%=pcv_intQtyFieldValue%>" onblur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>,<%=pcv_multiQty%><%if (pcv_Apparel="0") AND pNostock=0 AND pcv_intBackOrder=0 AND scOutofstockpurchase=-1 AND iBTOOutofstockpurchase=-1 then%><%strQtyCheck=strQtyCheck & vbcrlf & "if (!(qttverify(document.getElementById('" & "CAG" & tempVarCat & "QF" & icount-1 & "')," & pcv_qtyvalid &"," & pcv_minQty & "," & pcv_multiQty & "," & pstock & ",1))) {setTimeout(function() {fname.focus();}, 0); return(false);}" & vbcrlf%>,<%=pstock%><%end if%>)) calculate(document.additem.CAG<%=tempVarCat%>,2);">
                                    <% else %>
                                        <input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="<%=pcv_intQtyFieldValue%>">
                                    <% end if %>
                                        
                                </div>
                                
                                <div class="<%=pcv_strColumn2%>">
                                    <% '// Row 2: Image %>
                                    <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                        <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                    <% end if %>
                                </div>
                                
                                <div class="col-xxs-12 col-xs-12 col-sm-6">
                                
                                    <% '// Row 3: Details %>
                                    <span name="CAG<%=tempVarCat%>DESC<%=icount-1%>"><%if sp_IDProduct>"0" then%><%=sp_PrdName%><%else%><%=strDescription%><%end if%></span>
                        
                                    <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%=pcv_strPricingInfo%>" readonly class="transparentField" size="<%=len(pExt)%>">
                                    <% if not pClngShowSku = 0 then %>
                                        <div class="pcSmallText"><%=strSku%></div>
                                    <% end if %>
									
									<% if pcv_Apparel=1 then %>
										<div class="pcSmallText" id="show_CAG<%=tempVarCat&"P"&intTempIdProduct%>">
										<%if pcv_tmpArr(12,pcv_tmpN)=true then%>
										<a href="javascript:win1('<%=pcf_GeneratePopupPath()%>/popup_Apparel.asp?IDBTO=<%=pIdProduct%>&IDField=CAG<%=tempVarCat%>&IDProduct=<%=intTempIdProduct%>&vindex=<%=icount-1%>&Price=<%=prdPrice1%>&AddPrice=0',document.additem.CAG<%=tempVarCat%>,2)"><%=dictLanguage.Item(Session("language")&"_configPrd_spmsg1")%></a>
										<%else%>
										<a href="javascript:win1('<%=pcf_GeneratePopupPath()%>/popup_Apparel.asp?IDBTO=<%=pIdProduct%>&IDField=CAG<%=tempVarCat%>&IDProduct=<%=intTempIdProduct%>&vindex=<%=icount-1%>&Price=<%=prdPrice1%>&AddPrice=<%=prdPrice%>',document.additem.CAG<%=tempVarCat%>,2)"><%=dictLanguage.Item(Session("language")&"_configPrd_spmsg1")%></a>
										<%end if%>
										</div>
									<% end if %>
                        
                                    <%=func_DisplayBOMsg1(tempVarCat,intTempIdProduct)%>
                              
                        
                                    <% if pcv_ShowDesc="1" then %>
                                        <div class="row">
                                            <div class="col-xs-12">
                                                <span class="configDesc">
                                                    <%=pcv_prdSDesc%>
                                                </span>
                                            </div>
                                        </div>
                                    <% end if %>
                                
                            <% end if '// if pBTODisplayType=1 then %>
                        
                            <%
                            '// START: Apparel
                            If statusAPP="1" Then
                                if pBTODisplayType=1 then
                                    if pcv_Apparel=1 then
                                        if cdefault=true then
                                            prdPriceA=0
                                        else
                                            prdPriceA=prdPrice
                                        end if
                                        myApparel=myApparel & "if (j==" & (icount-1) & ")" & vbcrlf
                                        myApparel=myApparel & "{" & vbcrlf
                                        myApparel=myApparel & "document.additem.app_IDProduct_CAG" & tempVarCat & ".value=" & intTempIdProduct & ";" & vbcrlf
                                        myApparel=myApparel & "document.additem.app_Price_CAG" & tempVarCat & ".value=" & prdPrice1 & ";" & vbcrlf
                                        myApparel=myApparel & "document.additem.app_AddPrice_CAG" & tempVarCat & ".value=" & prdPriceA & ";" & vbcrlf
                                        myApparel=myApparel & "document.additem.app_VIndex_CAG" & tempVarCat & ".value=" & (icount-1) & ";" & vbcrlf
                                        myApparel=myApparel & "m=1;" & vbcrlf
                                        myApparel=myApparel & "display_CAG" & tempVarCat & "();" & vbcrlf
                                        myApparel=myApparel & "return(true);" & vbcrlf
                                        myApparel=myApparel & "break;" & vbcrlf
                                        myApparel=myApparel & "}" & vbcrlf
                                    end if
                                end if
                            End If
                            '// END: Apparel
                            %>
                        
                            <% IF pBTODisplayType<>1 THEN %>                                
                                    </div>                            
                                    <div class="col-xxs-12 col-xs-12 col-sm-2">
                                        
                                        <% if showInfoVar = "1" then %>
                                            <% if iBTODetLinkType=1 then%>	
                                                    <a class="" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&cd=<%=replace(strCategoryDesc,"""","%22")%>')"><%=pcv_strBTODetTxt %></a>
                                            <% else %>
                                                    <a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&cd=<%=replace(strCategoryDesc,"""","%22")%>')">
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
                                                if MyTest=1 then%>
                                                    <a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=ProductArray(i)%>')"><img alt="<%=dictLanguage.Item(Session("language")&"_viewPrd_16")%>" src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>"></a>
                                                <%end if
                                            end if
                                        next
                                        'End Show Option Discounts icon
                                        %>
                                        
                                    </div>
                                </div>
                            <% END IF %>
                        
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
                                <div <%=strCol%>><div class="col-xxs-12 col-xs-12 col-sm-9">
                            <%end if
                            if Cdbl(varTempDefaultPrice)<0 then 
                                if tempPrd=0 then %>
                                    <%pcv_tmpCustomizedPrice=varTempDefaultPrice
                                    if pBTODisplayType=1 then
                                    icount=icount+1%>
                                        <option value="0_<%=varTempDefaultPrice%>_0_0_0" selected><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%><%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(varTempDefaultPrice)%><%end if%></option>
                                        <% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
                                           strselectvalue = func_DisplayBOMsg
                                        %>
                                    <% else
                                    icount=icount+1 %>
                                            <input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0_0" checked onClick="CheckPreValue(this, 2, 0);" class="clearBorder">
                                            
                                            <input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                            <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
                                            <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*varTempDefaultPrice)%><%end if%>" readonly class="transparentField" size="<%=len(" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*varTempDefaultPrice))%>">
                                        <% end if %>
                                    <% else %>
                                        <% if pBTODisplayType=1 then
                                        icount=icount+1%>
                                            <option value="0_<%=varTempDefaultPrice%>_0_0_0"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%><%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(varTempDefaultPrice)%><%end if%></option>
                                           <% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
                                         else
                                        icount=icount+1 %>
                                            <input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0_0"  onClick="CheckPreValue(this, 2, 0);" class="clearBorder">
                                            
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
                                        <option value="0_<%=varTempDefaultPrice%>_0_0_0" selected><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%><%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%></option>
                                       <% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf 
                                          strselectvalue = func_DisplayBOMsg
                                    else
                                    icount=icount+1 %>
                                        <input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0_0" checked onClick="CheckPreValue(this, 2, 0);" class="clearBorder">
                                        <input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                        <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
                                        <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%>" readonly class="transparentField" size="<%=len(" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice))%>">
                                    <% end if %>
                                <% else %>
                                    <% if pBTODisplayType=1 then
                                    icount=icount+1%>
                                        <option value="0_<%=varTempDefaultPrice%>_0_0_0"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%><%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%></option>
                                        <% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='' ;"  &vbcrlf %>
                                    <% else
                                    icount=icount+1 %>
                                        <input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0_0" onClick="CheckPreValue(this, 2, 0);" class="clearBorder">
                                        
                                        <input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                        <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
                                        <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%>" readonly class="transparentField" size="<%=len(" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice))%>">
                                    <% end if %>
    
                            <% end if %>
                            <% else if cdVar="0" then 
                                if tempPrd=0 then %>
                                    <% if pBTODisplayType=1 then
                                    icount=icount+1 %>
                                        <option value="0_0_0_0_0" selected><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></option>
                                          <%  StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='' ;"  &vbcrlf 
                                          %>
                                    <% else
                                    icount=icount+1 %>
                                    <input type="radio" name="CAG<%=tempVarCat%>" value="0_0_0_0_0" checked onClick="CheckPreValue(this, 2, 0);" class="clearBorder">
                                    
                                    <input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                    <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
                                    <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="" readonly class="transparentField">
                                
                                    <% end if %>
                                <% else %>
                                    <% if pBTODisplayType=1 then
                                    icount=icount+1 %>
                                        <option value="0_0_0_0_0"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></option>
                                              <%StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='' ;"  &vbcrlf 
                                       else
                                            icount=icount+1 %>
                                            <input type="radio" name="CAG<%=tempVarCat%>" value="0_0_0_0_0" onClick="CheckPreValue(this, 2, 0);" class="clearBorder">
                                            
                                            <input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                            <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
                                            <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="" readonly class="transparentField">
                                        
                                            <% end if %>
                            <% end if %>
                            <% end if %>
                            <% end if
                            end if %>
                            <%if pBTODisplayType<>1 then%>
                            </div></div>
                            <%end if%>
                            <% end if %>
                            <% if pBTODisplayType=1 then %> 
                            </select>
    
                            <%=HiddenFields%>
                            <script type=text/javascript>
                                 <%=StrBackOrd %>
                                 function showAvail<%=tempVarCat%>(sel){
                                 document.getElementById("AV<%=tempVarCat%>").innerHTML = availArr<%=tempVarCat%>[sel.selectedIndex] + "" ;															 
                                 }
    
                                var ns6=document.getElementById&&!document.all
                                var ie=document.all
                                function display_CAG<%=tempVarCat%>()
                                {
                                    document.getElementById("show_CAG<%=tempVarCat%>").style.display="";
                                }
                                function hide_CAG<%=tempVarCat%>()
                                {
                                    document.getElementById("show_CAG<%=tempVarCat%>").style.display="none";
                                }
                                            
                                <%funcTestCat=funcTestCat & "testCAG" & tempVarCat & "();" & vbcrlf%>
                                function testCAG<%=tempVarCat%>()
                                {
                                    var oSelect=eval("document.additem.CAG<%=tempVarCat%>");
                                    var j=0;
                                    for (j=0;j<oSelect.options.length;j++)
                                    {
                                        if (oSelect.value==oSelect.options[j].value)
                                        {
                                            <%=myApparel%>
                                        }
                                    }
                                    <%if myApparel<>"" then%>
                                        hide_CAG<%=tempVarCat%>();
                                    <%end if%>
                                }
    
                            </script>
    
                            <span  id="AV<%=tempVarCat%>" ><%=strselectvalue%></span>
                            <%intOpCnt = intOpCnt + 1 %>
                            <% end if %>
                            
                            <%pcv_CustomizedPrice=pcv_CustomizedPrice+pcv_tmpCustomizedPrice%>
                            <%IF pBTODisplayType<>1 THEN%>
                        <!--<div <%=strCol%>> -->
                            <%END IF%>
                            <input name="currentValue<%=jCnt%>" type="HIDDEN" value="<%=pcv_tmpCustomizedPrice%>">
                            <input name="CAT<%=jCnt%>" type="HIDDEN" value="CAG<%=tempVarCat%>">
                            <input name="Discount<%=jCnt%>" type="HIDDEN" value="<%=Round(pcv_tmpIDiscount+0.001,2)%>">
                            <%IF pBTODisplayType<>1 THEN%>
                        <!--</div> -->
                            <%END IF%>
                            <%IF pBTODisplayType=1 THEN %>
                        </div>
                        <div class="col-xxs-12 col-xs-12 col-sm-2">
                            <% if pcv_HaveApparel=1 then %>
                                <div id="show_CAG<%=tempVarCat%>" style="display:none;">
                                    <a href="javascript:win1('<%=pcf_GeneratePopupPath()%>/popup_Apparel.asp?IDBTO=<%=pIdProduct%>&IDField=CAG<%=tempVarCat%>&IDProduct='+document.additem.app_IDProduct_CAG<%=tempVarCat%>.value+'&vindex='+document.additem.app_VIndex_CAG<%=tempVarCat%>.value+'&Price='+document.additem.app_Price_CAG<%=tempVarCat%>.value+'&AddPrice='+document.additem.app_AddPrice_CAG<%=tempVarCat%>.value,document.additem.CAG<%=tempVarCat%>,0)"><%=dictLanguage.Item(Session("language")&"_configPrd_spmsg1")%></a>
                                </div>
                            <% end if %>
                            <% if showInfoVar = "1" then %>
                                <% if iBTODetLinkType=1 then %>
                                    <a class="" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&amp;cd=<%=replace(strCategoryDesc,"""","%22")%>')"><%=pcv_strBTODetTxt %></a>
                                <% else %>
                                    <a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&amp;cd=<%=replace(strCategoryDesc,"""","%22")%>')">
                                        <span class="pcButtonText"><%=pcv_strBTODetTxt %></span>
                                        <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>">
                                    </a>
                                <% end if %>
                            <% end if %>
                        
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
                        <a href="javascript:openbrowser('<%=pcv_sffolder%>OptpriceBreaks.asp?type=<%=Session("customerType")%>&SIArray=<%=ShowInfoArray%>&cd=<%=Server.URLEnCode(replace(strCategoryDesc,"""","%22"))%>')">
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
                                        CATNotes=pcv_tmpArr(30,pcv_tmpN)
                                        if CATNotes<>"" then
                                        %>
                                        <div <%=strCol%>>  
                                            <div class="col-xs-12"><span class="catNotes"><%=CATNotes%></span></div>
                                        </div>
                                        <%
                                        end if
                                        %>
                                    <%'BTOCM-S%>
                                    <div class="row">
                                        <span name="CMMsg<%=pcv_tmpArr(0,pcv_tmpN)%>" id="CMMsg<%=pcv_tmpArr(0,pcv_tmpN)%>"></span>
                                    </div>
                                    <%'BTOCM-E%>
                            <% PrdCnt = 0 %>
                            
                            <% 
                            ShowInfoArray = ""
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
                                
                                        if pcv_tmpArr(28,pcv_tmpN)=1 then
                                            'Get apparel and compare
                                            query="SELECT products.idProduct FROM products WHERE (((products.pcprod_ParentPrd)="&intTempIdProduct&"));"
                                            set rsGetApp=server.CreateObject("ADODB.RecordSet")
                                            set rsGetApp=conntemp.execute(query)
                                            do until rsGetApp.eof or SelectVar="1"
                                                intTempGetAppId=rsGetApp(0)
                                                if Clng(ArrProduct(i))=Clng(intTempGetAppId) then
                                                    if Clng(ArrCategory(i))=Clng(checkVarCat) then
                                                        SelectVar="1"
                                                        tempQ=ArrQuantity(i)
                                                    end if
                                                end if
                                                rsGetApp.moveNext
                                            loop
                                            set rsGetApp=nothing
                                        else
                                            if Clng(ArrProduct(i))=Clng(intTempIdProduct) then
                                                if Clng(ArrCategory(i))=Clng(checkVarCat) then
                                                    SelectVar="1"
                                                    tempQ=ArrQuantity(i)
                                                end if
                                            end if
                                        end if
                                    
                                    next
                                                        
                                    pcv_prdDesc=pcv_tmpArr(29,pcv_tmpN)
                                    pcv_prdSDesc=pcv_tmpArr(30,pcv_tmpN)
                                    if IsNull(pcv_prdSDesc) or trim(pcv_prdSDesc)="" then
                                        pcv_prdSDesc=pcv_prdDesc
                                    end if
                                
                                    pcv_Apparel=pcv_tmpArr(28,pcv_tmpN)
                                    if IsNull(pcv_Apparel) or pcv_Apparel="" then
                                        pcv_Apparel=0
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
                                    pcv_multiQty=pcv_tmpArr(27,pcv_tmpN)
                                    if isNULL(pcv_multiQty) OR pcv_multiQty="" then
                                        pcv_multiQty="0"
                                    end if
                                                                        
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
    
                                prdPriceA=prdPrice
                                
                                sp_IDProduct=0
                                sp_PrdName=""
                                if pcv_Apparel="1" then
									HaveDiffPrice=0
                                    queryQ="SELECT price,btoBPrice FROM Products WHERE idProduct=" & intTempIdProduct & ";"
                                    set rsQ=connTemp.execute(queryQ)
                                    if not rsQ.eof then
                                        pcParentPrice=rsQ("price")
                                        pcParentWPrice=rsQ("btoBPrice")
                                        set rsQ=nothing
                                        queryQ="SELECT price,Wprice FROM configSpec_Products WHERE specProduct=" & pIDProduct & " AND configProduct=" & intTempIdProduct  & " AND configProductCategory=" & tempVarCat & ";"
                                        set rsQ=connTemp.execute(queryQ)
                                        if not rsQ.eof then
                                            pcPriceConf=rsQ("price")
                                            pcWPriceConf=rsQ("Wprice")
                                            set rsQ=nothing
                                            
                                            if ccur(pcPriceConf)<>ccur(pcParentPrice) then
                                                HaveDiffPrice=1
                                            else
                                                if (ccur(pcWPriceConf)<>ccur(pcParentWPrice)) AND (ccur(pcWPriceConf)<>ccur(pcParentPrice)) then
                                                    HaveDiffPrice=1
                                                end if
                                            end if
                                        end if
                                        set rsQ=nothing
                                    end if
                                    set rsQ=nothing
									
                                    query="SELECT idproduct,description,price,btoBPrice,stock,nostock,pcProd_BackOrder,pcProd_ShipNDays FROM products WHERE pcProd_ParentPrd=" & intTempIdProduct & " AND removed=0 AND ((stock>0) or (nostock>0) or (pcProd_BackOrder>0)) AND pcProd_SPInActive=0;" 
                                    set rsA=connTemp.execute(query)
                                    if not rsA.eof then
                                        pcA=rsA.getRows()
                                        intCount=ubound(pcA,2)
                                        for i = lbound(ArrProduct) to (UBound(ArrProduct)-1)
                                            for j = 0 to intCount
                                                if (Clng(ArrProduct(i))=Clng(pcA(0,j))) AND (Clng(ArrCategory(i))=Clng(checkVarCat)) then
                                                    SelectVar="1"
                                                    sp_IDProduct=pcA(0,j)
                                                    sp_PrdName=pcA(1,j)
                                                    sp_PrdName=replace(sp_PrdName,"""","&quot;")
                                                    prdPrice = Cdbl(pcA(2,j))
                                                    prdWPrice=pcA(3,j)
                                                    if IsNull(prdWPrice) or prdWPrice="" then
                                                        prdWPrice=0
                                                    end if
                                                    if Cdbl(prdWPrice)=0 then
                                                        prdWPrice=prdPrice
                                                    end if
                                                    pstock=pcA(4,j)
                                                    pNostock=pcA(5,j)	
                                                    if pNostock = "" or pNoStock = null then
                                                        pNostock = 0
                                                    end if						
                                                    pcv_intBackOrder = pcA(6,j)							
                                                    pcv_intShipNDays = pcA(7,j)
                                                                                        
                                                    pcv_SPPrice1=CheckParentPrices(sp_IDProduct,prdPrice,prdWPrice,0)
                                                    pcv_SPWPrice1=CheckParentPrices(sp_IDProduct,prdPrice,prdWPrice,1)
                                                                                        
                                                    if Cdbl(prdWPrice)>0 and session("customerType")=1 then
                                                        prdPrice = prdWPrice
                                                    end if
                
                                                    if session("customerCategory")<>0 then
                                                        prdPrice=pcv_SPPrice1
                                                    else
                                                        if (pcv_SPWPrice1>"0") and (session("customerType")=1) then
                                                            prdPrice=pcv_SPWPrice1
                                                        end if
                                                    end if
													
													if HaveDiffPrice=1 then
														prdPrice=pcPriceConf
														if Cdbl(pcWPriceConf)>0 and session("customerType")=1 then
															prdPrice = pcWPriceConf
														end if
													end if
                                                    tempQ=ArrQuantity(i)
                                                end if
                                            next
                                        next
                                    end if
                                    set rsA=nothing
                                    query="SELECT idproduct,description,price,btoBPrice FROM products WHERE pcProd_ParentPrd=" & intTempIdProduct & " AND removed=0 AND ((stock>0) or (nostock>0) or (pcProd_BackOrder>0)) AND pcProd_SPInActive=0;"
                                    set rsA=connTemp.execute(query)
                                    do while not rsA.eof
                                        call CheckDiscount(rsA("idproduct"),0,tmp_qty,prdPrice)
                                        rsA.MoveNext
                                    loop
                                    set rsA=nothing
                                end if
                            
                                pcv_tmpIDiscount=0
                                call CheckDiscount(intTempIdProduct,tmp_selected,tmp_qty,prdPrice)
                                
                                pcv_tmpCustomizedPrice=0
                                if tmp_selected then
                                    if cdefault<>"" and cdefault<>0 then
                                        pcv_tmpCustomizedPrice=(tempQ-pcv_minqty)*prdPrice+(prdPrice-prdPriceA)
                                    else
                                        pcv_tmpCustomizedPrice=tempQ*prdPrice
                                    end if
                                    pcv_CustomizedPrice=pcv_CustomizedPrice+pcv_tmpCustomizedPrice
                                else
                                    if cdefault<>"" and cdefault<>0 then
                                        pcv_tmpCustomizedPrice=cdbl(-pcv_minqty*prdPrice)
                                        pcv_CustomizedPrice=pcv_CustomizedPrice+pcv_tmpCustomizedPrice
                                    end if
                                end if
    
                            ShowInfoArray = ShowInfoArray & intTempIdProduct& ","
                            ShowInfoArray = intTempIdProduct& ","   
                            PrdCnt = PrdCnt + 1
                            jCnt = jCnt + 1%>
                            <input name="MS<%=jCnt%>" type="HIDDEN" value="<%=VarMS%>">
                            <input name="currentValue<%=jCnt%>" type="HIDDEN" value="<%=pcv_tmpCustomizedPrice%>">
                            <input name="Discount<%=jCnt%>" type="HIDDEN" value="<%=Round(pcv_tmpIDiscount+0.001,2)%>">
                            <input name="CAT<%=jCnt%>" type="HIDDEN" value="CAG<%=tempCcat%>">

                            <%
                            '// CLASS CONFIGURATION: CHECKBOX (RC)
                            If (displayQF=True) Then
                                pcv_strColumn1 = "col-xxs-12 col-xs-6 col-sm-2"
                                pcv_strColumn2 = "col-xxs-12 col-xs-6 col-sm-2"
                                pcv_strColumn3 = "col-xxs-12 col-xs-12 col-sm-6"
                            Else
                                pcv_strColumn1 = "col-xxs-12 col-xs-6 col-sm-2"
                                pcv_strColumn2 = "col-xxs-12 col-xs-6 col-sm-2"
                                pcv_strColumn3 = "col-xxs-12 col-xs-12 col-sm-6"
                            End If                            
                            %>                            
                            <div <%=strCol%>>
                                <div class="<%=pcv_strColumn1%>">
                                    
                                    <% '// Row 1: checkbox %>
                                    <%'Configurator Plus - S %>
                                        <input type="hidden" name="TXT<%=intTempIdProduct%>" value="<%=ClearHTMLTags2(pcv_tmpArr(7,pcv_tmpN),0)%>">
                                    <%'Configurator Plus - E %>
                                    
                                    <input type="hidden" name="Cat<%=intTempIdCategory%>_Prd<%=PrdCnt%>" value="<%=intTempIdProduct%>">
                                    <% If SelectVar="1" then %>
                                        <input type="checkbox" name="CAG<%=intTempIdCategory&intTempIdProduct%>" value="<%if sp_IDProduct>"0" then%><%=sp_IDProduct%><%else%><%=intTempIdProduct%><%end if%>_<%if cdefault<>0 then%><%=Round((prdPrice-prdPriceA)+0.001,2)%><%else%><%=prdPrice%><%end if%>_<%=weight%>_<%=prdPrice%>_<%=intTempIdProduct%>" onClick="javscript:document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>QF.value='<%=pcv_minQty%>'; CheckBoxPreValue(this,0);" checked class="clearBorder">
                                    <% else %>
                                        <input type="checkbox" name="CAG<%=intTempIdCategory&intTempIdProduct%>" value="<%if sp_IDProduct>"0" then%><%=sp_IDProduct%><%else%><%=intTempIdProduct%><%end if%>_<%if cdefault<>0 then%>0<%else%><%=prdPrice%><%end if%>_<%=weight%>_<%=prdPrice%>_<%=intTempIdProduct%>" onClick="javscript:document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>QF.value='<%=pcv_minQty%>'; CheckBoxPreValue(this,0);" class="clearBorder">
                                    <% end if %>
                                    
                                <% '// Row 1: Quantity %>
                                <%
                                SelectVar="0"
                                RTestStr=RTestStr & vbcrlf & "if (document.additem.CAG"& intTempIdCategory & intTempIdProduct & ".checked !=false) { RTest" & CB_CatCnt & "=" & "RTest" & CB_CatCnt & "+document.additem.CAG" & intTempIdCategory & intTempIdProduct & ".checked; }"& vbcrlf
                                %>
                                <% if pcv_Apparel="1" then
                                    CheckAPPStr=CheckAPPStr & "if (document.additem.CAG" & intTempIdCategory&intTempIdProduct & ".checked==true) {" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "tmp1=document.additem.CAG" & intTempIdCategory&intTempIdProduct & ".value; tmp2=tmp1.split('_');" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "if (eval(tmp2[0])=='" & intTempIdProduct & "') { alert('" & dictLanguage.Item(Session("language")&"_configPrd_spmsg3") & """" & replace(strDescription,"'","\'") & """'); return(false); }" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "}" & vbcrlf
                                end if %>
                                    <%if (displayQF=True) then%>
                                        <input class="form-control quantity" type="text" size="2" id="CAG<%=intTempIdCategory&intTempIdProduct%>QF" name="CAG<%=intTempIdCategory&intTempIdProduct%>QF" value="<%=tempQ%>" onblur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>,<%=pcv_multiQty%><%if (pcv_Apparel="0") AND pNostock=0 AND pcv_intBackOrder=0 AND scOutofstockpurchase=-1 AND iBTOOutofstockpurchase=-1 then%><%strQtyCheck=strQtyCheck & vbcrlf & "if (!(qttverify(document.getElementById('" & "CAG" & intTempIdCategory&intTempIdProduct & "QF" & "')," & pcv_qtyvalid &"," & pcv_minQty & "," & pcv_multiQty & "," & pstock & ",1))) {setTimeout(function() {fname.focus();}, 0); return(false);}" & vbcrlf%>,<%=pstock%><%end if%>)) calculate(document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>,0);">
                                    <%else%>
                                        <input type="hidden" name="CAG<%=intTempIdCategory&intTempIdProduct%>QF" value="<%=tempQ%>">
                                    <%end if%>
                                </div>    
    
                                <div class="<%=pcv_strColumn2%>">
                                    <% '// Row 2: Image %>
                                    <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                        <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                    <% end if %>
                                </div>	
                                
                                <%if (displayQF=True) then%><%end if%>						
                                <div class="col-xxs-12 col-xs-12 col-sm-6">
                                
                                    <% '// Row 4: Details %>
                                    <span name="CAG<%=intTempIdCategory&intTempIdProduct%>DESC0"><%if sp_IDProduct>"0" then%><%=sp_PrdName%><%else%><%=strDescription%><%end if%></span>
                            
                                    <%if pnoprices<2 then%>
                                        &nbsp;-&nbsp;<input name="CAG<%=intTempIdCategory&intTempIdProduct%>TX0" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%=scCurSign & money(prdPrice)%>" readonly class="transparentField" size="<%=len("" & scCurSign & money(prdPrice))%>">
                                    <%end if%>

                                    <% if not pClngShowSku = 0 then %>
                                        <div class="pcSmallText"><%=strSku%></div>
                                    <% end if %>
									
									<%if pcv_Apparel=1 then%>
                                        <div class="pcSmallText" id="show_CAG<%=intTempIdCategory&intTempIdProduct&"P"&intTempIdProduct%>"><a href="javascript:win1('<%=pcf_GeneratePopupPath()%>/popup_Apparel.asp?IDPROD=<%=intTempIdProduct%>&IDBTO=<%=pIdProduct%>&IDField=CAG<%=intTempIdCategory&intTempIdProduct%>&IDProduct=<%=intTempIdProduct%>&Price=<%=prdPrice%>&AddPrice=<%if (cdefault<>"") and (cdefault<>0) then%>0<%else%><%=prdPrice%><%end if%>&vindex=0',document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>,0)"><%=dictLanguage.Item(Session("language")&"_configPrd_spmsg1")%></a></div>
                                    <%end if%>

                                    <%if pcv_ShowDesc="1" then%>
                                        <div class="configDesc"><%=pcv_prdSDesc%></div>
                                    <%end if%>
                                    <%if (displayQF=True) then%><%end if%>
                                    
                                </div>
                                
                                <div class="col-xxs-12 col-xs-12 col-sm-2">
    
                                    <% if showInfoVar = "1" then %>
                                    
                                        <% if iBTODetLinkType=1 then %>
                                            <a class="" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&cd=<%=replace(strCategoryDesc,"""","%22")%>')"><%=pcv_strBTODetTxt %></a>
                                        <% else %>
                                            <a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&cd=<%=replace(strCategoryDesc,"""","%22")%>')">
                                                <span class="pcButtonText"><%=pcv_strBTODetTxt %></span>
                                                <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>">
                                            </a>
                                        <% end if %>
                                    
                                    <% end if %>
                                 
                                <%
                                'Show Option Discounts icon
                                ProductArray = Split(ShowInfoArray,",")
                                for i = lbound(ProductArray) to (UBound(ProductArray)-1)
                                    if ProductArray(i)<>"" then
                                        MyTest=CheckOptDiscount(ProductArray(i))
                                        if MyTest=1 then
                                            if pnoprices<2 then%>
                                                <a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=ProductArray(i)%>')"><img alt="<%=dictLanguage.Item(Session("language")&"_viewPrd_16")%>" src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>"  align="middle"></a>
                                            <%end if
                                        end if
                                    end if
                                next
                                'End Show Option Discounts icon
                                %> 
                                </div>
                            </div>
                            
                            <%'---- Check Boxes ---%>
                            <% if func_DisplayBOMsg <> "" then %>
                                <div <%=strCol%> style="vertical-align:top">
                                <%=func_DisplayBOMsg1(tempVarCat,intTempIdProduct)%>
                                </div>
                            <% end if %>
                            
                            <%if pcv_ShowDesc="1" then%>
                            <!--
                            <div <%=strCol%> style="vertical-align:top">
                                <%if (displayQF=True) then%><%end if%>
                                <div class="col-xxs-12 col-xs-12 col-sm-4"></div>
                                <div class="col-xxs-12 col-xs-12 col-sm-6">
                                    <span class="configDesc">
                                        <%=pcv_prdSDesc%>
                                    </span>
                                </div>
                            </div>
                            -->
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
        End if //'Have Configurator Categories
        set rsSSobj=nothing
        '******* END Configurator Categories
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
        response.write "return(eval(roundNumber(DisValue1,2)));" & VBCrlf
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
        response.write "return(eval(roundNumber(DisValue1,2)));" & VBCrlf
        response.write " } </script>" & VBCRlf
        %>
        <script type=text/javascript>
        function roundNumber(num, dec) {
        var tmp1 = Math.round(num*Math.pow(10,dec))/Math.pow(10,dec);
        return(tmp1);
        }
    
        function chkR()
        {
        <%
        'Configurator Plus - S
        if pcv_HaveRules=1 then%>
        var tmp1=CheckCatBeforeSubmit();
        if (tmp1=="no")
        {
            return(false);
        }
        <%end if
        'Configurator Plus - E 
        %>
        <%if ReqTestStr<>"" then%>
        <%=ReqTestStr%>
        <%end if%>
        <%=CheckAPPStr%>
        if (checkproqty(document.additem.quantity))
        {
            <%
            'Configurator Plus - S
            if pcv_HaveRules=1 then%>
            var i=0;
            var objElems = document.additem.elements;
            var j=objElems.length;
            do
            {
                i=j-1;
                objElems[i].disabled=false;
            }
            while (--j);
            <%end if
            'Configurator Plus - E 
            %>
            return CheckTotalItemQty();
        }
        else
        {
            return(false);
        }
        }
        //APP-S
        <%=funcTestCat%>
        //APP-E
        </script>
        <%Call pcs_GetDefaultBTOItemsMin%>
        <input type="hidden" name="FirstCnt" value="<%=jCnt%>">
        <input type="hidden" name="CB_CatCnt" value="<%=CB_CatCnt%>">
    </div>
    <%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  product configuration table - Reconfigure
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  product configuration table
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_BTOConfigTable
    Dim query,rsSSObj,tmpquery
    
	tmpquery=""
	if (scOutOfStockPurchase="-1") AND (iBTOOutofStockPurchase="-1") then
		tmpquery=" AND ((products.stock>0) OR (products.nostock<>0) OR (products.pcProd_BackOrder<>0))"
	end if

	call CreateAppPopUp()

	query="SELECT categories.idCategory, categories.categoryDesc, configSpec_products.multiSelect,products.pcprod_qtyvalidate,products.pcprod_minimumqty,products.idproduct, products.weight, products.description, configSpec_products.prdSort, configSpec_products.price, configSpec_products.Wprice, configSpec_products.showInfo, configSpec_products.cdefault, configSpec_products.requiredCategory, configSpec_products.displayQF,configSpec_products.pcConfPro_ShowDesc,configSpec_products.pcConfPro_ShowImg,configSpec_products.pcConfPro_ImgWidth,configSpec_products.pcConfPro_ShowSKU,products.sku,products.smallImageUrl,products.stock,products.noStock, products.pcProd_BackOrder, products.pcProd_ShipNDays,products.pcprod_minimumqty,configSpec_Products.pcConfPro_UseRadio,products.pcProd_multiQty,products.pcProd_Apparel,products.details,products.sDesc,configSpec_products.Notes FROM categories INNER JOIN (products INNER JOIN configSpec_products ON (products.idproduct=configSpec_products.configProduct AND products.active<>0 AND products.removed=0" & tmpquery & ")) ON categories.idCategory = configSpec_products.configProductCategory WHERE configSpec_products.specProduct="&pIdProduct&" ORDER BY configSpec_products.catSort, categories.idCategory, configSpec_products.prdSort,products.description;"
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
	<%
	funcTestCat=""
	%>
	<div class="">
        <% 
        CB_CatCnt = 0
        jcnt=0
    
        '*******************************************
        '******* START Configurator Categories
    
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
    
                        pcv_HaveApparel=0
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
                        query="SELECT configSpec_products.configProduct,configSpec_products.price, configSpec_products.Wprice, configSpec_products.cdefault FROM configSpec_products WHERE configSpec_products.configProductCategory="&tempVarCat&" AND configSpec_products.specProduct="&pIdProduct&" AND configSpec_products.cdefault<>0;"
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
                
                        <div class="panel-heading">
                            <%=pcv_tmpArr(1,pcv_tmpN)%>
                        </div>
                        <div class="panel-body">
                        
                        <%
                        '// If there are configuration instructions for this category, show them here.
                        CATNotes=pcv_tmpArr(31,pcv_tmpN)
                        if CATNotes<>"" then
                        %>
                        <div <%=strCol%>>
                            <div class="col-xs-12"><span class="catNotes"><%=CATNotes%></span></div>
                        </div>
                        <% end if %>
                        
                        <%'BTOCM-S%>
                        <div class="row">
                            <span name="CMMsg<%=pcv_tmpArr(0,pcv_tmpN)%>" id="CMMsg<%=pcv_tmpArr(0,pcv_tmpN)%>"></span>
                        </div>
                        <%'BTOCM-E%>
                        
                        <%
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
                        myApparel=""
                        
                        '// CLASS CONFIGURATION: SELECT W. NO IMAGES (C)
                        If (displayQF=True) Then 
                            pcv_strColumn1 = "col-xxs-12 col-xs-12 col-sm-1"
                            pcv_strColumn2 = "col-xxs-12 col-xs-12 col-sm-9"
                        Else 
                            pcv_strColumn1 = "col-xxs-12 col-xs-12 col-sm-1"
                            pcv_strColumn2 = "col-xxs-12 col-xs-12 col-sm-9"
                        End If 
    
                        '// START DROP-DOWN
                        if pBTODisplayType=1 then
                        
                            if (requiredCategory<>0) and (cdVar<>"1") then
                                pcv_ListForGenInfo=pcv_ListForGenInfo & "GenDropInfo(document.additem.CAG" & tempVarCat & ");" & vbcrlf
                            end if
                            %>
                            <div <%=strCol%>>
                            
                            
                                <div class='<%=pcv_strColumn1%>'>
                                    <% '// Row 1: Quantity %>
                                    <% if (displayQF=True) then %>
                                        <input class="form-control quantity" type="text" size=2 name="CAG<%=tempVarCat%>QF" value="<%=pcv_minqty%>" onBlur="javascript:testdropqty(this,'document.additem.CAG<%=tempVarCat%>');">                   <% else %>
                                        <input type="hidden" name="CAG<%=tempVarCat%>QF" value="<%=pcv_minqty%>">
                                    <% end if %>
                                </div>
                               <div class='<%=pcv_strColumn2%>'>
                                    <% '// Row 2: Select %>
                                    <input type=hidden name="app_IDProduct_CAG<%=tempVarCat%>" value="0">
                                    <input type=hidden name="app_Price_CAG<%=tempVarCat%>" value="0">
                                    <input type=hidden name="app_AddPrice_CAG<%=tempVarCat%>" value="0">
                                    <input type=hidden name="app_VIndex_CAG<%=tempVarCat%>" value="0">
                                    
                                    <select class="form-control" name="CAG<%=tempVarCat%>" onChange="testCAG<%=tempVarCat%>(); testdropdown('document.additem.CAG<%=tempVarCat%>'); CheckPreValue(this, 1, 0); showAvail<%=tempVarCat%>(this);">
            
                                    <% HiddenFields="" %>
                            
                        <% else '// if pBTODisplayType=1 then %>
                        
                            <input type="hidden" name="CAG<%=tempVarCat%>QF" value="<%=pcv_minqty%>">
                            <%
                            if Clng(requiredCategory)<>0 then
                                RTestStr="totalradio=document.additem.CAG" & tempVarCat & ".length;" & vbcrlf
                                RTestStr=RTestStr & "RadioChecked=0;" & vbcrlf
                                RTestStr=RTestStr & "if (totalradio>0) {" & vbcrlf
                                RTestStr=RTestStr & "for (var mk=0;mk<totalradio;mk++) {" & vbcrlf
                                RTestStr=RTestStr & "if (document.additem.CAG" & tempVarCat & "[mk].checked==true) { RadioChecked=1; break; } }" & vbcrlf
                                RTestStr=RTestStr & "} else { if (document.additem.CAG" & tempVarCat & ".checked==true) RadioChecked=1;}" & vbcrlf
                                RTestStr=RTestStr & "if (RadioChecked==0) {alert('"& dictLanguage.Item(Session("language")&"_alert_7") & replace(pcv_tmpArr(1,pcv_tmpN),"'","\'") & "'); return(false);}" & vbcrlf
                                ReqTestStr=ReqTestStr & RTestStr
                            end if
                            %>
                            
                        <% end if '// if pBTODisplayType=1 then %>
                
                        <% 
                        Dim requiredVar, showInfoVar, ShowInfoArray
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
                                                
                            pcv_Apparel=pcv_tmpArr(28,pcv_tmpN)
                            if IsNull(pcv_Apparel) or pcv_Apparel="" then
                                pcv_Apparel=0
                            end if
                            if pcv_Apparel="1" then
                                pcv_HaveApparel=1
                            end if
                        
                            pcv_prdDesc=pcv_tmpArr(29,pcv_tmpN)
                            pcv_prdSDesc=pcv_tmpArr(30,pcv_tmpN)
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
                            pcv_multiQty=pcv_tmpArr(27,pcv_tmpN)
                            if isNULL(pcv_multiQty) OR pcv_multiQty="" then
                                pcv_multiQty="0"
                            end if
                            
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
                            
                            if (requiredCategory<>0) and (cdVar<>"1") and (pcv_FirstItem=1) and (pBTODisplayType=1) then
                                call CheckDiscount(pcv_tmpArr(5,pcv_tmpN),true,tmp_qty,prdPrice)
                                pcv_FirstItem=3
                                pcv_tmpDefaultValue=prdPrice*tmp_qty
                                pcv_CustomizedPrice=pcv_CustomizedPrice+cdbl(pcv_tmpDefaultValue)
                                pcv_tmpDefaultDiscount=pcv_tmpIDiscount
                            end if
                            call CheckDiscount(pcv_tmpArr(5,pcv_tmpN),pcv_tmpArr(12,pcv_tmpN),tmp_qty,prdPrice)
                        
                            if pcv_Apparel="1" then
                                if pBTODisplayType=1 then
                                    CheckAPPStr=CheckAPPStr & "tmp1=document.additem.CAG" & tempVarCat & ".value; tmp2=tmp1.split('_');" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "if (eval(tmp2[0])=='" & intTempIdProduct & "') { alert('" & dictLanguage.Item(Session("language")&"_configPrd_spmsg3") & """" & replace(strDescription,"'","\'") & """'); return(false); }" & vbcrlf
                                else
                                    CheckAPPStr=CheckAPPStr & "totalradio=document.additem.CAG" & tempVarCat & ".length;" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "if (totalradio>0) {" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "for (i=0;i<document.additem.CAG" & tempVarCat & ".length;i++){" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "	if (document.additem.CAG" & tempVarCat & "[i].checked==true) {" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "tmp1=document.additem.CAG" & tempVarCat & "[i].value; tmp2=tmp1.split('_');" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "if (eval(tmp2[0])=='" & intTempIdProduct & "') { alert('" & dictLanguage.Item(Session("language")&"_configPrd_spmsg3") & """" & replace(strDescription,"'","\'") & """'); return(false); }" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "} }" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "} else {" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "	if (document.additem.CAG" & tempVarCat & ".checked==true) {" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "tmp1=document.additem.CAG" & tempVarCat & ".value; tmp2=tmp1.split('_');" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "if (eval(tmp2[0])=='" & intTempIdProduct & "') { alert('" & dictLanguage.Item(Session("language")&"_configPrd_spmsg3") & """" & replace(strDescription,"'","\'") & """'); return(false); }" & vbcrlf
                                    CheckAPPStr=CheckAPPStr & "} }" & vbcrlf
                                end if
                            end if
                            prdPrice1=prdPrice
                        
    
    'DEFAULT ITEM
                            '// Configurator Plus - S
                            if pBTODisplayType<>1 then %>
                                <input type="hidden" name="TXT<%=pcv_tmpArr(5,pcv_tmpN)%>" value="<%=ClearHTMLTags2(pcv_tmpArr(7,pcv_tmpN),0)%>">
                            <% end if
                            '// Configurator Plus - E


                            if cdefault=true then
                            
                                pExt = " "
                                ShowInfoArray = ShowInfoArray & intTempIdProduct& "," 

                                '// CLASS CONFIGURATION: RADIO W. DEFAULT (C)
                                If (displayQF=True) Then 
                                    pcv_strColumn1 = "col-xxs-12 col-xs-6 col-sm-2"
                                    pcv_strColumn2 = "col-xxs-12 col-xs-6 col-sm-2"
                                    pcv_strColumn3 = "col-xxs-12 col-xs-12 col-sm-6"
                                Else 
                                    pcv_strColumn1 = "col-xxs-12 col-xs-6 col-sm-1"
                                    pcv_strColumn2 = "col-xxs-12 col-xs-6 col-sm-3"
                                    pcv_strColumn3 = "col-xxs-12 col-xs-12 col-sm-6"
                                End If 
                                %>
                                
                                <% if pBTODisplayType=1 then %>
                                
                                    <option value="<%if sp_IDProduct>"0" then%><%=sp_IDProduct%><%else%><%=intTempIdProduct%><%end if%>_0.00_<%=weight%>_<%=prdPrice%>_<%=intTempIdProduct%>" selected><%if sp_IDProduct>"0" then%><%=sp_PrdName%><%else%><%=strDescription%><%end if%></option>
                                    
                                    <% HiddenFields=HiddenFields & "<input type=hidden name=""CAG" & tempVarCat & intTempIdProduct & "HF"" value=""" & pcv_qtyValid & "_" & pcv_minQty & """>" & vbcrlf %>
                                    <% StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  & vbcrlf 
                                    strselectvalue = func_DisplayBOMsg
                                    %>
                                        
                                <% else '// if pBTODisplayType=1 then %>


                                    <div class="<%=pcv_strColumn1%>">
                                        <% '// Row 1: Radio %>
                                        <input type="radio" name="CAG<%=tempVarCat%>" value="<%if sp_IDProduct>"0" then%><%=sp_IDProduct%><%else%><%=intTempIdProduct%><%end if%>_0.00_<%=weight%>_<%=prdPrice%>_<%=intTempIdProduct%>" checked onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>'; CheckPreValue(this, 2, 0);" class="clearBorder">
                                        <% '// Row 1: Quantity %>
                                        <% if (displayQF=True) then %>
                                            <input class="form-control quantity" type="text" size="2" id="CAG<%=tempVarCat%>QF<%=icount-1%>" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="<%=pcv_minQty%>" onblur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>,<%=pcv_multiQty%><%if (pcv_Apparel="0") AND pNostock=0 AND pcv_intBackOrder=0 AND scOutofstockpurchase=-1 AND iBTOOutofstockpurchase=-1 then%><%strQtyCheck=strQtyCheck & vbcrlf & "if (!(qttverify(document.getElementById('" & "CAG" & tempVarCat & "QF" & icount-1 & "')," & pcv_qtyvalid &"," & pcv_minQty & "," & pcv_multiQty & "," & pstock & ",1))) {setTimeout(function() {fname.focus();}, 0); return(false);}" & vbcrlf%>,<%=pstock%><%end if%>)) calculate(document.additem.CAG<%=tempVarCat%>,2);">
                                        <% else '// if (displayQF=True) then %>
                                            <input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="<%=pcv_minQty%>">
                                        <% end if '// if (displayQF=True) then %>
                                    </div>    

                                    <div class="<%=pcv_strColumn2%>"> 
                                        <% '// Row 2: Image %>
                                        <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                        <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                        <% end if %>
                                    </div>
                                    
                                    <div class="<%=pcv_strColumn3%>">
                                    
                                        <span name="CAG<%=tempVarCat%>DESC<%=icount-1%>"><%if sp_IDProduct>"0" then%><%=sp_PrdName%><%else%><%=strDescription%><%end if%></span>														

                                        <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<% if pnoprices<2 then %>TEXT<% else %>Hidden<% end if %>" value="" readonly size="<%=len(pExt)%>" class="transparentField">												
                                        <% if not pClngShowSku = 0 then %>
                                            <div class="pcSmallText"><%=strSku%></div>
                                        <% end if %>

                                        <%=func_DisplayBOMsg1(tempVarCat,intTempIdProduct) %>
                                        
                                        <% if pcv_ShowDesc="1" then %>
                                            <div class="row">
                                                <div class="col-xs-12">
                                                    <span class="configDesc">
                                                        <%=pcv_prdSDesc %>
                                                    </span>
                                                </div>
                                            </div>
                                        <% end if %>
                                    
                                <% end if 
                                
                            else '// if cdefault=true then
                            
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
                                        End If
                                    End If
                                End If
                                
                                If scDecSign="," then
                                    prdPrice=replace(prdPrice,",",".")
                                    prdPrice1=replace(prdPrice1,",",".")
                                End If 
                                
                                '// CLASS CONFIGURATION: RADIO W. NO DEFAULT (C)
                                If (displayQF=True) Then 
                                    pcv_strColumn1 = "col-xxs-12 col-xs-6 col-sm-2"
                                    pcv_strColumn2 = "col-xxs-12 col-xs-6 col-sm-2"
                                    pcv_strColumn3 = "col-xxs-12 col-xs-12 col-sm-6"
                                Else 
                                    pcv_strColumn1 = "col-xxs-12 col-xs-6 col-sm-1"
                                    pcv_strColumn2 = "col-xxs-12 col-xs-6 col-sm-3"
                                    pcv_strColumn3 = "col-xxs-12 col-xs-12 col-sm-6"
                                End If 
                                %>
                                
                                <% if pBTODisplayType=1 then %>
                                
                                    <option value="<%if sp_IDProduct>"0" then%><%=sp_IDProduct%><%else%><%=intTempIdProduct%><%end if%>_<%=prdPrice%>_<%=weight%>_<%=prdPrice1%>_<%=intTempIdProduct%>"><%if sp_IDProduct>"0" then%><%=sp_PrdName&pExt%><%else%><%=strDescription&pExt%><%end if%></option>
                                        <%
                                        HiddenFields=HiddenFields & "<input type=hidden name=""CAG" & tempVarCat & intTempIdProduct & "HF"" value=""" & pcv_qtyValid & "_" & pcv_minQty & """>" & vbcrlf
                                        strBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf %>
                                        
                                <% else '// if pBTODisplayType=1 then %>

                                        <div class="<%=pcv_strColumn1%>">
                                            <% '// Row 1: Radio %>
                                            <input type="radio" name="CAG<%=tempVarCat%>" value="<%if sp_IDProduct>"0" then%><%=sp_IDProduct%><%else%><%=intTempIdProduct%><%end if%>_<%=prdPrice%>_<%=weight%>_<%=prdPrice1%>_<%=intTempIdProduct%>" onClick="javascript:document.additem.CAG<%=tempVarCat%>QF<%=icount-1%>.value='<%=pcv_minQty%>';CheckPreValue(this, 2, 0);" class="clearBorder">
                                            <% '// Row 1: Quantity %>
                                            <% if (displayQF=True) then %>
                                                <input class="form-control quantity" type="text" size="2" id="CAG<%=tempVarCat%>QF<%=icount-1%>" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="0" onblur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>,<%=pcv_multiQty%><%if (pcv_Apparel="0") AND pNostock=0 AND pcv_intBackOrder=0 AND scOutofstockpurchase=-1 AND iBTOOutofstockpurchase=-1 then%><%strQtyCheck=strQtyCheck & vbcrlf & "if (!(qttverify(document.getElementById('" & "CAG" & tempVarCat & "QF" & icount-1 & "')," & pcv_qtyvalid &"," & pcv_minQty & "," & pcv_multiQty & "," & pstock & ",1))) {setTimeout(function() {fname.focus();}, 0); return(false);}" & vbcrlf%>,<%=pstock%><%end if%>)) calculate(document.additem.CAG<%=tempVarCat%>,2);">
                                            <% else '// if (displayQF=True) then %>
                                                <input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="0">
                                            <% end if '// if (displayQF=True) then %>
                                        </div>    
    
                                        <div class="<%=pcv_strColumn2%>">	
                                            <% '// Row 2: Image %>
                                            <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                            <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                                            <% end if %>
                                        </div>
                                        
                                        <div class="<%=pcv_strColumn3%>">																					
                                        
                                            <span name="CAG<%=tempVarCat%>DESC<%=icount-1%>"><%if sp_IDProduct>"0" then%><%=sp_PrdName%><%else%><%=strDescription%><%end if%></span>														
                                        
                                            <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%=pExt%>" readonly size="<%=len(pExt)%>" class="transparentField">																					
                                            <% if not pClngShowSku = 0 then %>
                                                <div class="pcSmallText"><%=strSku%></div>
                                            <% end if %>
                                            
											<% if pcv_Apparel=1 then %>
												<div>
													<div class="pcSmallText" id="show_CAG<%=tempVarCat&"P"&intTempIdProduct%>">
													<%if cdefault=true then%>
													<a href="javascript:win1('<%=pcf_GeneratePopupPath()%>/popup_Apparel.asp?IDBTO=<%=pIdProduct%>&IDField=CAG<%=tempVarCat%>&IDProduct=<%=intTempIdProduct%>&vindex=<%=icount-1%>&Price=<%=prdPrice%>&AddPrice=0',document.additem.CAG<%=tempVarCat%>,2)"><%=dictLanguage.Item(Session("language")&"_configPrd_spmsg2")%></a>
													<%else%>
													<a href="javascript:win1('<%=pcf_GeneratePopupPath()%>/popup_Apparel.asp?IDBTO=<%=pIdProduct%>&IDField=CAG<%=tempVarCat%>&IDProduct=<%=intTempIdProduct%>&vindex=<%=icount-1%>&Price=<%=prdPrice1%>&AddPrice=<%=prdPrice%>',document.additem.CAG<%=tempVarCat%>,2)"><%=dictLanguage.Item(Session("language")&"_configPrd_spmsg2")%></a>
													<%end if%>
													</div>
												</div>
                                        	<% end if %>

                                            <%=func_DisplayBOMsg1(tempVarCat,intTempIdProduct) %>
                                            
                                            <% if pcv_ShowDesc="1" then %>
                                                <div class="row">
                                                    <div class="col-xs-12">
                                                        <span class="configDesc">
                                                            <%=pcv_prdSDesc%>
                                                        </span>
                                                    </div>
                                                </div>
                                            <% end if %>
                                    
                                <% end if '// if pBTODisplayType=1 then %>
                            
                            <%
                            end if


                            if pBTODisplayType=1 then
                                if pcv_Apparel=1 then
                                    if cdefault=true then
                                        prdPriceA=0
                                    else
                                        prdPriceA=prdPrice
                                    end if
                                    myApparel=myApparel & "if (j==" & (icount-1) & ")" & vbcrlf
                                    myApparel=myApparel & "{" & vbcrlf
                                    myApparel=myApparel & "document.additem.app_IDProduct_CAG" & tempVarCat & ".value=" & intTempIdProduct & ";" & vbcrlf
                                    myApparel=myApparel & "document.additem.app_Price_CAG" & tempVarCat & ".value=" & prdPrice1 & ";" & vbcrlf
                                    myApparel=myApparel & "document.additem.app_AddPrice_CAG" & tempVarCat & ".value=" & prdPriceA & ";" & vbcrlf
                                    myApparel=myApparel & "document.additem.app_VIndex_CAG" & tempVarCat & ".value=" & (icount-1) & ";" & vbcrlf
                                    myApparel=myApparel & "m=1;" & vbcrlf
                                    myApparel=myApparel & "display_CAG" & tempVarCat & "();" & vbcrlf
                                    myApparel=myApparel & "return(true);" & vbcrlf
                                    myApparel=myApparel & "break;" & vbcrlf
                                    myApparel=myApparel & "}" & vbcrlf
                                end if
                            end if
																			
                            IF pBTODisplayType<>1 THEN
                                    %>
                                    </div>
                                    <div class="col-xxs-12 col-xs-12 col-sm-2">
                                        
                                        <% if showInfoVar="1" then %>
                                            
                                            <% if iBTODetLinkType=1 then%>	
                                                <a class="" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&cd=<%=replace(strCategoryDesc,"""","%22")%>')"><%=pcv_strBTODetTxt %></a>
                                            <%else%>
                                                <a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&cd=<%=replace(strCategoryDesc,"""","%22")%>')">
                                                    <span class="pcButtonText"><%=pcv_strBTODetTxt %></span>
                                                    <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>">
                                                </a>
                                            <%end if
                                            
                                        end if 
                                        %>
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
                                <%
                            END IF
    
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
                        
                            if pBTODisplayType<>1 then %>
                                <div <%=strCol%>>
                                <div class="col-xxs-12 col-xs-12 col-sm-9">
                            <%
                            end if
        
                            if Cdbl(varTempDefaultPrice)<0 then
                                if pBTODisplayType=1 then
                                    icount=icount+1 %>
                                    <option value="0_<%=varTempDefaultPrice%>_0_0_0"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%> 
                                    <%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-(varTempDefaultPrice))%><%end if%></option>
                                 <%  StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf %>
                                <% else 
                                    icount=icount+1%>
                                    <input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0_0" onClick="CheckPreValue(this, 2, 0);" class="clearBorder"><input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                    <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
                                    <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*varTempDefaultPrice)%><%end if%>" readonly size="<%=len(" - " & dictLanguage.Item(Session("language")&"_prodOpt_2") & scCurSign & money(-1*varTempDefaultPrice))%>" class="transparentField">
                                <% end if %>
                                
                            <% else if Cdbl(varTempDefaultPrice)<0 then	%>
                            
                                <% if pBTODisplayType=1 then
                                    icount=icount+1 %>
                                    <option value="0_<%=varTempDefaultPrice%>_0_0_0"><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%> 
                                    <%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%></option>
                                    <%  StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='" & func_DisplayBOMsg  &"' ;"  &vbcrlf %>
                                <% else
                                    icount=icount+1 %>
                                    <input type="radio" name="CAG<%=tempVarCat%>" value="0_<%=varTempDefaultPrice%>_0_0_0" onClick="CheckPreValue(this, 2, 0);" class="clearBorder"><input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                    <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
                                    <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%if pnoprices<2 then%><%=" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice)%><%end if%>" readonly size="<%=len(" - " & dictLanguage.Item(Session("language")&"_prodOpt_1") & scCurSign & money(varTempDefaultPrice))%>" class="transparentField">
                                <% end if %>
                            <% else if cdVar="0" then %>
                                <% if pBTODisplayType=1 then
                                    icount=icount+1 %>
                                    <option value="0_0.00_0_0_0" <% if cdVar="0" then Response.write "selected": strselectvalue = "" end if %>><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></option>
                                  <%  StrBackOrd = StrBackOrd & "availArr"&tempVarCat&"[" & intOpCnt &"]='' ;"  &vbcrlf %>
                                <% else
                                    icount=icount+1 %>
                                    <input type="radio" name="CAG<%=tempVarCat%>" value="0_0.00_0_0_0" <% if cdVar="0" then %> checked<% end if %> onClick="CheckPreValue(this, 2, 0);" class="clearBorder"><input type="hidden" name="CAG<%=tempVarCat%>QF<%=icount-1%>" value="1">
                                    <span><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_20")%></span>
                                    <input name="CAG<%=tempVarCat%>TX<%=icount-1%>" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="" readonly class="transparentField"><% end if %>
                                <% end if
                            end if 
                            end if
                            
                            if pBTODisplayType<>1 then %>
                            </div></div>
                            <% end if
                            
                        end if 
                        %>
    
                        <% if pBTODisplayType=1 then %> 
                            </select>
                    
                            <%=HiddenFields%>
                            <script type=text/javascript>
                                <%=StrBackOrd %>
                                function showAvail<%=tempVarCat%>(sel){
                                document.getElementById("AV<%=tempVarCat%>").innerHTML = availArr<%=tempVarCat%>[sel.selectedIndex] + "&nbsp;";															 
                                }
                                
                                var ns6=document.getElementById&&!document.all
                                var ie=document.all
                                function display_CAG<%=tempVarCat%>()
                                {
                                document.getElementById("show_CAG<%=tempVarCat%>").style.display="";
                                }
                                function hide_CAG<%=tempVarCat%>()
                                {
                                document.getElementById("show_CAG<%=tempVarCat%>").style.display="none";
                                }
                                
                                <%funcTestCat=funcTestCat & "testCAG" & tempVarCat & "();" & vbcrlf%>
                                function testCAG<%=tempVarCat%>()
                                {
                                var oSelect=eval("document.additem.CAG<%=tempVarCat%>");
                                var j=0;
                                for (j=0;j<oSelect.options.length;j++)
                                {
                                    if (oSelect.value==oSelect.options[j].value)
                                    {
                                        <%=myApparel%>
                                    }
                                }
                                <%if myApparel<>"" then%>
                                hide_CAG<%=tempVarCat%>();
                                <%end if%>
                                }                        
                                testdropdown('document.additem.CAG<%=tempVarCat%>');
                            </script>
                                                                                                
                            <% if intOpCnt = 0 then %>
                                <span  id="AV<%=tempVarCat%>" ><%=func_DisplayBOMsg%></span>
                             <% else %>
                                <span  id="AV<%=tempVarCat%>" ><%=strselectvalue%></span>
                             <% end if %>
                             
                        <% end if %>
                        <%
                        intOpCnt = intOpCnt + 1
                        '// END DROP-DOWN
                        %>

                            
                        <input name="currentValue<%=jCnt%>" type="HIDDEN" value="<%if (pcv_FirstItem=3) and (pBTODisplayType=1) then%><%=pcv_tmpDefaultValue%><%else%>0.00<%end if%>">
                        <input name="Discount<%=jCnt%>" type="HIDDEN" value="<%if (pcv_FirstItem=3) and (pBTODisplayType=1) then%><%=pcv_tmpDefaultDiscount%><%else%><%=pcv_tmpIDiscount%><%end if%>">
                        <input name="CAT<%=jCnt%>" type="HIDDEN" value="CAG<%=tempVarCat%>">

                        <% if pBTODisplayType=1 then %>
                            </div>
                            <div class="col-xxs-12 col-xs-12 col-sm-2">
                        <% end if %>
                        
                    <% IF pBTODisplayType=1 THEN %>
					
						<% if pcv_HaveApparel=1 then %>
                            <div id="show_CAG<%=tempVarCat%>" style="display:none;">
                                <a href="javascript:win1('<%=pcf_GeneratePopupPath()%>/popup_Apparel.asp?IDBTO=<%=pIdProduct%>&IDField=CAG<%=tempVarCat%>&IDProduct='+document.additem.app_IDProduct_CAG<%=tempVarCat%>.value+'&vindex='+document.additem.app_VIndex_CAG<%=tempVarCat%>.value+'&Price='+document.additem.app_Price_CAG<%=tempVarCat%>.value+'&AddPrice='+document.additem.app_AddPrice_CAG<%=tempVarCat%>.value,document.additem.CAG<%=tempVarCat%>,0)"><%=dictLanguage.Item(Session("language")&"_configPrd_spmsg2")%></a>
                            </div>
                        <% end if %>
                    
                        <% if showInfoVar="1" then %>
                        
                            <% if iBTODetLinkType=1 then %>
                                <a class="" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&cd=<%=replace(strCategoryDesc,"""","%22")%>')"><%=pcv_strBTODetTxt %></a>
                            <% else %>
                                <a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&cd=<%=replace(strCategoryDesc,"""","%22")%>')">
                                <span class="pcButtonText"><%=pcv_strBTODetTxt %></span>
                                <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>">
                                </a>
                            <% end if %>
                            
                        <% end if %>
                        
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
                    
                    <a href="javascript:openbrowser('<%=pcv_sffolder%>OptpriceBreaks.asp?type=<%=Session("customerType")%>&SIArray=<%=ShowInfoArray%>&cd=<%=Server.URLEnCode(replace(strCategoryDesc,"""","%22"))%>')"><img alt="<%=dictLanguage.Item(Session("language")&"_viewPrd_16")%>" src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>" align="middle"></a>
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
                    CATNotes=pcv_tmpArr(31,pcv_tmpN)
                    
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
                    <%'BTOCM-S%>
                        <div class="row">
                            <span name="CMMsg<%=pcv_tmpArr(0,pcv_tmpN)%>" id="CMMsg<%=pcv_tmpArr(0,pcv_tmpN)%>"></span>
                        </div>
                    <%'BTOCM-E%>
                    <% PrdCnt = 0 %>
                    <% ShowInfoArray = ""
                    showInfoVar="0"
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    ' START: SHOW CHECKBOXES WITH PRICE
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    
                    pcv_tmpTest=1

                DO WHILE ((pcv_tmpTest=1) AND (pcv_tmpN<=pcv_ArrCount))
                
                    pcv_Apparel=pcv_tmpArr(28,pcv_tmpN)
                    if IsNull(pcv_Apparel) or pcv_Apparel="" then
                        pcv_Apparel=0
                    end if
                    pcv_prdDesc=pcv_tmpArr(29,pcv_tmpN)
                    pcv_prdSDesc=pcv_tmpArr(30,pcv_tmpN)
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
                    pcv_multiQty=pcv_tmpArr(27,pcv_tmpN)
                    if isNULL(pcv_multiQty) OR pcv_multiQty="" then
                        pcv_multiQty="0"
                    end if
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
                            set rsCCObj2=server.CreateObject("ADODB.RecordSet")
                            set rsCCObj2=conntemp.execute(query)
                            if NOT rsCCObj2.eof then
                                intCC_BTO_Pricing=1
                                pcCC_BTO_Price=rsCCObj2("pcCC_Price")
                            end if
                            set rsCCObj2=nothing
                            
                        end if
                        set rsCCObj=nothing
                        
                    end if
                
                    '// customer logged in as ATB customer based on retail price
                    if session("ATBCustomer")=1 AND session("ATBPercentOff")=0 then
                        prdPrice=Cdbl(prdPrice)-(pcf_Round(Cdbl(prdPrice)*(cdbl(session("ATBPercentage"))/100),2))
                    end if
                
                    '// customer logged in as ATB customer based on wholesale price
                    if session("ATBCustomer")=1 AND session("ATBPercentOff")=1 then
                        prdBtoBPrice=Cdbl(prdBtoBPrice)-(pcf_Round(Cdbl(prdBtoBPrice)*(cdbl(session("ATBPercentage"))/100),2))
                        prdPrice=Cdbl(prdBtoBPrice)
                    end if
                    
                    '// customer logged in as a wholesale customer
                    if prdBtoBPrice>0 and session("customerType")=1 then
                        prdPrice=Cdbl(prdBtoBPrice)
                    end if
                    
                    '// customer logged in as a customer type with price different then the online price
                    if intCC_BTO_Pricing=1 then
                        if (pcCC_BTO_Price<>0) OR (pcCC_BTO_Price=0 AND intCC_BTO_Pricing=1) then
                            prdPrice=Cdbl(pcCC_BTO_Price)
                        end if
                    end if
                    
                    tmp_qty=pcv_minQty*ProQuantity
                
                    pcv_tmpIDiscount=0
                    call CheckDiscount(pcv_tmpArr(5,pcv_tmpN),pcv_tmpArr(12,pcv_tmpN),tmp_qty,prdPrice)
                
                    PrdCnt = PrdCnt + 1
                    jCnt = jCnt + 1 %>
                    <input name="MS<%=jCnt%>" type="HIDDEN" value="<%=VarMS%>">
                    <input name="currentValue<%=jCnt%>" type="HIDDEN" value="0">
                    <input name="Discount<%=jCnt%>" type="HIDDEN" value="<%=pcv_tmpIDiscount%>">
                    <input name="CAT<%=jCnt%>" type="HIDDEN" value="CAG<%=tempCcat%>">
                    
                    <div <%=strCol%> style="vertical-align:top">  
                        
                        <%
                        '// CLASS CONFIGURATION: CHECKBOX (C)
                        If (displayQF=True) Then 
                            pcv_strColumn1 = "col-xxs-12 col-xs-6 col-sm-2"
                            pcv_strColumn2 = "col-xxs-12 col-xs-6 col-sm-2"
                            pcv_strColumn3 = "col-xxs-12 col-xs-12 col-sm-6"
                        Else 
                            pcv_strColumn1 = "col-xxs-12 col-xs-6 col-sm-1"
                            pcv_strColumn2 = "col-xxs-12 col-xs-6 col-sm-3"
                            pcv_strColumn3 = "col-xxs-12 col-xs-12 col-sm-6"
                        End If 
                        %>
                        
                        <div class="<%=pcv_strColumn1%>">
                            
                            <% '// Row 1: Checkbox %>
                            <%'Configurator Plus - S %>
                            <input type="hidden" name="TXT<%=pcv_tmpArr(5,pcv_tmpN)%>" value="<%=ClearHTMLTags2(pcv_tmpArr(7,pcv_tmpN),0)%>">
                            <%'Configurator Plus - E %>
                            <input type="hidden" name="Cat<%=intTempIdCategory%>_Prd<%=PrdCnt%>" value="<%=intTempIdProduct%>">
                            <input type="checkbox" name="CAG<%=intTempIdCategory&intTempIdProduct%>" value="<%if sp_IDProduct>"0" then%><%=sp_IDProduct%><%else%><%=intTempIdProduct%><%end if%>_<%if (cdefault<>"") and (cdefault<>0) then%>0<%else%><%=prdPrice%><%end if%>_<%=weight%>_<%=prdPrice%>_<%=intTempIdProduct%>" onClick="javscript:document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>QF.value='<%=pcv_minQty%>'; CheckBoxPreValue(this,0);" <%if (cdefault<>"") and (cdefault<>0) then%>checked<%end if%> class="clearBorder">
                            
                            <% '// Row 1: Quantity %>
                            <%RTestStr=RTestStr & vbcrlf & "if (document.additem.CAG"& intTempIdCategory & intTempIdProduct & ".checked !=false) { RTest" & CB_CatCnt & "=" & "RTest" & CB_CatCnt & "+document.additem.CAG" & intTempIdCategory & intTempIdProduct & ".checked; }"& vbcrlf%>
                            
                            <%
                            if pcv_Apparel="1" then
                                CheckAPPStr=CheckAPPStr & "if (document.additem.CAG" & intTempIdCategory&intTempIdProduct & ".checked==true) {" & vbcrlf
                                CheckAPPStr=CheckAPPStr & "tmp1=document.additem.CAG" & intTempIdCategory&intTempIdProduct & ".value; tmp2=tmp1.split('_');" & vbcrlf
                                CheckAPPStr=CheckAPPStr & "if (eval(tmp2[0])=='" & intTempIdProduct & "') { alert('" & dictLanguage.Item(Session("language")&"_configPrd_spmsg3") & """" & replace(strDescription,"'","\'") & """'); return(false); }" & vbcrlf
                                CheckAPPStr=CheckAPPStr & "}" & vbcrlf
                            end if
                            %>
                    
                            <% if (displayQF=True) then %>
                            
                                <input class="form-control quantity" type="text" size="2" id="CAG<%=intTempIdCategory&intTempIdProduct%>QF" name="CAG<%=intTempIdCategory&intTempIdProduct%>QF" value="<%if (cdefault<>"") and (cdefault<>0) then%><%=pcv_minQty%><%else%>0<%end if%>" onBlur="if (qttverify(this,<%=pcv_qtyvalid%>,<%=pcv_minQty%>,<%=pcv_multiQty%><%if (pcv_Apparel="0") AND pNostock=0 AND pcv_intBackOrder=0 AND scOutofstockpurchase=-1 AND iBTOOutofstockpurchase=-1 then%><%strQtyCheck=strQtyCheck & vbcrlf & "if (!(qttverify(document.getElementById('" & "CAG" & intTempIdCategory&intTempIdProduct & "QF" & "')," & pcv_qtyvalid &"," & pcv_minQty & "," & pcv_multiQty & "," & pstock & ",1))) {setTimeout(function() {fname.focus();}, 0); return(false);}" & vbcrlf%>,<%=pstock%><%end if%>)) calculate(document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>,0);">
                                
                            <%else%>
                            
                                <input type="hidden" name="CAG<%=intTempIdCategory&intTempIdProduct%>QF" value="<%if (cdefault<>"") and (cdefault<>0) then%><%=pcv_minQty%><%else%>0<%end if%>">
                                
                            <%end if%>
                        
                        </div>    
                    
                        <div class="<%=pcv_strColumn2%>">
                            <% '// Row 2: Image %>
                            <% if strSmallImage <> "hide" and pClngShowItemImg <> 0 then %>
                                <img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",strSmallImage)%>" style="width: <%=pClngSmImgWidth%>px;" alt="<%=strDescription%>" align="top">
                            <% end if %>
                        </div>
                
                        <div class="<%=pcv_strColumn3%>">
                            <% '// Row 3: Details %>
                            <span name="CAG<%=intTempIdCategory&intTempIdProduct%>DESC0"><%if sp_IDProduct>"0" then%><%=sp_PrdName%><%else%><%=strDescription%><%end if%></span>
                            <%if pnoprices<2 then%>
                                &nbsp;-&nbsp;<input name="CAG<%=intTempIdCategory&intTempIdProduct%>TX0" type="<%if pnoprices<2 then%>TEXT<%else%>Hidden<%end if%>" value="<%=scCurSign & money(prdPrice)%>" readonly class="transparentField" size="<%=len("" & scCurSign & money(prdPrice))%>">
                            <%end if%>
                    
                            <% if not pClngShowSku = 0 then %>
                                <div class="pcSmallText"><%=strSku%></div>
                            <% end if %>
							
							<%if pcv_Apparel=1 then%><div class="pcSmallText" id="show_CAG<%=intTempIdCategory&intTempIdProduct&"P"&intTempIdProduct%>"><a href="javascript:win1('<%=pcf_GeneratePopupPath()%>/popup_Apparel.asp?IDPROD=<%=intTempIdProduct%>&IDBTO=<%=pIdProduct%>&IDField=CAG<%=intTempIdCategory&intTempIdProduct%>&IDProduct=<%=intTempIdProduct%>&Price=<%=prdPrice%>&AddPrice=<%if (cdefault<>"") and (cdefault<>0) then%>0<%else%><%=prdPrice%><%end if%>&vindex=0',document.additem.CAG<%=intTempIdCategory&intTempIdProduct%>,0)"><%=dictLanguage.Item(Session("language")&"_configPrd_spmsg2")%></a></div><%end if%>
                    
                            <%if pcv_ShowDesc="1" then%>
                                <div class="configDesc"><%=pcv_prdSDesc%></div>
                            <%end if%>
                        </div>
                        
                        <div class="col-xxs-12 col-xs-12 col-sm-2">
                            <% '// Row: More %>
                    
                            <% if showInfoVar="1" then %>
                    
                                <% if iBTODetLinkType=1 then %>	
                                    <a class="" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&cd=<%=replace(strCategoryDesc,"""","%22")%>')"><%=pcv_strBTODetTxt %></a>
                                <% else %>
                                    <a class="pcButton pcConfigDetail tiny" href="javascript:viewWin('ShowInfo.asp?<% if pBTODisplayType<>1 then %>IDPROD=<%=intTempIdProduct%>&<% end if %>IDBTO=<%=pIdProduct%>&IDCat=<%=tempVarCat%>&cd=<%=replace(strCategoryDesc,"""","%22")%>')">
										<span class="pcButtonText"><%=pcv_strBTODetTxt %></span>
										<img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>"></a>
                                <% end if
                                
                            end if 
                            %>
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
                    
                    <% if func_DisplayBOMsg <>"" then %>
                        <div <%=strCol%>>  
                            <%=func_DisplayBOMsg1(tempVarCat,intTempIdProduct)%>    
                        </div>
                    <% end if %>
                    

                    
                    <%
                    pcv_tmpN=pcv_tmpN+1
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
        
    end if //'Have Configurator Categories
    set rsSSobj=nothing
    '******* END Configurator Categories
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
    response.write "return(eval(roundNumber(DisValue1,2)));" & VBCrlf
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
    response.write "return(eval(roundNumber(DisValue1,2)));" & VBCrlf
    response.write " } </script>" & VBCRlf
    %>
    <script type=text/javascript>
        function roundNumber(num, dec) {
        var tmp1 = Math.round(num*Math.pow(10,dec))/Math.pow(10,dec);
        return(tmp1);
        }
        function chkR()
        {
        <%
        'Configurator Plus - S
        if pcv_HaveRules=1 then%>
        var tmp1=CheckCatBeforeSubmit();
        if (tmp1=="no")
        {
        return(false);
        }
        <%end if
        'Configurator Plus - E
        %>
        <%if ReqTestStr<>"" then%>
        <%=ReqTestStr%>
        <%end if%>
        <%=CheckAPPStr%>
        if (checkproqty(document.additem.quantity))
        {
        <%
        'Configurator Plus - S
        if pcv_HaveRules=1 then%>
        var tmp1=CheckCatBeforeSubmit();
        if (tmp1=="no")
        {
        return(false);
        }
        <%end if
        'Configurator Plus - E
        %>
        return CheckTotalItemQty();
        }
        else
        {
        return(false);
        }
        }    
        <%=funcTestCat%>    
    </script>
    <% Call pcs_GetDefaultBTOItemsMin %>
    <input type="hidden" name="FirstCnt" value="<%=jCnt%>">
    <input type="hidden" name="CB_CatCnt" value="<%=CB_CatCnt%>">
</div>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  product configuration table
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Get Minimum Quantity of Default Configurator Items
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
' END:  Get Minimum Quantity of Default Configurator Items
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Display Totals
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_DisplayTotals
%>	
<div id="pcBTOfloatPrices">
<div class="pcTable">
	<div class="pcTableRowFull">
	<div class="pcTableColumn60">
		<b><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_4")%></b>
	</div>
	<div class="pcTableColumn1"></div>
	<div class="pcTableColumn39"> 
		<input name="curPrice" type="hidden" value="<%=scCurSign & money(pPrice) %>">
		<input name="TLcurPrice" type="TEXT" style="text-align:right;" value="<%=scCurSign & money(pPrice) %>" readonly size="10" class="transparentField">
		<input name="TLPriceDefault" type="hidden">
	</div>
	</div> 
	
	<div class="pcTableRowFull">
	<div class="pcTableColumn60">
		<input name="currentValue0" type="HIDDEN" value="<%=pPrice%>">
		<input name="jCnt" type="HIDDEN" value="<%=jCnt%>">
		<input name="ConfigCartIndex" type="HIDDEN" value="<%=f%>">
		<input name="ConfigSession" type="HIDDEN" value="<%=pConfigSession%>">
		<b><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_5")%></b> 
	</div>
	<div class="pcTableColumn1">
		<% 
		'// If this is the Quote Page, then display a different total value.
		if len(pConfigWishlistSession)>0 then 
		%>
		<input name="total" type="hidden" value="<%=pcv_CustomizedPrice%>">		
		<% else %>
		<input name="total" type="hidden" value="<%=pcv_CustomizedPrice%>">
		<% end if %>
	</div>
	<div class="pcTableColumn39">
		<input name="TLtotal" type="TEXT" style="text-align:right;" value="<%=scCurSign & money(pcv_CustomizedPrice)%>" readonly size="10" class="transparentField">
	</div>
	</div>
	
	<%if TempDiscountStr<>"" then%>
	<div class="pcTableRowFull">
	<div class="pcTableColumn60"> 
		<b><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_8")%></b>
	</div>
	<div class="pcTableColumn1"></div>
	<div class="pcTableColumn39">
		<input name="Discounts" type="TEXT" style="text-align:right;" value="<%=scCurSign & money(pcv_ItemDiscounts)%>" readonly size="10" class="transparentField">
	</div>
	</div>
	<%end if%>
	
	<%if pDiscountPerQuantity=-1 then%>
	<div class="pcTableRowFull">
	<div class="pcTableColumn60"> 
		<b><%=bto_dictLanguage.Item(Session("language")&"_CustviewPastD_3")%></b>
	</div>
	<div class="pcTableColumn1"></div>
	<div class="pcTableColumn39">
		<input name="QDiscounts" type="TEXT" style="text-align:right;" value="<%=scCurSign & money(0)%>" readonly size="10" class="transparentField">
		<input name="QDiscounts0" type="hidden" value="<%=scCurSign & money(0)%>">
	</div>
	</div>
	<%else%>
		<input name="QDiscounts" type="hidden" value="<%=scCurSign & money(0)%>">
		<input name="QDiscounts0" type="hidden" value="<%=scCurSign & money(0)%>">
	<%end if%>
	
	<div class="pcTableRowFull">
	<div class="pcTableColumn60"> 
		<b><%=bto_dictLanguage.Item(Session("language")&"_configurePrd_6")%></b>
	</div>
	<div class="pcTableColumn1">
		<%if TempDiscountStr="" then%>
		<input name="Discounts" type="hidden" value="<%=scCurSign & money(pcv_ItemDiscounts)%>">
		<%end if%>
	
		<%
		'// If this is the Quote Page, then display a different total value.
		if len(pConfigWishlistSession)>0 then
		%>
		<input name="GrandTotal" type="hidden" value="<%=scCurSign & money(pPrice)%>">
		<input name="UGrandTotal" type="hidden" value="<%=scCurSign & money(pPrice)%>">
		<% else %>
		<input name="GrandTotal" type="hidden" value="<%=scCurSign & money(pDefaultPrice)%>">
		<input name="UGrandTotal" type="hidden" value="<%=scCurSign & money(pPrice)%>">
		<% end if %>
	</div>
	<div class="pcTableColumn39">
		<input name="TotalWithQD" type="TEXT" style="text-align:right;" value="<%=scCurSign & money(pPrice+pcv_CustomizedPrice*ProQuantity-Round(pcv_ItemDiscounts+0.001,2))%>" readonly size="10" class="transparentField">
		<input name="TLGrandTotal" type="hidden" value="<%=scCurSign & money(pPrice)%>">
		<input name="CMDefault" type="hidden">
		<input name="CMWQD" type="hidden">
	</div>
	</div>       
</div>
</div>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Display Totals
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_BTODiscounts
%>
	<%if pDiscountPerQuantity=-1 then %>
	<div class="pcTable pcShowList">
		<% If session("customerType")="1" then %>
			<div class="col-xs-12">
				<a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=pidProduct%>')"><img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>"></a>
				<a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?idproduct=<%=pidProduct%>&type=1')"><%= dictLanguage.Item(Session("language")&"_viewPrd_16")%></a>
			</div>
		<% else %>
			<div class="col-xs-12"><a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=pidProduct%>')"><img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("discount"))%>"></a>
			<a href="javascript:openbrowser('<%=pcv_sffolder%>priceBreaks.asp?idproduct=<%=pidProduct%>')"><%= dictLanguage.Item(Session("language")&"_viewPrd_16")%></a>
			</div>
		<% end if %>
	</div>
	<%end if%>
<%
End Sub	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Discounts
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  X Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Dim pcXFArr,intXFCount
Public Sub pcs_BTOXFields
Dim i

xrequired="0"
xfieldCnt=0
reqstring=""

dim tmpCount
tmpCount=0

'Get XFields
intXFCount=-1
query="SELECT IdXField,pcPXF_XReq FROM pcPrdXFields WHERE idProduct=" & pidProduct & ";"
set rs=connTemp.execute(query)
if not rs.eof then
	pcXFArr=rs.getRows()
	intXFCount=ubound(pcXFArr,2)
end if
set rs=nothing

IF intXFCount>=0 THEN
For i=0 to intXFCount
	'select from the database more info 
	query= "SELECT xfield,textarea,widthoffield,rowlength,maxlength FROM xfields WHERE idxfield="&pcXFArr(0,i)
	set rsfieldObj=server.createobject("adodb.recordset")
	set rsfieldObj=conntemp.execute(query)
	
	if not rsfieldObj.EOF then '// Check for no field in DB, although referenced by the Configurable Product
		pxfield_OK=1
		xfield=rsfieldObj("xfield")
		textarea=rsfieldObj("textarea")
		widthoffield=rsfieldObj("widthoffield")
		rowlength=rsfieldObj("rowlength")
		maxlength=rsfieldObj("maxlength")
		set rsfieldObj=nothing

		tmpCount=tmpCount+1
		pxreq=pcXFArr(1,i)
						
		if pxreq="-1" then
			xfieldCnt=xfieldCnt+1
			xrequired="1"
			if reqstring<>"" then
				reqstring=reqstring & ","
			end if
			reqstring=reqstring&"additem.xfield" & tmpCount & ".value,'"&replace(xfield,"'","\'")&"'"
		end if
		
		tmpxfield=request("xfield" & tmpCount)
		if tmpxfield="" then
			tmpxfield=session("SFxfield" & tmpCount & "_" & pIdProduct)
		end if
		%>
		<div class="pcFormItem">
			<input type="hidden" name="xf<%=tmpCount%>" value="<%=pcXFArr(0,i)%>">

			<div class="pcFormField">
			<% if textarea="-1" then %>

                <label for="xfield<%=tmpCount%>"><%=xfield%>:</label>
                <textarea name="xfield<%=tmpCount%>" cols="<%=widthoffield%>" rows="<%=rowlength%>" <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'<%=tmpCount%>',<%=maxlength%>);"<%end if%>><%=tmpxfield%></textarea>
                <%if maxlength>"0" then%>
                    <br />
                    <%= dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar<%=tmpCount%>" name="countchar<%=tmpCount%>" style="font-weight: bold"><%=maxlength%></span> <%= dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
                <%end if%>
                
			<% else %>

                <label for="xfield<%=tmpCount%>"><%=xfield%>:</label>
				<input class="form-control quantity" type="text" name="xfield<%=tmpCount%>" size="<%=widthoffield%>" maxlength="<%=maxlength%>" value="<%=tmpxfield%>" <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'<%=tmpCount%>',<%=maxlength%>);"<%end if%>>
				<%if maxlength>"0" then%>
                    <br />				
                    <%= dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar<%=tmpCount%>" name="countchar<%=tmpCount%>" style="font-weight: bold"><%=maxlength%></span> <%= dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
				<%end if%>
                
			<% end if %>
            
			</div>
		</div>
	<%end if
Next
END IF%>
		<input type="hidden" name="XFCount" value="<%=tmpCount%>" />
		<%if (pxfield_OK=1) then%>
			<%if tmpCount>0 then%>
				<script type=text/javascript>
				function testchars(tmpfield,idx,maxlen)
				{
					var tmp1=tmpfield.value;
					if (tmp1.length>maxlen)
					{
						alert("<%= dictLanguage.Item(Session("language")&"_CheckTextField_1")%>" + maxlen + "<%= dictLanguage.Item(Session("language")&"_CheckTextField_1a")%>");
						tmp1=tmp1.substr(0,maxlen);
						tmpfield.value=tmp1;
						document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
						tmpfield.focus();
					}
					document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
				}
				</script>
			<%end if%>
			<hr>
		<%end if%>
<%
End sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  X Fields
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  X Fields - Reconfigure
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_XFieldsReconfigure
Dim i,j
%>
	<%
	xrequired="0"
	xfieldCnt=0
	reqstring=""
	
	dim tmpCount
	tmpCount=0
	
	if xstr<>"" then
		if pcv_strAdminPrefix="1" then
			xstr=replace(xstr, "<BR>", vbCRLF)
			xarray=split(xstr,"|")
		else
			xarray=split(xstr,"<br>")
		end if
	end if	 
	
	'Get XFields
	intXFCount=-1
	query="SELECT IdXField,pcPXF_XReq FROM pcPrdXFields WHERE idProduct=" & pidProduct & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcXFArr=rs.getRows()
		intXFCount=ubound(pcXFArr,2)
	end if
	set rs=nothing
	
	IF intXFCount>=0 THEN
	For i=0 to intXFCount
		'select from the database more info 
		query= "SELECT * FROM xfields WHERE idxfield="&pcXFArr(0,i)
		set rsfieldObj=server.createobject("adodb.recordset")
		set rsfieldObj=conntemp.execute(query)
		
		tmpCount=tmpCount+1
		maxlength=rsfieldObj("maxlength")
		pxreq=pcXFArr(1,i)
									
		if pxreq="-1" then
			xfieldCnt=xfieldCnt+1
			xrequired="1"
			if reqstring<>"" then
				reqstring=reqstring & ","
			end if
			reqstring=reqstring&"additem.xfield" & tmpCount & ".value,'"&replace(rsfieldObj("xfield"),"'","\'")&"'"
		end if
		%>		
		<div class="pcFormItem">
		<input type="hidden" name="xf<%=tmpCount%>" value="<%=pcXFArr(0,i)%>">
				
		<% if rsfieldObj("textarea")="-1" then %> 
			<div class="pcFormField"> 

                <label for="xfield<%=tmpCount%>"><%=xfield%>:</label>
			    <textarea class="form-control" name="xfield<%=tmpCount%>" cols="<%=rsfieldObj("widthoffield")%>" rows="<%=rsfieldObj("rowlength")%>" <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'<%=tmpCount%>',<%=maxlength%>);"<%end if%>><% if xstr<>"" then%><% for j=0 to ubound(xarray) %><% tempstr=xarray(j) %><% strArray=split(tempstr,":") %><% if trim(strArray(0))=trim(rsfieldObj("xfield")) then %><%=replace(trim(pcf_ReverseGetUserInput(strArray(1))),"<BR>",VbCrLf) %><%exit for%><% end if %><% next %><% end if %></textarea>
                <%if maxlength>"0" then%>
                    <br />
                    <%= dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar<%=tmpCount%>" name="countchar<%=tmpCount%>" style="font-weight: bold"><%=maxlength%></span> <%= dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
                <%end if%>
			</div>
            
		<% else %>
        
			<div class="pcFormField">
            
                <label for="xfield<%=tmpCount%>"><%=xfield%>:</label>	
			    <input class="form-control" type="text" name="xfield<%=tmpCount%>" size="<%=rsfieldObj("widthoffield")%>" maxlength="<%=rsfieldObj("maxlength")%>" value="<% if xstr<>"" then %><% for j=0 to ubound(xarray) %><% tempstr=xarray(j) %><% strArray=split(tempstr,":") %><% if trim(strArray(0))=trim(rsfieldObj("xfield")) then %><%=trim(pcf_ReverseGetUserInput(strArray(1))) %><%exit for%><% end if %><% next %><% end if %>" <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'<%=tmpCount%>',<%=maxlength%>);"<%end if%>>
			    <%if maxlength>"0" then%>
			        <br />			
			        <%= dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar<%=tmpCount%>" name="countchar<%=tmpCount%>" style="font-weight: bold"><%=maxlength%></span> <%= dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
			    <%end if%>
			</div>
            
		<% end if %>
		</div>
	
	<%Next
	END IF%>
	
	<input type="hidden" name="XFCount" value="<%=tmpCount%>" />
	
	<%if tmpCount>0 then%>
		<script type=text/javascript>
			function testchars(tmpfield,idx,maxlen)
			{
				var tmp1=tmpfield.value;
				if (tmp1.length>maxlen)
				{
					alert("<%= dictLanguage.Item(Session("language")&"_CheckTextField_1")%>" + maxlen + "<%= dictLanguage.Item(Session("language")&"_CheckTextField_1a")%>");
					tmp1=tmp1.substr(0,maxlen);
					tmpfield.value=tmp1;
					document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
					tmpfield.focus();
				}
				document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
			}
		</script>
	<%end if%>
<%
End sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  X Fields - Reconfigure
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  X Fields - Reconfigure Quote
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_XFieldsQuote
Dim i,j%>
	<%
	xrequired="0"
	xfieldCnt=0
	reqstring=""
	
	dim tmpCount
	tmpCount=0
	
	if xstr<>"" then
	xarray=split(xstr,"||")
	end if 
	
	'Get XFields
	intXFCount=-1
	query="SELECT IdXField,pcPXF_XReq FROM pcPrdXFields WHERE idProduct=" & pidProduct & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcXFArr=rs.getRows()
		intXFCount=ubound(pcXFArr,2)
	end if
	set rs=nothing
	
	IF intXFCount>=0 THEN
	For i=0 to intXFCount
		'select from the database more info 
		mySQL= "SELECT * FROM xfields WHERE idxfield="&pcXFArr(0,i)
		set rsfieldObj=server.createobject("adodb.recordset")
		set rsfieldObj=conntemp.execute(mySQL)
		
		if not rsfieldObj.EOF then
			tmpCount=tmpCount+1
			maxlength=rsfieldObj("maxlength")
			pxreq=pcXFArr(1,i)
									
			if pxreq="-1" then
				xfieldCnt=xfieldCnt+1
				xrequired="1"
				if reqstring<>"" then
					reqstring=reqstring & ","
				end if
				reqstring=reqstring&"additem.xfield" & tmpCount & ".value,'"&replace(rsfieldObj("xfield"),"'","\'")&"'"
			end if
			%>
			
			<div class="pcFormItem">
            
			    <input type="hidden" name="xf<%=tmpCount%>" value="<%=pcXFArr(0,i)%>">
                
                
			<% if rsfieldObj("textarea")="-1" then %>
				
				<div class="pcFormField">
                    <label for="xfield<%=tmpCount%>"><%=rsfieldObj("xfield")%>:</label>
                    <textarea class="form-control" name="xfield<%=tmpCount%>" cols="<%=rsfieldObj("widthoffield")%>" rows="<%=rsfieldObj("rowlength")%>" <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'<%=tmpCount%>',<%=maxlength%>);"<%end if%>><% if xstr<>"" then%><%for j=0 to (ubound(xarray)-1) %><% tempstr=xarray(j) %><% strArray=split(tempstr,"|") %><% if Clng(strArray(0))=Clng(pcXFArr(0,i)) then %><%=trim(pcf_ReverseGetUserInput(strArray(1))) %><%exit for%><% end if %><% next %><% end if %></textarea>
                    <%if maxlength>"0" then%>                    
                        <%= dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar<%=tmpCount%>" name="countchar<%=tmpCount%>" style="font-weight: bold"><%=maxlength%></span> <%= dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>                    
                    <%end if%>                
				</div>
                
			<% else %>
				
				<div class="pcFormField">                
                    <label for="xfield<%=tmpCount%>"><%=rsfieldObj("xfield")%>:</label>
                    <input class="form-control quantity" type="text" name="xfield<%=tmpCount%>" size="<%=rsfieldObj("widthoffield")%>" maxlength="<%=rsfieldObj("maxlength")%>" value="<% if xstr<>"" then %><% for j=0 to (ubound(xarray)-1) %><% tempstr=xarray(j) %><% strArray=split(tempstr,"|") %><% if Clng(strArray(0))=Clng(pcXFArr(0,i)) then %><%=trim(pcf_ReverseGetUserInput(strArray(1))) %><% end if %><% next %><% end if %>" <%if maxlength>"0" then%>onkeyup="javascript:testchars(this,'<%=tmpCount%>',<%=maxlength%>);"<%end if%>>
                    <%if maxlength>"0" then%>				
                        <%= dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar<%=tmpCount%>" name="countchar<%=tmpCount%>" style="font-weight: bold"><%=maxlength%></span> <%= dictLanguage.Item(Session("language")&"_GiftWrap_5b")%>
                    <%end if%>
				</div>
			<% end if %>
			</div> 
		<%
		end if ' EOF
	Next
	END IF%>

	<input type="hidden" name="XFCount" value="<%=tmpCount%>" />
	
	<%if tmpCount>0 then%>
		<script type=text/javascript>
			function testchars(tmpfield,idx,maxlen)
			{
				var tmp1=tmpfield.value;
				if (tmp1.length>maxlen)
				{
					alert("<%= dictLanguage.Item(Session("language")&"_CheckTextField_1")%>" + maxlen + "<%= dictLanguage.Item(Session("language")&"_CheckTextField_1a")%>");
					tmp1=tmp1.substr(0,maxlen);
					tmpfield.value=tmp1;
					document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
					tmpfield.focus();
				}
				document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
			}
		</script>
	<%end if%>
<%
End sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  X Fields - Reconfigure Quote
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_ProductImage
	if len(pImageUrl) > 0 then
	'// A)  The image exists
	%>
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
                    <a href="<%=pcv_strZoomLink%>" <%=pcv_strZoomLocation%>><img id='mainimg' class='img-responsive' name='mainimg' src='<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pImageUrl)%>' alt="<%=replace(pDescription,"""","&quot;")%>" /></a>
                <% else %>
                    <img id='mainimg' class='img-responsive' name='mainimg' src='<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pImageUrl)%>' alt="<%=replace(pDescription,"""","&quot;")%>" />
                <% end if %>
                <% if pcv_strUseEnhancedViews = True then %>
                	<div class="<%=pcv_strHighSlide_Heading%>"><%=replace(pDescription,"""","&quot;")%></div>
                <% end if %>
            </div>
    
		<% if len(pLgimageURL)>0 and pcv_strUseEnhancedViews = False then %>
				<div style="width:100%; text-align:right;">
                    <a href="<%=pcv_strZoomLink%>" <%=pcv_strZoomLocation%>><img src="<%=pcf_getImagePath(pcv_tmpNewPath,rsIconObj("zoom"))%>" hspace="10" alt="<%=dictLanguage.Item(Session("language")&"_altTag_5")%>"></a>
                    <% if pcv_strUseEnhancedViews = True then %>
                    	<div class="<%=pcv_strHighSlide_Heading%>"><%=replace(pDescription,"""","&quot;")%></div>
                    <% end if %>
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
                <% if bCounter>0 AND pcv_Apparel<>"1" then %>
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
			<div class="pcShowMainImage">
				<img name='mainimg' src='<%=pcf_getImagePath(pcv_tmpNewPath & "catalog","no_image.gif")%>' alt="<%=replace(pDescription,"""","&quot;")%>">
			</div>
	<% 
	end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Product Image (If there is one)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Show Configurator Prices
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_BTOPrices
%>
				<div class="pcShowProductPrice" id="pcBTOhideTopPrices">
					<%
					if pBtoBPrice>0 and session("customerType")=1 then
						pPrice=pBtoBPrice
					End if
					
						if (pPrice>0) and (pnoprices<2) then
							if session("customerType")=1 then
								response.write dictLanguage.Item(Session("language")&"_viewPrd_15")%>
								<input name="TLdefaultprice" type="text" value="<%=scCurSign%><%=money(pPrice)%>" readonly size="14" class="transparentField">
								<%
							else
								response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_2")%>
								<input name="TLdefaultprice" type="text" value="<%=scCurSign%><%=money(pPrice)%>" readonly size="14" class="transparentField">
								<%
							end if
						else 
						%>
						<input name="TLdefaultprice" type="hidden" value="<%=scCurSign%><%=money(pPrice)%>">	
						<% 
						end if
						
						if pnoprices=2 then %>
						<input name="GrandTotal2" type="hidden" value="<%=scCurSign%><%=money(pPrice)%>">
						<input name="TLGrandTotal2" type="hidden" value="<%=scCurSign%><%=money(pPrice)%>">
						<input name="TLGrandTotal2QD" type="hidden" value="<%=scCurSign%><%=money(pPrice)%>">
						<% else
						response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_3")%>
						<input name="GrandTotal2" type="hidden" value="<%=scCurSign%><%=money(pPrice)%>">
						<input name="TLGrandTotal2" type="hidden" value="<%=scCurSign%><%=money(pPrice)%>">
						<input name="TLGrandTotal2QD" type="Text" value="<%=scCurSign%><%=money(pPrice)%>" readonly size="14" class="transparentField">
						<% end if %>
					
					</div>
<%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Show Configurator Prices
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Sub CreateAppPopUp()
%>
<script>
	var pcv_field="";
	var pcv_type=0;
	function win1(fileName,tmpfield,ctype)
	{
		myFloater=window.open('','myWindow','resizable=yes,width=650,height=500,status=1,scrollbars=yes')
		myFloater.location.href=fileName;
		pcv_field=tmpfield;
		pcv_type=ctype;
		checkwin();
	}
	function checkwin()
	{
		if (myFloater.closed)
		{
			calculate(pcv_field,pcv_type);
		}
		else
		{
			setTimeout('checkwin()',500);
		}
	}
</script>
<%
End Sub




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
	tmpquery = Request.ServerVariables("QUERY_STRING")
	If Left(tmpquery,4) = "404;" Then
		tmppath = Request.ServerVariables("PATH_INFO")
		tmppath = Left(tmppath,InStr(tmppath,"/pc/")-1)
		tmpquery = Mid(tmpquery,InStr(tmpquery,"/pc/"),Len(tmpquery)-InStr(tmpquery,"/pc/")+1)
		tmpquery = tmppath & tmpquery
		If InStr(tmpquery, "?") = 0 Then
			tmpquery = tmpquery & "?"
		Else
			tmpquery = tmpquery & "&"
		End If
	Else
		tmpquery = Request.ServerVariables("PATH_INFO") & "?" & tmpquery & "&"
	End If
	
	If (Request.ServerVariables("HTTPS") = "off") Then
		pcv_srtRestoreFreshURL = "http://" & Request.ServerVariables("SERVER_NAME") & tmpquery
	Else
		pcv_srtRestoreFreshURL = "https://" & Request.ServerVariables("SERVER_NAME") & tmpquery
	End If
	'// Set a hidden field value. We will use this value as an environmental indicator. Its our "Coal Mine Canary".
	response.write "<form name=""pct"" id=""pct""><input name=""pctf"" value=""1"" type=""hidden""></form>"
	'// This block of javascript will check the hidden field value each time the page loads.
	response.write "<script type=""text/javascript"">"& chr(10)
	response.write "var intro=""1"";"& chr(10)
	response.write "var outro;"& chr(10)
	response.write "function pcf_isRefreshNeeded() {"& chr(10)
	response.write "   pcf_isBTOReady();"& chr(10) '// Lock "Add To Cart" button until page loads
	response.write "   outro = (intro!=document.pct.pctf.value);"& chr(10)
	response.write "   document.pct.pctf.value=2;"& chr(10)
	response.write "   pcf_doRefreshNeeded();"& chr(10)
	response.write "}"& chr(10)
	response.write "function pcf_doRefreshNeeded() {"& chr(10)
	response.write "	if (outro==true) {"& chr(10)
	response.write " 			// re-load the page"& chr(10)
	response.write "			 window.location = '"&pcv_srtRestoreFreshURL&"&time=" & Year(Date())&Month(Date())&Day(Date())&Hour(Time())&Minute(Time())&Second(Time()) & "';"& chr(10)
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
	response.write "function pcf_isBTOReady() {"& chr(10)
	response.write "   if (document.getElementById('addtocart')!=null) {"& chr(10)
	response.write "   		document.getElementById('addtocart').disabled='';" & chr(10)
	response.write "   }"& chr(10)
	response.write "}"& chr(10)
	response.write "</script>"& chr(10)
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Cache Errors
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

Function func_DisplayBOMsg1(tempCat,tempPrd)
	if isNULL(pMinPurchase) or pMinPurchase="" then
		pMinPurchase=0
	end if
	If (scOutofStockPurchase=-1) AND ((CLng(pStock)<1) OR (clng(pStock)<clng(pMinPurchase))) AND (Clng(pNoStock)=0) AND (Clng(pcv_intBackOrder)=1) Then
		If clng(pcv_intShipNDays)>0 then		  
			func_DisplayBOMsg1 = "<span style=""padding-left: 22px;"" id=""AV" & tempCat & "P" & tempPrd & """>" & dictLanguage.Item(Session("language")&"_viewPrd_60")&dictLanguage.Item(Session("language")&"_sds_viewprd_1") & pcv_intShipNDays & dictLanguage.Item(Session("language")&"_sds_viewprd_1b") & "</span>"
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
' START:  Find Path to PC Folder
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function pcf_GeneratePopupPath()
	Dim pcv_strConstructPath, pcv_intStartLoc
	pcv_strConstructPath=Request.ServerVariables("PATH_INFO")	
	pcv_intStartLoc=(InStrRev(pcv_strConstructPath,"/")-1)
	pcv_strConstructPath=mid(pcv_strConstructPath,1,InStrRev(pcv_strConstructPath,"/",pcv_intStartLoc)-1)
	If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
		pcf_GeneratePopupPath="http://" & Request.ServerVariables("HTTP_HOST") & pcv_strConstructPath & "/pc"
	Else
		pcf_GeneratePopupPath="https://" & Request.ServerVariables("HTTP_HOST") & pcv_strConstructPath & "/pc"
	End if
End Function	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Find Path to PC Folder
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