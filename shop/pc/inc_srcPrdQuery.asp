<%
If statusAPP="1" Then
	pcv_strHideSubSearch=False '// Set to "True" to disable sub-product search
Else
	pcv_strHideSubSearch=True
End If

iPageSize=getUserInput(request("resultCnt"),10)
if iPageSize="" then
	iPageSize=getUserInput(request("iPageSize"),0)
end if
if (iPageSize="") then
	iPageSize = (cdbl(scPrdRow) * cdbl(scPrdRowsPerPage))
end if
if (not IsNumeric(iPageSize)) then
	iPageSize=6
end if
if request("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=server.HTMLEncode(request("iPageCurrent"))
end if

pSKU=getUserInput(request("SKU"),150)
pKeywords=getUserInput(request("keyWord"),100)
pCValues=getUserInput(request("SearchValues"),0)
tKeywords=pKeywords
tIncludeSKU=getUserInput(request("includeSKU"),10)
if tIncludeSKU = "" then
	tIncludeSKU = "true"
end if
pPriceFrom=getUserInput(request("priceFrom"),20)
if NOT isNumeric(pPriceFrom) then
	pPriceFrom=0
end if
if Instr(pPriceFrom,",")>Instr(pPriceFrom,".") then
	pPriceFrom=replace(pPriceFrom,",",".")
end if
pPriceUntil=getUserInput(request("priceUntil"),20)
if NOT isNumeric(pPriceUntil) then
	pPriceUntil=9999999
end if
if Instr(pPriceUntil,",")>Instr(pPriceUntil,".") then
	pPriceUntil=replace(pPriceUntil,",",".")
end if
if src_ForCats="1" then
	pIdCategory=0
	else
	pIdCategory=getUserInput(request("idCategory"),4)
	if NOT validNum(pIdCategory) or trim(pIdCategory)="" then
		pIdCategory=0
	end if
end if
pIdSupplier=getUserInput(request("idSupplier"),4)
if NOT validNum(pIdSupplier) or trim(pIdSupplier)="" then
	pIdSupplier=0
end if
pWithStock=getUserInput(request("withStock"),2)
pcustomfield=getUserInput(request("customfield"),0)
if pcustomfield="" then
	pcustomfield="0"
end if
	
IDBrand=getUserInput(request("IDBrand"),20)
if NOT validNum(IDBrand) or trim(IDBrand)="" then
	IDBrand=0
end if

incSale=getUserInput(request("incSale"),4)
if NOT validNum(incSale) or trim(incSale)="" then
	incSale=0
end if
tmpIDSale=getUserInput(request("IDSale"),4)
if NOT validNum(tmpIDSale) or trim(tmpIDSale)="" then
	tmpIDSale=0
end if

strORD=getUserInput(request("order"),4)
if NOT validNum(strORD) or trim(strORD)="" then
	strORD=3
end if

if strORD<>"" then
	Select Case StrORD
		Case "0": strORD1="A.sku ASC, A.idproduct DESC"
		Case "1": strORD1="A.description ASC"
		Case "2":
			If Session("customerType")=1 then
				strORD1 = "(CASE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE A.bToBPrice WHEN 0 THEN A.Price ELSE A.bToBPrice END) ELSE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN A.pcProd_BTODefaultPrice ELSE A.pcProd_BTODefaultWPrice END) END) DESC"
			else
				strORD1 = "(CASE (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) WHEN 0 THEN A.Price ELSE A.pcProd_BTODefaultPrice END) DESC"
			End if
		Case "3":
			If Session("customerType")=1 then
				strORD1 = "(CASE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE A.bToBPrice WHEN 0 THEN A.Price ELSE A.bToBPrice END) ELSE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN A.pcProd_BTODefaultPrice ELSE A.pcProd_BTODefaultWPrice END) END) ASC"
			else
				strORD1 = "(CASE (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) WHEN 0 THEN A.Price ELSE A.pcProd_BTODefaultPrice END) ASC"
			End if
		Case Else: strORD1="A.description ASC"
	End Select
Else
	strORD="1"
	strORD1="A.idproduct ASC"
End If
	
PrdTypeStr=""

'src_IncNormal=getUserInput(request("src_IncNormal"),0)
src_IncNormal=1
'src_IncBTO=getUserInput(request("src_IncBTO"),0)
src_IncBTO=1
src_IncItem=getUserInput(request("src_IncItem"),0)
src_Special=getUserInput(request("src_Special"),0)
src_Featured=getUserInput(request("src_Featured"),0)

if src_IncNormal="" then
	src_IncNormal="0"
end if

if src_IncBTO="" then
	src_IncBTO="0"
end if

if src_IncItem="" then
	src_IncItem="0"
end if

if src_Special="" then
	src_Special="0"
end if

if src_Featured="" then
	src_Featured="0"
end if

if (src_IncBTO="0") and (src_IncItem="0") then
	src_IncNormal="1"
end if

if (src_IncBTO="1") and (src_IncItem="0") and (src_IncNormal="0") then
	PrdTypeStr=" AND serviceSpec<>0 "
end if

if (src_IncBTO="0") and (src_IncItem="1") and (src_IncNormal="0") then
	PrdTypeStr=" AND configOnly<>0 "
end if

if (src_IncBTO="1") and (src_IncItem="1") and (src_IncNormal="0") then
	PrdTypeStr=" AND ((serviceSpec<>0) OR (configOnly<>0)) "
end if

if (src_IncBTO="0") and (src_IncItem="1") and (src_IncNormal="1") then
	PrdTypeStr=" AND serviceSpec=0 "
end if

if (src_IncBTO="1") and (src_IncItem="0") and (src_IncNormal="1") then
	PrdTypeStr=" AND configOnly=0 "
end if

if (src_IncBTO="0") and (src_IncItem="0") and (src_IncNormal="1") then
	PrdTypeStr=" AND ((serviceSpec=0) AND (configOnly=0)) "
end if

	
' create sql statement
strSQL=""
tmpSQL=""
tmpSQL2=""

tmpSQL1=",categories_products,categories "

tmp_StrQuery=""
if session("customerCategory")="" or session("customerCategory")=0 then
	If session("customerType")=1 then
		tmp_StrQuery="(A.serviceSpec<>0 AND A.pcProd_BTODefaultWPrice>="&pPriceFrom&" And A.pcProd_BTODefaultWPrice<=" &pPriceUntil&")"
	else
		tmp_StrQuery="(A.serviceSpec<>0 AND A.pcProd_BTODefaultPrice>="&pPriceFrom&" And A.pcProd_BTODefaultPrice<=" &pPriceUntil&")"
	end if
else
	tmp_StrQuery="(A.serviceSpec<>0 AND A.idproduct IN (SELECT idproduct FROM pcBTODefaultPriceCats WHERE pcBTODefaultPriceCats.idCustomerCategory=" & session("customerCategory") & " AND pcBTODefaultPriceCats.pcBDPC_Price>="&pPriceFrom&" AND pcBTODefaultPriceCats.pcBDPC_Price<=" &pPriceUntil&"))"
end if

pcv_strMaxResults=SRCH_MAX
If pcv_strMaxResults>"0" Then
	pcv_strLimitPhrase="TOP " & pcv_strMaxResults
Else
	pcv_strLimitPhrase=""
End If

tmpCatquery=""
tmpSubquery=""

if src_ForCats="1" then
	
	'// Category Search
	strSQL= "SELECT "& pcv_strLimitPhrase &" COUNT(categories.idcategory) AS ProductCount, categories.idcategory, categories.categoryDesc FROM "
	strSQL=strSQL& "(categories_products INNER JOIN categories ON categories_products.idcategory=categories.idcategory) "
	strSQL=strSQL& "LEFT OUTER JOIN products as A ON A.idProduct=categories_products.idProduct "
	strSQL=strSQL& "WHERE (" & tmp_StrQuery & " OR (A.serviceSpec=0 AND A.price>="&pPriceFrom&" And A.price<=" &pPriceUntil&")) AND A.active=-1 AND A.removed=0 "  
  	strSQL=strSQL & " AND categories.iBTOhide=0"
  	if session("CustomerType")<>"1" then
		tmpiHide=1
	  	strSQL=strSQL & " AND categories.pccats_RetailHide=0"
	else
		tmpiHide=0
  	end if
	tmpHiddenCat=pcf_FindHiddenCatList(tmpiHide,1)
	if tmpHiddenCat<>"" then
		strSQL=strSQL & " AND (NOT categories.idCategory IN (" & tmpHiddenCat & ")) "
	end if
	
	If statusAPP="1" Then
		tmpSubquery= "SELECT "& pcv_strLimitPhrase &" A.idProduct FROM "
		tmpSubquery=tmpSubquery& "(categories_products INNER JOIN categories ON categories_products.idcategory=categories.idcategory) "
		tmpSubquery=tmpSubquery& "LEFT OUTER JOIN products as A ON A.idProduct=categories_products.idProduct "
		tmpSubquery=tmpSubquery& "WHERE (A.active=-1 AND A.removed=0 "  
	  	tmpSubquery=tmpSubquery & " AND categories.iBTOhide=0"
		tmpCatquery=tmpCatquery & " AND categories.iBTOhide=0"
	  	if session("CustomerType")<>"1" then
		  	tmpSubquery=tmpSubquery & " AND categories.pccats_RetailHide=0"
			tmpCatquery=tmpCatquery & " AND categories.pccats_RetailHide=0"
	  	end if
	End If
	
else
	
	'// Product Search
	strSQL= "SELECT "& pcv_strLimitPhrase &" A.idProduct, A.sku, A.description, A.price, A.listHidden, A.listPrice, A.serviceSpec, A.bToBPrice, A.smallImageUrl, A.noprices, A.stock, A.noStock, A.pcprod_HideBTOPrice, A.pcProd_BackOrder, A.FormQuantity, A.pcProd_BackOrder, A.pcProd_BTODefaultPrice " '// , "& zSQL &" "
	strSQL=strSQL& "FROM products A "
	strSQL=strSQL& " WHERE (A.active=-1 AND A.removed=0 AND A.idProduct IN (" 

	'// START: Category Sub-Query
	strSQL=strSQL& "SELECT B.idProduct FROM categories_products B INNER JOIN categories C ON "
	strSQL=strSQL & "C.idCategory=B.idCategory WHERE C.iBTOhide=0 "
	if pIdCategory<>"0" then
		if (schideCategory = "1") OR (SRCH_SUBS = "1") then			
			if len(TmpCatList)>0 then
				strSQL=strSQL & " AND B.idCategory IN ("& TmpCatList &")" '// include sub cats
			else
				strSQL=strSQL & " AND B.idCategory=" &pIdCategory	
			end if
		else
			strSQL=strSQL & " AND B.idCategory=" &pIdCategory	
		end if
	end if
	if session("CustomerType")<>"1" then
		tmpiHide=1
		strSQL=strSQL & " AND C.pccats_RetailHide=0"
	else
		tmpiHide=0
	end if
	tmpHiddenCat=pcf_FindHiddenCatList(tmpiHide,1)
	if tmpHiddenCat<>"" then
		strSQL=strSQL & " AND (NOT C.idCategory IN (" & tmpHiddenCat & ")) "
	end if
	'// END: Category Sub-Query

	strSQL=strSQL& ") AND (" & tmp_StrQuery & " OR (A.serviceSpec=0 AND A.configOnly=0 AND A.price>="&pPriceFrom&" AND A.price<=" &pPriceUntil&")) " 

end if

if (incSale>"0") then
	if tmpIDSale="0" then
		strSQL=strSQL & " AND A.pcSC_ID>0"
	else
		strSQL=strSQL & " AND A.pcSC_ID=" & tmpIDSale
	end if
end if

if len(pSKU)>0 then
	strSQL=strSQL & " AND A.sku like '%"&pSKU&"%'"
end if

if (pIdSupplier<>"0") and (pIdSupplier<>"10") then
	strSQL=strSQL & " AND A.idSupplier=" &pIdSupplier
end if

if pWithStock="-1" then
    If statusAPP="1" Then    
	    strSQLSP=strSQLSP & " AND (((A.stock>0 OR A.noStock<>0) AND (A.pcProd_Apparel=0)) "
        strSQLSP=strSQLSP & " OR (((SELECT SUM(SP.stock) FROM products SP WHERE SP.pcProd_ParentPrd=A.idProduct) > 0 OR A.noStock<>0) AND (A.pcProd_Apparel=1))) "   
    Else    
        strSQLSP=strSQLSP & " AND (A.stock>0 OR A.noStock<>0) "        
    End IF    
    strSQL = strSQL & strSQLSP    
end if

if (IDBrand&""<>"") and (IDBrand&""<>"0") then
	strSQL=strSQL & " AND A.IDBrand=" & IDBrand
end if

if src_Special="1" then
   strSQL=strSQL & " AND A.hotdeal<>0" 
end if

if src_Special="2" then
   strSQL=strSQL & " AND A.hotdeal=0" 
end if

if src_Featured="1" then
   strSQL=strSQL & " AND A.showInHome<>0" 
end if

if src_Featured="2" then
   strSQL=strSQL & " AND A.showInHome=0" 
end if

TestWord=""
if request("exact")<>"1" then
	if Instr(pKeywords," AND ")>0 then
		keywordArray=split(pKeywords," AND ")
		TestWord=" AND "
	else
		if Instr(pKeywords," and ")>0 then
			keywordArray=split(pKeywords," and ")
			TestWord=" AND "
		else
			if Instr(pKeywords,",")>0 then
				keywordArray=split(pKeywords,",")
				TestWord=" OR "
			else
				if (Instr(pKeywords," OR ")>0) then
					keywordArray=split(pKeywords," OR ")
					TestWord=" OR "
				else
					if (Instr(pKeywords," or ")>0) then
						keywordArray=split(pKeywords," or ")
						TestWord=" OR "
					else
						if (Instr(pKeywords," ")>0) then
							keywordArray=split(pKeywords," ")
							TestWord=" AND "
						else
							keywordArray=split(pKeywords,"***")	
							TestWord=" OR "
						end if
					end if
				end if
			end if
		end if
	end if
else
	pKeywords=trim(pKeywords)
	if pKeywords<>"" then
		pKeywords="'" & pKeywords & "'***'%[^a-zA-z0-9]" & pKeywords & "[^a-zA-z0-9]%'***'" & pKeywords & "[^a-zA-z0-9]%'***'%[^a-zA-z0-9]" & pKeywords & "'"
	end if
	keywordArray=split(pKeywords,"***")	
	TestWord=" OR "
end if

tmpStrEx=""
if pCValues<>"" AND pCValues<>"0" then
	tmpSValues=split(pCValues,"||")
	For k=lbound(tmpSValues) to ubound(tmpSValues)
		if tmpSValues(k)<>"" then	
			sfquery=""
			sfquery = "SELECT pcSearchFields_Products.idproduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData=" & tmpSValues(k)
			set rsSearchFields=Server.CreateObject("ADODB.Recordset")
			set rsSearchFields=connTemp.execute(sfquery)
			If NOT rsSearchFields.eof Then
				SearchFieldArray = pcf_ColumnToArray(rsSearchFields.getRows(),0)
				SearchFieldString = Join(SearchFieldArray,",")		
				If len(SearchFieldString)>0 Then
					tmpStrEx=tmpStrEx & " AND A.idproduct IN ("& SearchFieldString &")"
				End If
			Else
				tmpStrEx=tmpStrEx & " AND A.idproduct IN (0)"				
			End If
			set rsSearchFields = nothing
		end if
	Next
end if

IF request("exact")<>"1" THEN

if pKeywords<>"" then

	strSQl=strSql & " AND ("
	
	tmpSQL="(A.details LIKE "
	tmpSQL2="(A.description LIKE "
	tmpSQL3="(A.sDesc LIKE "
	tmpSQL5="(A.pcProd_MetaKeywords LIKE "
	if tIncludeSKU="true" then
		tmpSQL4="(A.SKU LIKE "
	end if
	Pos=0
	For L=LBound(keywordArray) to UBound(keywordArray)
		if trim(keywordArray(L))<>"" then
		Pos=Pos+1
		if Pos>1 Then
			tmpSQL=tmpSQL  & TestWord & " A.details LIKE "
			tmpSQL2=tmpSQL2 & TestWord & " A.description LIKE "
			tmpSQL3=tmpSQL3 & TestWord & " A.sDesc LIKE "
			tmpSQL5=tmpSQL5 & TestWord & " A.pcProd_MetaKeywords LIKE "
			if tIncludeSKU="true" then
				tmpSQL4=tmpSQL4 & TestWord & " A.SKU LIKE "
			end if
		end if
			tmpSQL=tmpSQL  & "'%" & trim(keywordArray(L)) & "%'"
			tmpSQL2=tmpSQL2 & "'%" & trim(keywordArray(L)) & "%'"
			tmpSQL3=tmpSQL3 & "'%" & trim(keywordArray(L)) & "%'"
			tmpSQL5=tmpSQL5 & "'%" & trim(keywordArray(L)) & "%'"
			if tIncludeSKU="true" then
				tmpSQL4=tmpSQL4 & "'%" & trim(keywordArray(L)) & "%'"
			end if
		end if
	Next
	tmpSQL=tmpSQL & ")"
	tmpSQL2=tmpSQL2 & ")"
	tmpSQL3=tmpSQL3 & ")"
	tmpSQL5=tmpSQL5 & ")"
	if tIncludeSKU="true" then
		tmpSQL4=tmpSQL4 & ")"
	end if
	
	strSQL=strSQL & tmpSQL
	strSQL=strSQL & " OR " & tmpSQL2
	strSQL=strSQL & " OR " & tmpSQL5
	if tIncludeSKU="true" then
		strSQL=strSQL & " OR " & tmpSQL3
		strSQL=strSQL & " OR " & tmpSQL4 & ")"
	else	
		strSQL=strSQL & " OR " & tmpSQL3 & ")"
	end if
	strSQL=strSQL & PrdTypeStr & tmpStrEx
	if src_ForCats<>"1" then
		strSQL=strSQL& ")"
		query=strSQL & " ORDER BY " & strORD1
	end if
else
	strSQL=strSQL & PrdTypeStr & tmpStrEx
	if src_ForCats<>"1" then
		strSQL=strSQL& ")"
		query=strSQL & " ORDER BY " & strORD1
	end if
end if

ELSE 'Exact=1

if pKeywords<>"" then

	strSQl=strSql & " AND ("
	
	tmpSQL="(A.details LIKE "
	tmpSQL2="(A.description LIKE "
	tmpSQL3="(A.sDesc LIKE "
	tmpSQL5="(A.pcProd_MetaKeywords LIKE "
	if tIncludeSKU="true" then
		tmpSQL4="(A.SKU LIKE "
	end if
	Pos=0
	For L=LBound(keywordArray) to UBound(keywordArray)
		if trim(keywordArray(L))<>"" then
		Pos=Pos+1
		if Pos>1 Then
			tmpSQL=tmpSQL  & TestWord & " A.details LIKE "
			tmpSQL2=tmpSQL2 & TestWord & " A.description LIKE "
			tmpSQL3=tmpSQL3 & TestWord & " A.sDesc LIKE "
			tmpSQL5=tmpSQL5 & TestWord & " A.pcProd_MetaKeywords LIKE "
			if tIncludeSKU="true" then
				tmpSQL4=tmpSQL4 & TestWord & " A.SKU LIKE "
			end if
		end if
			tmpSQL=tmpSQL & trim(keywordArray(L))
			tmpSQL2=tmpSQL2 & trim(keywordArray(L))
			tmpSQL3=tmpSQL3 & trim(keywordArray(L))
			tmpSQL5=tmpSQL5 & trim(keywordArray(L))
			if tIncludeSKU="true" then
				tmpSQL4=tmpSQL4 & trim(keywordArray(L))
			end if
		end if
	Next
	tmpSQL=tmpSQL & ")"
	tmpSQL2=tmpSQL2 & ")"
	tmpSQL3=tmpSQL3 & ")"
	tmpSQL5=tmpSQL5 & ")"
	if tIncludeSKU="true" then
		tmpSQL4=tmpSQL4 & ")"
	end if
	
	strSQL=strSQL & tmpSQL
	strSQL=strSQL & " OR " & tmpSQL2
	strSQL=strSQL & " OR " & tmpSQL5
	if tIncludeSKU="true" then
		strSQL=strSQL & " OR " & tmpSQL3
		strSQL=strSQL & " OR " & tmpSQL4 & ")"
	else	
		strSQL=strSQL & " OR " & tmpSQL3 & ")"
	end if
	strSQL=strSQL & PrdTypeStr & tmpStrEx
	if src_ForCats<>"1" then
		strSQL=strSQL& ")"
		query=strSQL & " ORDER BY " & strORD1
	end if
else
	strSQL=strSQL & PrdTypeStr & tmpStrEx
	if src_ForCats<>"1" then
		strSQL=strSQL& ")"
		query=strSQL & " ORDER BY " & strORD1
	end if
end if

END IF 'Exact

if src_ForCats="1" then
	pIdCategory=getUserInput(request("idCategory"),4)
	if NOT validNum(pIdCategory) then
		pIdCategory=0
	end if
	query=strSQL& " GROUP BY categories.idcategory, categories.categoryDesc "
	query=query& " ORDER BY categories.idcategory; "
End If

If statusAPP="1" Then
	If src_ForCats<>"1" Then
	
	    If (pcv_strHideSubSearch=False) AND (IDBrand=0) Then
		
			tmpNewQuery=""
			tmpNewQuery=strSQL
			tmpMarker = instr(strSQL,"AND ((A.serviceSpec<>0")
			tmpNewQuery1=left(strSQL,tmpMarker-1)
			tmpNewQuery1=replace(tmpNewQuery1,"AND A.removed=0","AND A.removed=0 AND A.configOnly=0")
			tmpNewQuery2=right(strSQL,len(strSQL)-tmpMarker+1)
			tmpNewQuery2=replace(tmpNewQuery2,tmpStrEx,"")
            tmpNewQuery2=replace(tmpNewQuery2, strSQLSP, " AND (A.stock>0 OR A.noStock<>0) ")
			tmpNewQuery2=replace(tmpNewQuery2,"(A.","(D.")
			tmpNewQuery2=replace(tmpNewQuery2,"AND A.","AND D.")
			tmpNewQuery2=replace(tmpNewQuery2,"OR A.","OR D.")
			tmpNewQuery2=replace(tmpNewQuery2," A."," D.")
                       
			tmpSubQuery = "SELECT D.pcProd_ParentPrd FROM products D WHERE ("
			tmpSubQuery = tmpSubQuery & "D.active=0 AND D.pcProd_SPInActive=0 AND D.pcProd_ParentPrd>0 AND D.removed=0 "
			tmpSubQuery = tmpSubQuery & tmpNewQuery2 & " group by D.pcProd_ParentPrd"
	                                                                                                
			Set rs=Server.CreateObject("ADODB.Recordset")
			rs.Open tmpSubQuery, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
	                                
			pcv_strParentMatching = ""
			if NOT rs.eof then        
				pcv_strParentMatching = pcf_ColumnToArray(rs.getRows(),0)
				pcv_strParentMatching = JOIN(pcv_strParentMatching,",")
			end if
	
			if pcv_strParentMatching<>"" then
				tmpNewQuery3=tmpNewQuery1&"AND (A.idProduct IN ("&pcv_strParentMatching&")))" & tmpStrEx
				Set rs=Server.CreateObject("ADODB.Recordset")
	
				rs.Open tmpNewQuery3, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
				pcv_strParentMatching = ""
				if NOT rs.eof then      
					pcv_strParentMatching = pcf_ColumnToArray(rs.getRows(),0)
					pcv_strParentMatching = JOIN(pcv_strParentMatching,",")
				end if
			else
				pcv_strParentMatching = ""                  
			end if
	
			set rs=nothing
			If len(pcv_strParentMatching)>0 Then
				tmpSubQuery2=" OR (A.idProduct IN ( "& pcv_strParentMatching &" ))"
				query=replace(query, " ORDER", tmpSubQuery2 & " ORDER")
			End If
	
	    End If
	
	Else
	
	    If pcv_strHideSubSearch=False Then
		
		    tmpSQL1="A.idProduct=categories_products.idProduct WHERE ("
		    tmp1=split(query,tmpSQL1)
		    tmp2=split(tmp1(1),"GROUP BY")
		    tmp3=replace(tmp2(0)," AND A.active=-1"," AND A.active=0 AND A.pcProd_SPInActive=0 AND A.pcProd_ParentPrd>0")
		    tmp3a=split(tmp3," OR (A.serviceSpec=0")
		    tmp3="((A.serviceSpec=0" & tmp3a(1)
		    tmp3=replace(tmp3,"A.","B.")	
		    tmp4 = "SELECT B.pcProd_ParentPrd FROM products B INNER JOIN categories_products ON categories_products.idProduct = B.idProduct INNER JOIN categories ON categories_products.idcategory=categories.idcategory WHERE " & tmp3
			Set rs=Server.CreateObject("ADODB.Recordset")
		    rs.Open tmp4, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
		    tmp_strParentMatching = ""
		    if NOT rs.eof then	
			    tmp_strParentMatching = pcf_ColumnToArray(rs.getRows(),0)
			    tmp_strParentMatching = JOIN(tmp_strParentMatching,",")
		    end if
		    set rs=nothing
	
			if tmp_strParentMatching<>"" then
			    tmp5=" ) AND ( A.idProduct IN ( "& tmp_strParentMatching &" )"
			    tmp5=tmp5 & ") "
			    tmp5=tmp5 & " GROUP BY A.idProduct "
			    tmpNewQuery3=tmpSubquery
			    tmpNewQuery3=replace(tmpNewQuery3, "WHERE ((", "WHERE ( ((")
			    tmpNewQuery3=tmpNewQuery3 & tmp5
	
			    Set rs=Server.CreateObject("ADODB.Recordset")
			    rs.Open tmpNewQuery3, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
			    pcv_strParentMatching = ""
			    if NOT rs.eof then	
				    pcv_strParentMatching = pcf_ColumnToArray(rs.getRows(),0)
				    pcv_strParentMatching = JOIN(pcv_strParentMatching,",")
			    end if
			else
				pcv_strParentMatching = ""
			end if
		    set rs=nothing
		    If len(pcv_strParentMatching)>0 Then
			    tmpSubQuery2=" ) OR (A.idProduct IN ( "& pcv_strParentMatching &" ))"
			    query=replace(query, " GROUP", tmpSubQuery2 & " GROUP")
			    query=replace(query, "WHERE ((", "WHERE ( ((")
		    End If
		
	    End If
	
	End If
End If
%>