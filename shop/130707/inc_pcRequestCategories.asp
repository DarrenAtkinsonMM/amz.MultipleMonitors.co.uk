<%


Dim pcArrP,pcArrP1,pcArrP2,tmpDropCatList
Dim intCountQ,intCountQ1
intCountQ=-1
intCountQ1=-1
tmpDropCatList=""

if NOT isNumeric(idRootCat) or idRootCat="" then
	idRootCat=1
end if

if myCats="" then
	myCats="0,"
end if

if pcv_CP="" then
	pcv_CP="0"
end if

queryQ="SELECT categories.idcategory,categories.categorydesc,categories.idParentCategory, categories.pccats_BreadCrumbs FROM categories ORDER BY categories.categoryDesc ASC;"
set rsQ=ConnTemp.execute(queryQ)
if not rsQ.eof then
	pcArrP=rsQ.getRows()
	intCountQ=ubound(pcArrP,2)
end if
set rsQ=nothing


if pcv_CP="1" then
	queryQ = "SELECT DISTINCT categories.idcategory,categories.categorydesc,categories.idParentCategory, categories.pccats_BreadCrumbs FROM categories INNER JOIN categories_products ON categories.idCategory=categories_products.idCategory ORDER BY categories.categoryDesc ASC;"
	Set rsQ = Server.CreateObject("ADODB.Recordset")
	set rsQ=ConnTemp.execute(queryQ)		
	if not rsQ.eof then
		pcArrP1=rsQ.getRows()
		intCountQ1=ubound(pcArrP1,2)
	end if
	set rsQ=nothing
	pcArrP2=pcArrP1
else
	if pcv_CP="3" then
		queryQ = "SELECT DISTINCT categories.idcategory,categories.categorydesc,categories.idParentCategory, categories.pccats_BreadCrumbs FROM categories INNER JOIN (categories_products INNER JOIN options_optionsGroups ON categories_products.idProduct=options_optionsGroups.idProduct) ON categories.idCategory=categories_products.idCategory ORDER BY categories.categoryDesc ASC;"
		Set rsQ = Server.CreateObject("ADODB.Recordset")
		set rsQ=ConnTemp.execute(queryQ)		
		if not rsQ.eof then
			pcArrP1=rsQ.getRows()
			intCountQ1=ubound(pcArrP1,2)
		end if
		set rsQ=nothing
		pcArrP2=pcArrP1
	else
		if idRootCat>"1" then
			queryQ = "SELECT idCategory,categoryDesc,idParentCategory,pccats_BreadCrumbs FROM [categories] WHERE (idParentCategory="&idRootCat& ") OR (idCategory="&idRootCat&") OR (pccats_BreadCrumbs LIKE '" & idRootCat & "||%') OR (pccats_BreadCrumbs LIKE '%|,|" & idRootCat & "||%');"
			set rsQ=ConnTemp.execute(queryQ)		
			if not rsQ.eof then
				pcArrP1=rsQ.getRows()
				intCountQ1=ubound(pcArrP1,2)
			end if
			set rsQ=nothing
			pcArrP2=pcArrP1
		else
			pcArrP2=pcArrP
			intCountQ1=intCountQ
		end if
	end if
end if

if intCountQ1>=0 then
	dim tmp_A
	tmp_A=split(myCats,",")
	For kQ=0 to intCountQ1
		pcv_BC=pcf_catGetParent(pcArrP2(0,kQ),pcArrP2(2,kQ),pcArrP2(3,kQ))
		if pcv_BC<>"" then
			pcv_BC=replace(pcv_BC,"""", "&quot;")
			pcv_BC=replace(pcv_BC,"<", "&lt;")
			pcv_BC=replace(pcv_BC,">", "&gt;")
		end if
		x_categoryDesc=pcArrP2(1,kQ)
		if x_categoryDesc<>"" then
			x_categoryDesc=replace(x_categoryDesc,"""", "&quot;")
			x_categoryDesc=replace(x_categoryDesc,"<", "&lt;")
			x_categoryDesc=replace(x_categoryDesc,">", "&gt;")
		end if
		
		Dim x_pcSelected, lQ
		x_pcSelected = ""
		For lQ=lbound(tmp_A) to ubound(tmp_A)
			if trim(tmp_A(lQ))<>"" then
			if clng(tmp_A(lQ))=clng(pcArrP2(0,kQ)) then
				x_pcSelected=" selected"
				exit for
			end if
			end if
		Next
		
		tmpDropCatList=tmpDropCatList & "<option value=""" & pcArrP2(0,kQ) & """" & x_pcSelected & ">" & x_categoryDesc & " " & pcv_BC & "</option>" & vbCrLf
	Next
end if

if tmpDropCatList="" then
	tmpDropCatList="NONE"
end if
response.write tmpDropCatList

Function pcf_catGetParent(pcv_idcategory,pcv_parentCategory,pcv_BreadCrumbs)	
	Dim tmp_ParentText,tmp_C,tmp_D,p
	tmp_ParentText=""
	if isNULL(pcv_BreadCrumbs) then
		pcv_BreadCrumbs=""
	end if
	IF trim(pcv_BreadCrumbs)<>"" THEN
		tmp_C=split(pcv_BreadCrumbs,"|,|")
		For p=lbound(tmp_C) to ubound(tmp_C)
			if trim(tmp_C(p))<>"" then
			tmp_D=split(tmp_C(p),"||")
			if (clng(tmp_D(0))<>clng(pcv_idcategory)) AND (clng(tmp_D(0))<>"1") then
				if tmp_ParentText="" then
					tmp_ParentText="["
				else
					tmp_ParentText=tmp_ParentText & "/"
				end if
				tmp_ParentText=tmp_ParentText & tmp_D(1)
			end if
			end if
		Next
		if tmp_ParentText<>"" then
			tmp_ParentText=tmp_ParentText & "]"
		else
			tmp_ParentText=""
		end if
	ELSE
		pcv_tmpParent=""
		tmpParent=""
		if pcv_parentCategory="1" then
		else 
			pcv_tmpParent = pcf_FindParent(pcv_parentCategory)
			if pcv_tmpParent<>"" then
				pcv_tmpParent="[" & pcv_tmpParent & "]"
			end if
		end if		
		if pcv_tmpParent="" then
			pcv_tmpParent=""
		end if
		tmp_ParentText=pcv_tmpParent
	END IF
	pcf_catGetParent=tmp_ParentText
End Function

Function pcf_FindParent(idCat)
	Dim k
	if clng(idCat)<>1 then
	For k=0 to intCountQ
		if (clng(pcArrP(0,k))=clng(idCat)) and (clng(pcArrP(0,k))<>1)	then
			if tmpParent<>"" then
			tmpParent="/" & tmpParent
			end if
			tmpParent=pcArrP(1,k) & tmpParent
			pcf_FindParent(pcArrP(2,k))
			exit for
		end if
	Next
	pcf_FindParent=tmpParent
	end if
End function


%>