<%
	'// Clear unspecified error so the breadcrumbs load
	if err.number<>0 then
		err.clear
	end if
	
'// Get category information

	'// START - Check if category still exists in the database
	
		Dim pcCatExists
		pcCatExists=1
		query="SELECT idCategory FROM categories WHERE idCategory="&pIdCategory
		SET rs=Server.CreateObject("ADODB.RecordSet")
		SET rs=conntemp.execute(query)
	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		if rs.eof then ' The category no longer exists
			pcCatExists=0
		end if
		
		set rs=nothing
		
		if pIdProduct="" then
			pIdProduct=0
		end if
	
		'// If category ID doesn't exist and we have the product ID, 
		'// get the first category that the product has been assigned to, filtering out hidden categories
		if pcCatExists=0 and validNum(pIdProduct) and pIdProduct>0 then
		
			' If customer is not wholesale, disallow wholesale-only categories
			if not session("customerType")="1" then
				queryW = " AND categories.pccats_RetailHide<>1"
			end if
			' If admin preview, ignore hidden categories
			if session("pcv_intAdminPreview")<>1 then
				queryHC = " AND categories.iBTOhide<>1" & queryW
			end if		
			query="SELECT categories_products.idCategory FROM categories_products INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE categories_products.idProduct="& pIdProduct & queryHC &";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if not rs.EOF then
				pIdCategory=rs("idCategory")
				pcCatExists=1
			end if
			set rs=nothing
			
		end if
	
		if pcCatExists=0 then ' The category no longer exists and we did not find another one
			call closeDb()
			response.redirect "msg.asp?message=86"   
		end if
	'// END - Check if category still exists in the database
	
	'// Proceed to retrieving category information

	' If customer is wholesale, allow wholesale-only categories
		if not session("customerType")="1" then
			queryW = " AND categories.pccats_RetailHide<>1"
		end if
		' If admin preview, ignore hidden categories
		if session("pcv_intAdminPreview")<>1 and session("admin") <> 0 then
			queryHC = " AND categories.iBTOhide<>1" & queryW
		end if	

	query="SELECT categories.categoryDesc, categories.idCategory, categories.idParentCategory, categories.image, categories.largeimage, categories.SDesc, categories.LDesc, categories.HideDesc, categories.pccats_BreadCrumbs FROM categories WHERE (((categories.idCategory)="&pIdCategory&")" & queryHC & ") ORDER BY categories.priority, categories.categoryDesc;"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rs.eof then
		set rs=nothing
		call closeDb()
		response.redirect "msg.asp?message=86"           
	end if

	pCategoryName=rs("categoryDesc")
	plargeImage=rs("largeimage")
	if pLargeImage = "no_image.gif" then
		pLargeImage = ""
	end if
	SDesc=rs("SDesc")
	LDesc=rs("LDesc")
	HideDesc=rs("HideDesc")
	if isNULL(HideDesc) OR HideDesc="" then
		HideDesc="0"
	end if
	pccats_BreadCrumbs=rs("pccats_BreadCrumbs")
	set rs=nothing
	
'// GET breadcrumb information (location of this category in the category tree):
' (a) if it exists, parse and display
' (b) if it does not exist, create and save for future use

' (a)
IF pccats_BreadCrumbs<>"" AND instr(pccats_BreadCrumbs,"||") THEN
	pcArrayBreadCrumbs=split(pccats_BreadCrumbs,"|,|")
	strBreadCrumb=""
	for i=0 to ubound(pcArrayBreadCrumbs)
		pcArrayCrumb=split(pcArrayBreadCrumbs(i),"||")
		intBCId = pcArrayCrumb(0) 
		strBCDesc = pcArrayCrumb(1)

		'// Call SEO Routine
		pcGenerateSeoLinks
		'//
		if i=0 then
			IF (i = ubound(pcArrayBreadCrumbs)) AND (request("idproduct")="") AND (request("idcategory")<>"") THEN
				strBreadCrumb=strBreadCrumb & "<span itemprop=""itemListElement"" itemscope itemtype=""http://schema.org/ListItem"">"
				strBreadCrumb=strBreadCrumb & "<span itemprop=""name"">" & pcArrayCrumb(1) &"</span>"
				strBreadCrumb=strBreadCrumb & "<meta itemprop=""position"" content=""" & i+1 & """ />"
			ELSE
				strBreadCrumb=strBreadCrumb & "<span itemprop=""itemListElement"" itemscope itemtype=""http://schema.org/ListItem"">"
				strBreadCrumb=strBreadCrumb & "<a itemprop=""item"" href=""" & pcStrBCLink & """"
				strBreadCrumb=strBreadCrumb & "><span itemprop=""name"">" & pcArrayCrumb(1) &"</span></a>"
				strBreadCrumb=strBreadCrumb & "<meta itemprop=""position"" content=""" & i+1 & """ />"
				strBreadCrumb=strBreadCrumb & "</span>"
			END IF
		else
			IF (I = ubound(pcArrayBreadCrumbs)) AND (request("idproduct")="") AND (request("idcategory")<>"") THEN
				strBreadCrumb=strBreadCrumb & " > "
				strBreadCrumb=strBreadCrumb & "<span itemprop=""itemListElement"" itemscope itemtype=""http://schema.org/ListItem"">"
				strBreadCrumb=strBreadCrumb & "<span itemprop=""name"">" & pcArrayCrumb(1) &"</span>"
				strBreadCrumb=strBreadCrumb & "<meta itemprop=""position"" content=""" & i+1 & """ />"
			ELSE 
				strBreadCrumb=strBreadCrumb & "<span itemprop=""itemListElement"" itemscope itemtype=""http://schema.org/ListItem"">"
				strBreadCrumb=strBreadCrumb & " > " & "<a itemprop=""item"" href=""" & pcStrBCLink & """"
				strBreadCrumb=strBreadCrumb & "><span itemprop=""name"">" & pcArrayCrumb(1) &"</span></a>"
				strBreadCrumb=strBreadCrumb & "<meta itemprop=""position"" content=""" & i+1 & """ />"
				strBreadCrumb=strBreadCrumb & "</span>"
			END IF
		end if
	next
	
ELSE
' (b)
	dim arrCategories(999,4)
	indexCategories=0
	pUrlString=Cstr("")
	pIdCategory2=pidCategory

	' load category array with all categories until parent
	do while pIdCategory2>1
		queryW=""
		queryHC=""
		if not session("customerType")="1" then
			queryW = " AND categories.pccats_RetailHide<>1"
		end if
		' If admin preview, ignore hidden categories
		if session("pcv_intAdminPreview")<>1 and session("admin") <> 0 then
			queryHC = " AND categories.iBTOhide<>1" & queryW
		end if
		query="SELECT categoryDesc, idCategory, idParentcategory, largeimage, SDesc, LDesc, HideDesc FROM categories WHERE idCategory=" & pIdCategory2 & queryHC & pcv_strTemp & "  ORDER BY priority, categoryDesc ASC"
		SET rs=Server.CreateObject("ADODB.RecordSet")
		SET rs=conntemp.execute(query)

		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
 
		if rs.eof then
			set rs=nothing
			call closeDb()
			response.redirect "msg.asp?message=86"           
		end if
		
		'categoryDesc, idCategory, idParentcategory, largeimage, SDesc, LDesc, HideDesc
		if pIdCategory2=pidCategory then
			pCatName=rs("categoryDesc")
			intIdCategory=rs("idCategory")
			intIdParentCategory=rs("idParentCategory")
			plargeImage=rs("largeimage")
			if pLargeImage = "no_image.gif" then
				pLargeImage = ""
			end if
			SDesc=rs("SDesc")
			LDesc=rs("LDesc")
			HideDesc=rs("HideDesc")
			if isNULL(HideDesc) OR HideDesc="" then
				HideDesc="0"
			end if
		else
			pCatName=rs("categoryDesc")
			intIdCategory=rs("idCategory")
			intIdParentCategory=rs("idParentCategory")
		end if
		
		pIdCategory3=intIdParentCategory 
		arrCategories(indexCategories,0)=pCatName
		arrCategories(indexCategories,1)=intIdCategory
		arrCategories(indexCategories,2)=intIdParentCategory
		pIdCategory2=pIdCategory3
		indexCategories=indexCategories + 1   
	loop
	set rs=nothing
	
	'create new breadcrumb and enter it into database
	strBreadCrumb=""
	for f=indexCategories-1 to 0 step -1
		intBCId = arrCategories(f,1)
		strBCDesc = arrCategories(f,0)

		'// Call SEO Routine
		pcGenerateSeoLinks
		'//
		
		If arrCategories(f,2)="1" Then
			strDBBreadCrumb=strDBBreadCrumb&arrCategories(f,1)&"||"&arrCategories(f,0)
			IF (f = 0) AND (request("idproduct")="") AND (request("idcategory")<>"") THEN
				strBreadCrumb=strBreadCrumb & "<span itemprop=""itemListElement"" itemscope itemtype=""http://schema.org/ListItem"">"
				strBreadCrumb=strBreadCrumb & "<span itemprop=""name"">" & arrCategories(f,0) &"</span>"
				strBreadCrumb=strBreadCrumb & "<meta itemprop=""position"" content=""" & f+1 & """ />"
			ELSE
				strBreadCrumb=strBreadCrumb & "<span itemprop=""itemListElement"" itemscope itemtype=""http://schema.org/ListItem"">"
				strBreadCrumb=strBreadCrumb & "<a itemscope itemtype=""http://schema.org/Thing"" itemprop=""item"" href=""" & pcStrBCLink & """"
				strBreadCrumb=strBreadCrumb & "><span itemprop=""name"">" & arrCategories(f,0) &"</span></a>"
				strBreadCrumb=strBreadCrumb & "<meta itemprop=""position"" content=""" & f+1 & """ />"
				strBreadCrumb=strBreadCrumb & "</span>"
			END IF
		Else
			strDBBreadCrumb=strDBBreadCrumb&"|,|"&arrCategories(f,1)&"||"&arrCategories(f,0)
			IF (f = 0) AND (request("idproduct")="") AND (request("idcategory")<>"") THEN
				strBreadCrumb=strBreadCrumb & " > "
				strBreadCrumb=strBreadCrumb & "<span itemprop=""itemListElement"" itemscope itemtype=""http://schema.org/ListItem"">"
				strBreadCrumb=strBreadCrumb & "<span itemprop=""name"">" & arrCategories(f,0) &"</span>"
				strBreadCrumb=strBreadCrumb & "<meta itemprop=""position"" content=""" & f+1 & """ />"
			ELSE
				strBreadCrumb=strBreadCrumb & " > "
				strBreadCrumb=strBreadCrumb & "<span itemprop=""itemListElement"" itemscope itemtype=""http://schema.org/ListItem"">"
				strBreadCrumb=strBreadCrumb & "<a itemscope itemtype=""http://schema.org/Thing"" itemprop=""item"" href=""" & pcStrBCLink & """"
				strBreadCrumb=strBreadCrumb & "><span itemprop=""name"">" & arrCategories(f,0) &"</span></a>"
				strBreadCrumb=strBreadCrumb & "<meta itemprop=""position"" content=""" & f+1 & """ />"
				strBreadCrumb=strBreadCrumb & "</span>"
			END IF
		End If
	next
	'enter BreadCrumb into database
	query="UPDATE categories SET pccats_BreadCrumbs=N'"&replace(strDBBreadCrumb,"'","''")&"' WHERE idCategory="&pIdCategory&";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	set rs=nothing
END IF
%>