<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<%
Dim strQ, intProductId, strServerTransfer, strPrdDetails, strCatDetails, strOther404, strQProductIdCount, strRestOfQueryString, pcIntValidPath, pcStrBrandLink2

'====================================
'== Set page location ===============
'====================================

strCatDetails = "viewCategories.asp"

'====================================
'== You should not have to edit	  ===
'== any code after this point	  ===
'====================================

' Name of product details page
strPrdDetails = "viewPrd.asp" ' Location of product details page

' Name of Brand details page
strBrandDetails = "viewBrands.asp" ' Location of brands page

' Name of content page
strCntDetails = "viewcontent.asp" ' Location of content page

' Name of parent content page
strCntParent = "viewcontent.asp" ' Location of parent content page

' Redirection to standard 404 error page is peformed by 404b.asp.
' Open and edit 404b.asp to change the name of the default 404 error page.
strOther404 = "404b.asp"

' Get the Page Name
strQ = Request.ServerVariables("QUERY_STRING")
'// Troubleshooting: show page address
strQ=replace(strQ,"404;","")
strQ=replace(strQ,":80","")
strQ=replace(strQ,":443","")
strQ=replace(strQ,":8443","")

SPath1=Request.ServerVariables("PATH_INFO")
mycount1=0
do while mycount1<1
	if mid(SPath1,len(SPath1),1)="/" then
		mycount1=mycount1+1
	end if
	if mycount1<1 then
		SPath1=mid(SPath1,1,len(SPath1)-1)
	end if
loop
if Ucase(Request.ServerVariables("HTTPS"))="ON" then
	SPathInfo="https://"
else
	SPathInfo="http://"
end if
SPathInfo=SPathInfo & Request.ServerVariables("HTTP_HOST") & SPath1


'Check for valid path
pcIntValidPath=0
if instr(Ucase(strQ),Ucase(SPathInfo))=1 then 
    pcIntValidPath=1
end if

strQORG=strQ

Function SEOcheckAff(tmpQ)
Dim tmpStr1,tmpStr2,k,tmp1,tmp2
	tmp1=Cint(1)
	if Instr(tmpQ,"?")>0 then
	tmp2=split(tmpQ,"?")
	if tmp2(1)<>"" then
		tmpStr1=split(tmp2(1),"&")
		For k=lbound(tmpStr1) to ubound(tmpStr1)
			if tmpStr1(k)<>"" then
				if Instr(Ucase(tmpStr1(k)),"IDAFFILIATE")>0 then
					tmpStr2=split(tmpStr1(k),"=")
					if tmpStr2(1)<>"" then
						if IsNumeric(tmpStr2(1)) then
							tmp1=Clng(tmpStr2(1))
						end if
					end if
				end if
			end if
		Next
	end if
	end if
	SEOcheckAff=tmp1
End Function

Function GetParam(tmpQ,tmpName)
Dim tmpStr1,tmpStr2,k,tmp1,tmp2
	tmp1=""
	if Instr(tmpQ,"?")>0 then
	tmp2=split(tmpQ,"?")
	if tmp2(1)<>"" then
		tmpStr1=split(tmp2(1),"&")
		For k=lbound(tmpStr1) to ubound(tmpStr1)
			if tmpStr1(k)<>"" then
				if Instr(Ucase(tmpStr1(k)),Ucase(tmpName))>0 then
					tmpStr2=split(tmpStr1(k),"=")
					if tmpStr2(1)<>"" then
						if IsNumeric(tmpStr2(1)) then
							tmp1=Clng(tmpStr2(1))
						end if
					end if
				end if
			end if
		Next
	end if
	end if
	GetParam=tmp1
End Function

' Find the Product, Category, or Page ID
nIndex = InStrRev(strQ,"/")
If (nIndex>0) Then
	' Look for affiliate ID, set special session variable
	session("strSEOAffiliate")=SEOcheckAff(strQ)

	' Remove last character added on refresh (BTO configuration page)
	if (InStrRev(strQ,"=",-1,1)) AND (InStrRev(Ucase(strQ),"PAGESTYLE=",-1,1)=0) then
	 strQCount=len(strQ)
	 strQtemp=left(strQ,strQCount-1)
	 strQ=strQtemp
	 if session("strSEOAffiliate")="1" then
	 	strQORG=strQ
	 end if
	end if
	strQProductId = split(strQ,"?")
	strQProductIdCount = ubound(strQProductId)
	intProductId = Right(strQProductId(0),Len(strQProductId(0))-nIndex)
		nIndex2 = InStrRev(intProductId,"-",-1,1)
		If (nIndex2>0) Then
			intProductId = Right(intProductId,Len(intProductId)-nIndex2)
		end if
	if strQProductIdCount > 0 then
		strRestOfQueryString=strQProductId(1)
		session("strSeoQueryString")=strRestOfQueryString
		else
		session("strSeoQueryString")=""
	end if
else
	Server.Transfer(strOther404)
End If

' Detect whether this is an htm page
If Instr(LCase(intProductId),".htm") = 0 Then
	Server.Transfer(strOther404)
End If

' Detect whether this is a product, a category, or a content page
If Instr(LCase(intProductId),"c") <> 0 Then

	' START - CATEGORY PAGE

		intProductId=replace(intProductId,"c","")
		' Trim Off .htm from category ID
		If	((Mid(intProductId,Len(intProductId),1)="m") And _
			(Mid(intProductId,Len(intProductId)-1,1)="t") And _
			(Mid(intProductId,Len(intProductId)-2,1)="h") And _
			(Mid(intProductId,Len(intProductId)-3,1)=".")) Then
			intProductId = Left(intProductId,Len(intProductId)-4)
			if IsNull(intProductId) or trim(intProductId)="" then
				intProductId="NA"
			end if
		End If

		' Look Up the category Id In the Database
		if isNumeric(intProductid)=True then
        
            '// Cat 0 is used for SEO check in AdminSetting.asp
            If intProductid = "0" Then
                Server.Transfer("viewcategories.asp")
            End If
            
            query = "SELECT idCategory,categoryDesc FROM categories WHERE idCategory = " & intProductId
            set rs = Server.CreateObject("ADODB.Recordset")
            Set rs = ConnTemp.Execute(query)
            If (Not rs.EOF) Then
                strServerTransfer = 1
                strCategoryDesc=rs("categoryDesc")
            else
                strServerTransfer = 0
            End If
            set rs = nothing
				
		else
		    strServerTransfer = 0
		end if

		' Go to the new page
		If strServerTransfer=1 and pcIntValidPath=1 then
			intIdCategory=intProductId

			pcGenerateSeoLinks
			if session("strSEOAffiliate")>"1" then
				tmpREURL=SPathInfo & pcStrCatLink & "?idaffiliate=" & session("strSEOAffiliate")
			else
				tmpREURL=SPathInfo & pcStrCatLink
			end if

			session("idCategoryRedirect") = intProductId
			session("idCategoryRedirectSF") = intProductId

            if instr(Ucase(strQORG), Ucase(tmpREURL))>0 then
			
				Server.Transfer(strCatDetails)
			else
				pcf_do301Redirect(tmpREURL)
			end if
		else
			Server.Transfer(strOther404)
		end if

	' END - CATEGORY PAGE
	
ElseIf Instr(LCase(intProductId),"b") <> 0 Then

	' START - BRAND PAGE

		intProductId=replace(intProductId,"b","")
		' Trim Off .htm from category ID
		If	((Mid(intProductId,Len(intProductId),1)="m") And _
			(Mid(intProductId,Len(intProductId)-1,1)="t") And _
			(Mid(intProductId,Len(intProductId)-2,1)="h") And _
			(Mid(intProductId,Len(intProductId)-3,1)=".")) Then
			intProductId = Left(intProductId,Len(intProductId)-4)
			if IsNull(intProductId) or trim(intProductId)="" then
				intProductId="NA"
			end if
		End If

		' Look Up the Brand Id In the Database
		if isNumeric(intProductid)=True then
				
            query = "SELECT idBrand, BrandName FROM Brands WHERE idBrand = " & intProductId
            set rs = Server.CreateObject("ADODB.Recordset")
            Set rs = ConnTemp.Execute(query)
            If (Not rs.EOF) Then
                pcIntIDBrand=rs("idBrand")
                BrandName=rs("BrandName")
                strServerTransfer = 1
            else
                strServerTransfer = 0
            End If
            set rs = nothing
            
        else
            strServerTransfer = 0
		end if

		' Go to the new page
		If strServerTransfer=1  and pcIntValidPath=1 then
			
            pcGenerateSeoLinks
			if session("strSEOAffiliate")>"1" then
				tmpREURL=SPathInfo & pcStrBrandLink2 & "?idaffiliate=" & session("strSEOAffiliate")
			else
				tmpREURL=SPathInfo & pcStrBrandLink2
			end if
			if InStrRev(Ucase(strQ),"PAGESTYLE=",-1,1)>0 then
				tmpREURL=tmpREURL & "?pagestyle=" & GetParam(strQORG,"pagestyle")
			end if

			session("idBrandRedirect") = intProductId
			session("idBrandRedirectSF") = intProductId

            if instr(Ucase(strQORG), Ucase(tmpREURL))>0 then
				
				Server.Transfer(strBrandDetails)
			else
				pcf_do301Redirect(tmpREURL)
			end if
		else
			Server.Transfer(strOther404)
		end if

	' END - BRAND PAGE

elseif Instr(LCase(intProductId),"d") <> 0 then ' "d" stands for "document"

	' START - CONTENT PAGE

		intProductId=replace(intProductId,"d","")
		' Trim Off .htm from content page ID
		If	((Mid(intProductId,Len(intProductId),1)="m") And _
			(Mid(intProductId,Len(intProductId)-1,1)="t") And _
			(Mid(intProductId,Len(intProductId)-2,1)="h") And _
			(Mid(intProductId,Len(intProductId)-3,1)=".")) Then
			intProductId = Left(intProductId,Len(intProductId)-4)
			if IsNull(intProductId) or trim(intProductId)="" then
				intProductId="NA"
			end if
		End If

		' Look Up the content page Id In the Database
		if isNumeric(intProductid)=True then
				
				query="SELECT pcCont_IDPage,pcCont_PageName FROM pcContents WHERE pcCont_IDPage = "  & intProductId
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=ConnTemp.execute(query)
				If (Not rs.EOF) Then
					strServerTransfer = 1
					pcIntContentPageID=rs("pcCont_IDPage")
					pcvContentPageName=rs("pcCont_PageName")
					Set rs=nothing
					pcvPageType = ""
				else
					strServerTransfer = 0
				End If
				set rs = nothing
				
			else
				strServerTransfer = 0
		end if

		' Go to the content page
		If strServerTransfer=1  and pcIntValidPath=1 then
			pcGenerateSeoLinks
			if session("strSEOAffiliate")>"1" then
				tmpREURL=SPathInfo & pcStrCntPageLink & "?idaffiliate=" & session("strSEOAffiliate")
			else
				tmpREURL=SPathInfo & pcStrCntPageLink
			end if

			session("idContentPageRedirect") = intProductId
			session("MobileURL")=""
			
            if instr(Ucase(strQORG), Ucase(tmpREURL))>0 then

				Server.Transfer(strCntDetails)
			else
				pcf_do301Redirect(tmpREURL)
			end if
		else
			Server.Transfer(strOther404)
		end if

	' END - CONTENT PAGE

elseif Instr(LCase(intProductId),"e") <> 0 then ' This handles a parent content page

	' START - PARENT CONTENT PAGE

		intProductId=replace(intProductId,"e","")
		' Trim Off .htm from content page ID
		If	((Mid(intProductId,Len(intProductId),1)="m") And _
			(Mid(intProductId,Len(intProductId)-1,1)="t") And _
			(Mid(intProductId,Len(intProductId)-2,1)="h") And _
			(Mid(intProductId,Len(intProductId)-3,1)=".")) Then
			intProductId = Left(intProductId,Len(intProductId)-4)
			if IsNull(intProductId) or trim(intProductId)="" then
				intProductId="NA"
			end if
		End If

		' Look Up the content page Id In the Database
		if isNumeric(intProductid)=True then				
				query="SELECT pcCont_IDPage,pcCont_PageName FROM pcContents WHERE pcCont_IDPage = "  & intProductId
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=ConnTemp.execute(query)
				If (Not rs.EOF) Then
					strServerTransfer = 1
					pcIntContentPageID=rs("pcCont_IDPage")
					pcvContentPageName=rs("pcCont_PageName")
					Set rs=nothing
					pcvPageType = "parent"
				else
					strServerTransfer = 0
				End If
				set rs = nothing
			else
				strServerTransfer = 0
		end if

		' Go to the content page
		If strServerTransfer=1  and pcIntValidPath=1 then
			pcGenerateSeoLinks
			if session("strSEOAffiliate")>"1" then
				tmpREURL=SPathInfo & pcStrCntPageLink & "?idaffiliate=" & session("strSEOAffiliate")
			else
				tmpREURL=SPathInfo & pcStrCntPageLink
			end if

			session("idParentContentPageRedirect") = intProductId
			session("MobileURL")=""
            
            if instr(Ucase(strQORG), Ucase(tmpREURL))>0 then

				Server.Transfer(strCntParent)
			else
				pcf_do301Redirect(tmpREURL)
			end if
		else
			Server.Transfer(strOther404)
		end if

	' END - PARENT CONTENT PAGE
	
elseif Instr(LCase(intProductId),"f") <> 0 then ' This handles SEO URL check landing page
	
	' START - SEO URL Check
	
		intProductId=replace(intProductId,"f","")
		' Trim Off .htm from category ID
		If	((Mid(intProductId,Len(intProductId),1)="m") And _
			(Mid(intProductId,Len(intProductId)-1,1)="t") And _
			(Mid(intProductId,Len(intProductId)-2,1)="h") And _
			(Mid(intProductId,Len(intProductId)-3,1)=".")) Then
			intProductId = Left(intProductId,Len(intProductId)-4)
			if IsNull(intProductId) or trim(intProductId)="" then
				intProductId="NA"
			end if
		End If
		
		' Look Up the category Id In the Database
		if isNumeric(intProductid)=True then
        
            '// Cat 0 is used for SEO check in AdminSetting.asp
            If intProductid = "0" Then
                Server.Transfer("404d.asp")
            End If
				
		else
		    strServerTransfer = 0
		end if
	
	' END - SEO URL Check
	
else
	' This is a product page
		If Instr(LCase(intProductId),"p") <> 0 Then
			'This product is with a category
			strPrdCatArry=split(intProductId,"p")
			intTempCatId=strPrdCatArry(0)
			intProductId=strPrdCatArry(1)
			if IsNull(intProductId) or trim(intProductId)="" then
				intProductId="NA"
			end if
		end if
		' Trim Off .htm from Product ID
		If	((Mid(intProductId,Len(intProductId),1)="m") And _
			(Mid(intProductId,Len(intProductId)-1,1)="t") And _
			(Mid(intProductId,Len(intProductId)-2,1)="h") And _
			(Mid(intProductId,Len(intProductId)-3,1)=".")) Then
			intProductId = Left(intProductId,Len(intProductId)-4)
		End If
		'// Troubleshooting: show Product ID

		' Look Up the Product Id In the Database
		tmpHiddenCat=0
		if isNumeric(intProductid)=True then
				
				query = "SELECT idProduct,description FROM products WHERE idProduct = " & intProductId
				set rs = Server.CreateObject("ADODB.Recordset")
				Set rs = ConnTemp.Execute(query)
				If (Not rs.EOF) Then
					strServerTransfer = 1
					pDescription=rs("description")
					set rs=nothing
					query="SELECT categories.idcategory FROM categories INNER JOIN categories_products ON categories.idcategory=categories_products.idcategory WHERE categories_products.idProduct=" & intProductId & " AND categories.iBTOhide=0;"
					set rs = Server.CreateObject("ADODB.Recordset")
					Set rs = ConnTemp.Execute(query)
					if rs.eof then
						session("intTempCatId")="0"
						tmpHiddenCat=1
					end if
					set rs=nothing
				else
					strServerTransfer = 0
				End If
				set rs = nothing
				
			else
			strServerTransfer = 0
		end if

		' Go to the new page
		If strServerTransfer=1 and pcIntValidPath=1 then
			pIdCategory="0"
			pIdProduct=intProductId
			session("idProductRedirect") = intProductId
			if tmpHiddenCat=0 then
				session("intTempCatId") = intTempCatId
				pIdCategory="" & intTempCatId
			end if
			pcGenerateSeoLinks
			if session("strSEOAffiliate")>"1" then
				tmpREURL=SPathInfo & pcStrPrdLink & "?idaffiliate=" & session("strSEOAffiliate")
			else
				tmpREURL=SPathInfo & pcStrPrdLink
			end if

            if instr(Ucase(strQORG), Ucase(tmpREURL))>0 then			
				Server.Transfer(strPrdDetails)
			else
				pcf_do301Redirect(tmpREURL)
			end if
		else
			Server.Transfer(strOther404)
		end if

end if ' End category vs product link
%>