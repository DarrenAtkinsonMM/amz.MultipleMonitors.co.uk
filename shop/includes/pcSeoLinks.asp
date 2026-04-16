<%
dim pcStrCategoryDesc, pcStrCatLink, pcStrFeaturedCatLink, pcStrPrdLink, pcStrBCLink, intBCId, strBCDesc, pcv_strPreviousPageDesc, pcv_strNextPageDesc, pcv_strPreviousPage, pcv_strNextPage, pcStrPrdPreLink, pcStrPrdNextLink, pIdSeoCat, pIdCategoryTemp, pcStrPrdCSLink, pidrelation, pcStrCntPageLink, pcvContentPageName, pcIntContentPageID, pcStrPrdLinkCan

Public Sub pcGenerateSeoLinks

	'//=====================================
	'// CATEGORY LINKS - START
	'//=====================================

		if intIdCategory="1" then
			pcStrCatLink="viewcategories.asp"
		else
			if pcStrCategoryDesc = "" then
				pcStrCategoryDesc = strCategoryDesc
			end if
			pcStrCatLink=pcStrCategoryDesc & "-c" & intIdCategory & ".htm"
			pcStrCatLink=removeChars(pcStrCatLink)
		end if
		if scSeoURLs<>1 then
			pcStrCatLink="viewCategories.asp?idCategory="&intIdCategory
			pcStrCatLink2="viewCategories.asp"
		end if
		
			
		'Build Featured Category Link
		pcStrFeaturedCatLink=pcStrCategoryDesc & "-c" & pFeaturedCategory & ".htm"
		pcStrFeaturedCatLink=removeChars(pcStrFeaturedCatLink)
		if scSeoURLs<>1 then
			pcStrFeaturedCatLink="viewCategories.asp?idCategory=" & pFeaturedCategory
		end if
	'//=====================================
	'// CATEGORY LINKS - END
	'//=====================================
	
	'//=====================================
	'// BRAND LINKS - START
	'//=====================================

		if (len(pcIntBrandID)=0) AND (len(pcIntIDBrand)=0) then
			pcStrBrandLink="viewBrands.asp"
		else
            pcStrBrandLink=parentBrandName & "-b" & pcIntBrandID & ".htm"
            pcStrBrandLink=removeChars(pcStrBrandLink)
            pcStrBrandLink2=BrandName  & "-b" & pcIntIDBrand & ".htm"
            pcStrBrandLink2=removeChars(pcStrBrandLink2)
		end if
        
		if scSeoURLs<>1 then
			pcStrBrandLink="viewBrands.asp?idbrand=" & pcIntBrandID
			pcStrBrandLink2="viewBrands.asp?idbrand=" & pcIntIDBrand
		end if
	'//=====================================
	'// BRAND LINKS - END
	'//=====================================

	
	'//=====================================
	'// PRODUCT LINKS - START
	'//=====================================

		if pIdCategory<>"" and pIdCategory<>"0" then
			pIdSeoCat=pIdCategory
		else
			pIdSeoCat=pIdCategoryTemp		
		end if
		
		tmpDescription = trim(pDescription)
		tmpDescription = RegExpFilter(removeChars(tmpDescription))
        
        'Build Basic Product Link
		pcStrPrdLink = pcGenerateSeoProductLink(tmpDescription, pIdSeoCat, pIdProduct)
		
		'Build Canonical URL Link
		'Since the same product could be assigned to multiple categories, it makes sense not to include the category in the Canonical URL
		pcStrPrdLinkCan=tmpDescription & "-p" & pIdProduct & ".htm"
		pcStrPrdLinkCan=removeChars(pcStrPrdLinkCan)
		if scSeoURLs<>1 then
			if lcase(pcStrPageName)="configureprd.asp" then
				pcStrPrdLinkCan="configurePrd.asp?idproduct="&pIdProduct
			else
				pcStrPrdLinkCan="viewPrd.asp?idproduct="&pIdProduct
			end if
		end if
	
		'Build BreadCrumbs Link
		pcStrBCLink=strBCDesc & "-c" & intBCId & ".htm"
		pcStrBCLink=removeChars(pcStrBCLink)
		if scSeoURLs<>1 then
			pcStrBCLink="viewCategories.asp?idCategory="&intBCId
		end if		
		
		'Build Previous Product Link
		if pcv_strPreviousPageDesc<>"" then
			tmpPreDescription = trim(pcv_strPreviousPageDesc)
			tmpPreDescription = RegExpFilter(removeChars(tmpPreDescription))
			pcStrPrdPreLink = pcGenerateSeoProductLink(tmpPreDescription, pIdCategory, pcv_strPreviousPage)	
			if scSeoURLs<>1 then
				pcStrPrdPreLink="viewPrd.asp?idproduct="&pcv_strPreviousPage&"&idcategory="&pIdCategory
			end if
		end if
		
		'Build Next Product Link
		if pcv_strNextPageDesc<>"" then
			tmpNextDescription = trim(pcv_strNextPageDesc)
			tmpNextDescription = RegExpFilter(removeChars(tmpNextDescription))
			pcStrPrdNextLink = pcGenerateSeoProductLink(tmpNextDescription, pIdCategory, pcv_strNextPage)	
			if scSeoURLs<>1 then
				pcStrPrdNextLink="viewPrd.asp?idproduct="&pcv_strNextPage&"&idcategory="&pIdCategory
			end if
		end if
		
		'Build Cross Selling Product Link
		pcStrPrdCSLink=tmpDescription & "-" & pIdCategoryTemp & "p" & pidrelation & ".htm"
		pcStrPrdCSLink=removeChars(pcStrPrdCSLink)
		if scSeoURLs<>1 then
			pcStrPrdCSLink="viewPrd.asp?idproduct=" & pidrelation
		end if
	
	'//=====================================
	'// PRODUCT LINKS - END
	'//=====================================

	
	'//=====================================
	'// CONTENT LINKS - START
	'//=====================================
		if pcvContentPageName="" then pcvContentPageName=pcv_PageNameH
		if pcvPageType = "parent" then
			pcStrCntPageLink=pcvContentPageName & "-e" & pcIntContentPageID & ".htm"
			pcStrCntPageLink=removeChars(pcStrCntPageLink)
			if scSeoURLs<>1 then
				pcStrCntPageLink="viewcontent.asp?idpage=" & pcIntContentPageID
			end if
		else
			pcStrCntPageLink=pcvContentPageName & "-d" & pcIntContentPageID & ".htm"
			pcStrCntPageLink=removeChars(pcStrCntPageLink)
			if scSeoURLs<>1 then
				pcStrCntPageLink="viewcontent.asp?idpage=" & pcIntContentPageID
			end if
		end if
	'//=====================================
	'// PRODUCT LINKS - END
	'//=====================================

End Sub


Function pcGenerateSeoProductLink(pDescription, pIdSeoCat, pIdProduct)
    pcStrPrdLink = pDescription & "-" & pIdSeoCat & "p" & pIdProduct & ".htm"
	pcStrPrdLink = removeChars(pcStrPrdLink)
    If scSeoURLs <> 1 Then
        pcStrPrdLink = "viewPrd.asp?idproduct="&pIdProduct&"&idcategory="&pIdCategory
    End If
    pcGenerateSeoProductLink = pcStrPrdLink
End Function
%>
<!--#include file="pcSeoFunctions.asp"-->