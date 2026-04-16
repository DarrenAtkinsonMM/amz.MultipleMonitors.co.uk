<%
'//////////////////////////////////////////////////////////////////////////////////////////
'// START: pcf_converToArray
'//////////////////////////////////////////////////////////////////////////////////////////
'// Summary: 	Converts the query string into a JavaScript array
'// Params:		A search array separated by pipes
'// Returns:	A Javascript array separated by commas and encapsulated with (')  
Function pcf_converToArray(pipes)
    On Error Resume Next
    pcf_converToArray = pipes
    pcf_converToArray=replace(pcf_converToArray,"||","','")
    pcf_converToArray="'"&pcf_converToArray&"'"
End Function
'//////////////////////////////////////////////////////////////////////////////////////////
'// END: pcf_converToArray
'//////////////////////////////////////////////////////////////////////////////////////////



'//////////////////////////////////////////////////////////////////////////////////////////
'// START: pcf_CountResults & pcf_CountSearchResults
'//////////////////////////////////////////////////////////////////////////////////////////
'// Summary: 	Counts the number of products that will be returned when the filter is added
'// Params:		SearchValueId - a string of search value ids separated by two pipes
'// Returns:	Boolean  
Function pcf_CountResults(SearchValueId,GroupValues)
   ' On Error Resume Next
    
    tmpStrEx = SearchValueId '// ""
    pcv_HavingCount = 1
    if len(pcv_strCValues)>0 then		
        tmpSValues=split(pcv_strCValues,"||")
        For k=0 to ubound(tmpSValues)		
            if tmpSValues(k)<>"" AND pcf_InUse(GroupValues,tmpSValues(k))=0 then
                tmpStrEx = tmpStrEx & ","& tmpSValues(k)
                pcv_HavingCount = pcv_HavingCount + 1	
            end if
        Next	
    end if	

    queryCSF = "SELECT pcSearchFields_Products.idProduct "
    queryCSF = queryCSF & "FROM pcSearchFields_Products "
    queryCSF = queryCSF & "INNER JOIN products ON products.idProduct=pcSearchFields_Products.idProduct "
    queryCSF = queryCSF & "INNER JOIN categories_products ON products.idProduct=categories_products.idProduct "
    queryCSF = queryCSF & "WHERE pcSearchFields_Products.idSearchData in (" & tmpStrEx & ") "
    queryCSF = queryCSF & "AND categories_products.idCategory="& pcv_strCSFCatID &" AND active=-1 AND configOnly=0 AND removed=0 "
    queryCSF = queryCSF & "GROUP BY pcSearchFields_Products.idProduct "
    queryCSF = queryCSF & "HAVING COUNT(DISTINCT pcSearchFields_Products.idSearchData) = " & pcv_HavingCount

    set rsCSF=Server.CreateObject("ADODB.Recordset")  
    set rsCSF=connTemp.execute(queryCSF)
    if NOT rsCSF.eof then
        pcarray_RowCount = rsCSF.GetRows
        pcf_CountResults = UBound(pcarray_RowCount, 2) + 1
    else 
        pcf_CountResults = 0
    end if 	
    set rsCSF=nothing

End Function

Function pcf_CountSearchResults(SearchValueId,GroupValues)
    On Error Resume Next				
    tmpStrEx = ""
    if len(pcv_strCValues)>0 then
        queryCSF = "SELECT pcSearchFields_Products.idProduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData>0 "			
        tmpSValues=split(pcv_strCValues,"||")
        For k=0 to ubound(tmpSValues)		
            if tmpSValues(k)<>"" AND pcf_InUse(GroupValues,tmpSValues(k))=0 then					
                SubQuery = "SELECT pcSearchFields_Products.idProduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData = " & tmpSValues(k) & ""
                set rsSubQuery=Server.CreateObject("ADODB.Recordset")
                set rsSubQuery=connTemp.execute(SubQuery)
                If NOT rsSubQuery.eof Then
                    ProductIdArray = pcf_ColumnToArray(rsSubQuery.getRows(),0)
                    ProductIdString = Join(ProductIdArray,",")
                    tmpStrEx=tmpStrEx & " AND pcSearchFields_Products.idProduct IN "
                    tmpStrEx=tmpStrEx & "(" & ProductIdString & ")"	
                End If
                set rsSubQuery = nothing		
            end if
        Next	
        tmpStrEx = tmpStrEx & " AND pcSearchFields_Products.idSearchData=" & SearchValueId & ""
        queryCSF = queryCSF & tmpStrEx
    else
        queryCSF = "SELECT pcSearchFields_Products.idProduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData="& SearchValueId &" "
    end if		
    set rsCSF=Server.CreateObject("ADODB.Recordset")  
    set rsCSF=connTemp.execute(queryCSF)
    if err.number<>0 then
        call LogErrorToDatabase()
        set rsCSF=nothing
        call closedb()
        response.redirect "techErr.asp?err="&pcStrCustRefID
    end if
    if NOT rsCSF.eof then
        ProductIdArray = pcf_ColumnToArray(rsCSF.getRows(),0)
        ProductIdString = Join(ProductIdArray,",")
        pcv_strCSFilters = " AND (A.idProduct In ("& ProductIdString &"))"
    else 
        pcv_strCSFilters = " AND (A.idProduct In (0))"
    end if 	
    set rsCSF=nothing

    '*******************************
    ' Create Search Query
    '*******************************
    tmp_StrQuery=""
    if session("customerCategory")="" or session("customerCategory")=0 then
        If session("customerType")=1 then
            tmp_StrQuery="(A.serviceSpec<>0 AND A.pcProd_BTODefaultWPrice>="&pPriceFrom&" AND A.pcProd_BTODefaultWPrice<=" &pPriceUntil&")"
        else
            tmp_StrQuery="(A.serviceSpec<>0 AND A.pcProd_BTODefaultPrice>="&pPriceFrom&" AND A.pcProd_BTODefaultPrice<=" &pPriceUntil&")"
        end if
    else
        tmp_StrQuery="(A.serviceSpec<>0 AND A.idproduct IN (SELECT DISTINCT idproduct FROM pcBTODefaultPriceCats WHERE pcBTODefaultPriceCats.idCustomerCategory=" & session("customerCategory") & " AND pcBTODefaultPriceCats.pcBDPC_Price>="&pPriceFrom&" AND pcBTODefaultPriceCats.pcBDPC_Price<=" &pPriceUntil&"))"
    end if
    
    zSQL="cast(A.sDesc as varchar(8000)) sDesc"

    pcv_strMaxResults=SRCH_MAX
    If pcv_strMaxResults>"0" Then
        pcv_strLimitPhrase="TOP " & pcv_strMaxResults
    Else
        pcv_strLimitPhrase=""
    End If
    
    strSQL= "SELECT "& pcv_strLimitPhrase &" A.idProduct, A.sku, A.description, A.price, A.listHidden, A.listPrice, A.serviceSpec, A.bToBPrice, A.smallImageUrl, A.noprices, A.stock, A.noStock, A.pcprod_HideBTOPrice, A.pcProd_BackOrder, A.FormQuantity, A.pcProd_BackOrder, A.pcProd_BTODefaultPrice, "& zSQL &" " 
    strSQL=strSQL& "FROM products A "
    strSQL=strSQL& " WHERE (A.active=-1 AND A.removed=0 AND A.idProduct IN (" 

        '// START: Category Sub-Query
        strSQL=strSQL& "SELECT B.idProduct FROM categories_products B INNER JOIN categories C ON "
        strSQL=strSQL & "C.idCategory=B.idCategory WHERE C.iBTOhide=0 "
        if pIdCategory<>"0" then
            if (schideCategory = "1") OR (SRCH_SUBS = "1") then					
                TmpCatList=""
                call pcs_GetSubCats(pIdCategory) '// get sub cats
                TmpCatList = pIdCategory&TmpCatList
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
            strSQL=strSQL & " AND C.pccats_RetailHide=0"
        end if
        '// END: Category Sub-Query
    
    strSQL=strSQL& ") AND (" & tmp_StrQuery & " OR (A.serviceSpec=0 AND A.configOnly=0 AND A.price>="&pPriceFrom&" AND A.price<=" &pPriceUntil&")) " 
    
    if len(pSearchSKU)>0 then
        strSQL=strSQL & " AND A.sku like '%"&pSearchSKU&"%'"
    end if
    
    if pIdSupplier<>"0" then
        strSQL=strSQL & " AND A.idSupplier=" &pIdSupplier
    end if
    
    if pWithStock="-1" then
        strSQL=strSQL & " AND (A.stock>0 OR A.noStock<>0)" 
    end if
    
    if (IDBrand&""<>"") and (IDBrand&""<>"0") then
        strSQL=strSQL & " AND A.IDBrand=" & IDBrand
    end if
    
    TestWord=""
    if intExact<>"1" then
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
            if scDB="SQL" then
                pKeywords="'" & pKeywords & "'***'%[^a-zA-z0-9]" & pKeywords & "[^a-zA-z0-9]%'***'" & pKeywords & "[^a-zA-z0-9]%'***'%[^a-zA-z0-9]" & pKeywords & "'"
            else
                pKeywords="'" & pKeywords & "'***'%[!a-zA-z0-9]" & pKeywords & "[!a-zA-z0-9]%'***'" & pKeywords & "[!a-zA-z0-9]%'***'%[!a-zA-z0-9]" & pKeywords & "'"
            end if
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
                End If
                set rsSearchFields = nothing
            end if
        Next
    end if
    
    '// Category Seach Fields 
    tmpStrEx = tmpStrEx & pcv_strCSFilters
    
    IF intExact<>"1" THEN
    
        if pKeywords<>"" then
        
            strSQl=strSql & " AND ("
            
            tmpSQL="(A.details LIKE "
            tmpSQL2="(A.description LIKE "
            tmpSQL3="(A.sDesc LIKE "
            if tIncludeSKU="true" then
                tmpSQL4="(A.SKU LIKE "
            end if
            Dim Pos
            Pos=0
            For L=LBound(keywordArray) to UBound(keywordArray)
                if trim(keywordArray(L))<>"" then
                Pos=Pos+1
                if Pos>1 Then
                    tmpSQL=tmpSQL  & TestWord & " A.details LIKE "
                    tmpSQL2=tmpSQL2 & TestWord & " A.description LIKE "
                    tmpSQL3=tmpSQL3 & TestWord & " A.sDesc LIKE "
                    if tIncludeSKU="true" then
                        tmpSQL4=tmpSQL4 & TestWord & " A.SKU LIKE "
                    end if
                end if
                    tmpSQL=tmpSQL  & "'%" & trim(keywordArray(L)) & "%'"
                    tmpSQL2=tmpSQL2 & "'%" & trim(keywordArray(L)) & "%'"
                    tmpSQL3=tmpSQL3 & "'%" & trim(keywordArray(L)) & "%'"
                    if tIncludeSKU="true" then
                        tmpSQL4=tmpSQL4 & "'%" & trim(keywordArray(L)) & "%'"
                    end if
                end if
            Next
            tmpSQL=tmpSQL & ")"
            tmpSQL2=tmpSQL2 & ")"
            tmpSQL3=tmpSQL3 & ")"
            if tIncludeSKU="true" then
                tmpSQL4=tmpSQL4 & ")"
            end if
            
            strSQL=strSQL & tmpSQL
            strSQL=strSQL & " OR " & tmpSQL2
            if tIncludeSKU="true" then
                strSQL=strSQL & " OR " & tmpSQL3
                strSQL=strSQL & " OR " & tmpSQL4 & ")"
            else	
                strSQL=strSQL & " OR " & tmpSQL3 & ")"
            end if
            strSQL=strSQL& ")" & tmpStrEx
            query=strSQL & " ORDER BY " & strORD1
        else
            strSQL=strSQL& ")" & tmpStrEx
            query=strSQL & " ORDER BY " & strORD1
        end if
    
    ELSE 'Exact=1
    
        if pKeywords<>"" then
        
            strSQl=strSql & " AND ("
            
            tmpSQL="(A.details LIKE "
            tmpSQL2="(A.description LIKE "
            tmpSQL3="(A.sDesc LIKE "
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
                    if tIncludeSKU="true" then
                        tmpSQL4=tmpSQL4 & TestWord & " A.SKU LIKE "
                    end if
                end if
                    tmpSQL=tmpSQL & trim(keywordArray(L))
                    tmpSQL2=tmpSQL2 & trim(keywordArray(L))
                    tmpSQL3=tmpSQL3 & trim(keywordArray(L))
                    if tIncludeSKU="true" then
                        tmpSQL4=tmpSQL4 & trim(keywordArray(L))
                    end if
                end if
            Next
            tmpSQL=tmpSQL & ")"
            tmpSQL2=tmpSQL2 & ")"
            tmpSQL3=tmpSQL3 & ")"
            if tIncludeSKU="true" then
                tmpSQL4=tmpSQL4 & ")"
            end if
            
            strSQL=strSQL & tmpSQL
            strSQL=strSQL & " OR " & tmpSQL2
            if tIncludeSKU="true" then
                strSQL=strSQL & " OR " & tmpSQL3
                strSQL=strSQL & " OR " & tmpSQL4 & ")"
            else	
                strSQL=strSQL & " OR " & tmpSQL3 & ")"
            end if
            strSQL=strSQL& ")" & tmpStrEx
            query=strSQL & " ORDER BY " & strORD1
        else
            strSQL=strSQL& ")" & tmpStrEx
            query=strSQL & " ORDER BY " & strORD1
        end if
    END IF 'Exact

    queryCSF=query
    set rsCSF=Server.CreateObject("ADODB.Recordset")
    rsCSF.Open queryCSF, connTemp, adOpenStatic, adLockReadOnly, adCmdText
    if not rsCSF.eof then
        pcf_CountSearchResults = rsCSF.recordcount
    else
        pcf_CountSearchResults = 0
    end if
    set rsCSF=nothing
    
    tmp_StrQuery=""
    zSQL=""
    strSQL=""
    TestWord=""
    keywordArray=""
    pKeywords=""
    tmpStrEx=""
    tmpSQL=""
    tmpSQL2=""
    tmpSQL3=""
End Function
'//////////////////////////////////////////////////////////////////////////////////////////
'// END: pcf_CountResults & pcf_CountSearchResults
'//////////////////////////////////////////////////////////////////////////////////////////



'//////////////////////////////////////////////////////////////////////////////////////////
'// START: pcf_InUse
'//////////////////////////////////////////////////////////////////////////////////////////
'// Summary: 	Determines a specific value is contained within an array
'// Params:		A string of value separated by pipes
'// Returns:	Boolean 
Function pcf_InUse(theString,theValue)
    Dim r, tmpSValuesInUse
    pcf_InUse = 0		
    If instr(theString,"||")>0 Then			
        tmpSValuesInUse=split(theString,"||")
        For r=lbound(tmpSValuesInUse) to ubound(tmpSValuesInUse)
            if tmpSValuesInUse(r)<>"" then
                if cdbl(tmpSValuesInUse(r)) = cdbl(theValue) then
                    pcf_InUse = 1
                    Exit For
                end if
            end if
        Next
    Else
        if theString<>"" AND theValue<>"" then
            if cdbl(theString) = cdbl(theValue) then
                pcf_InUse = 1
            end if	
        end if
    End If
End Function
'//////////////////////////////////////////////////////////////////////////////////////////
'// END: pcf_InUse
'//////////////////////////////////////////////////////////////////////////////////////////



'//////////////////////////////////////////////////////////////////////////////////////////
'// START: pcf_SanitizeJava
'//////////////////////////////////////////////////////////////////////////////////////////
'// Summary: 	Escape Apostrophees in JavaScript
'// Params:		A string
'// Returns:	A safe string 
Function pcf_SanitizeJava(theSting)
    pcf_SanitizeJava = replace(theSting,"'","\'")	
End Function
'//////////////////////////////////////////////////////////////////////////////////////////
'// END: pcf_SanitizeJava
'//////////////////////////////////////////////////////////////////////////////////////////



'//////////////////////////////////////////////////////////////////////////////////////////
'// START: pcf_CSFieldQuery
'//////////////////////////////////////////////////////////////////////////////////////////
'// Summary: 	Creates all the parameters needed for the query string
'// Params:		Request variables for query string
'// Returns:	A new query string 
Function pcf_CSFieldQuery()
    pcf_CSFieldQuery = "&SFID="& pcv_intSFID &"&SFNAME="& Server.URLEncode(pcv_strSFNAME) &"&SFVID="& pcv_strCValues &"&SFVALUE="& Server.URLEncode(pcv_strSFVALUE) &"&SFCount="& pcv_intSFCount
End Function
'//////////////////////////////////////////////////////////////////////////////////////////
'// END: pcf_SanitizeJava
'//////////////////////////////////////////////////////////////////////////////////////////



'//////////////////////////////////////////////////////////////////////////////////////////
'// START: pcs_CSFSetVariables
'//////////////////////////////////////////////////////////////////////////////////////////
'// Summary: 	Sets all the needed variables from the query string
'// Params:		Request variables from query string
'// Returns:	All the needed variables 
Sub pcs_CSFSetVariables()
    pcv_intSFID = getUserInput(Request("SFID"),0)
    pcv_strSFNAME = getUserInput(Request("SFNAME"),0)
    pcv_strSFNAME = replace(pcv_strSFNAME,"''","'")
    pcv_strCValues = getUserInput(Request("SFVID"),0)
    pcv_strSFVALUE = getUserInput(Request("SFVALUE"),0)
    pcv_strSFVALUE = replace(pcv_strSFVALUE,"''","'")	
    pcv_intSFCount = getUserInput(Request("SFCount"),10)
    if not validNum(pcv_intSFCount) then
        pcv_intSFCount=-1
    end if
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////
'// END: pcf_SanitizeJava
'//////////////////////////////////////////////////////////////////////////////////////////
%>


<%
'//////////////////////////////////////////////////////////////////////////////////////////
'// START: pcs_SolrCatalog
'//////////////////////////////////////////////////////////////////////////////////////////
Public Function pcs_SolrCatalog(idCategory)

    query = Request.ServerVariables("QUERY_STRING")

    'endpoint = "http://localhost:14211/SolrViews/Details/" & scSearch_Handle & "?idCategory=" & idCategory & query
    endpoint = "http://localhost:14211/SolrViews/Details/" & scSearch_Handle & "?" & query
    'pcv_strUrl = "http://service.productcartlive.com/v1/SolrViews/Details/"
    
    'If scSeoURLs = 1 Then
    '    idCategory = session("idCategoryRedirectSF")
    '    If instr(query,"404;")>0 Then
    '        If instr(query,"?")>0 Then
    '            query = Right(query, Len(query) - InStr(query, "?"))
    '        End If
    '    End If
    '    endpoint = pcv_strUrl & scSearch_Handle & "?idCategory=" & idCategory & "&" & query
    'Else
    '    endpoint = pcv_strUrl & scSearch_Handle & "?" & query
    'End IF
    endpoint = replace(endpoint,"[]","")
    endpoint = replace(endpoint,"%5B%5D","")
    
    'response.Write(endpoint)
    'response.End()

    '// START: POST
    Set objXMLhttp = Server.CreateObject("MSXML2.serverXMLHTTP" & scXML) 
    objXMLhttp.open "GET", endpoint, false
    objXMLhttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objXMLhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    objXMLhttp.send query
    cfuResult = objXMLhttp.responseText
    
    'response.Write(cfuResult)
    'response.End()
    
    Set objXMLhttp = Nothing
    '// END: POST

    pcs_SolrCatalog = trim(cfuResult)

End Function
'//////////////////////////////////////////////////////////////////////////////////////////
'// END: pcs_SolrCatalog
'//////////////////////////////////////////////////////////////////////////////////////////
%>


<%
'//////////////////////////////////////////////////////////////////////////////////////////
'// START: pcf_FindHiddenCatList
'//////////////////////////////////////////////////////////////////////////////////////////
Function pcf_FindHiddenCatList(iHide,rHide)
    Dim queryQ,rsQ,tmpArrQ,intCountQ,iQ,tmpList,jQ,tmpH
    
    tmpList=""
    tmpH=0
    
    queryQ="SELECT idCategory,idParentCategory,iBTOhide,pccats_RetailHide,0 FROM categories ORDER BY idCategory ASC;"
    set rsQ=connTemp.execute(queryQ)
    
    if not rsQ.eof then
        tmpArrQ=rsQ.getRows()
        intCountQ=ubound(tmpArrQ,2)
        set rsQ=nothing
        
        For iQ=0 to intCountQ
            if ((iHide=1) AND (tmpArrQ(3,iQ)="1")) OR ((rHide=1) AND (tmpArrQ(2,iQ)="1")) OR (tmpArrQ(4,iQ)="1") then
                tmpArrQ(4,iQ)=1
                tmpH=1
                For jQ=iQ+1 to intCountQ
                    if tmpArrQ(1,jQ)=tmpArrQ(0,iQ) then
                        tmpArrQ(4,jQ)=1
                        tmpH=1
                    end if
                Next
            end if
        Next
    end if
    set rsQ=nothing
    
    if tmpH=1 then
        For iQ=0 to intCountQ
            if (tmpArrQ(4,iQ)="1") then
                if tmpList<>"" then
                    tmpList=tmpList & ","
                end if
                tmpList=tmpList & tmpArrQ(0,iQ)
            end if
        Next
    end if
    
    pcf_FindHiddenCatList=tmpList	

End Function
'//////////////////////////////////////////////////////////////////////////////////////////
'// END: pcf_FindHiddenCatList
'//////////////////////////////////////////////////////////////////////////////////////////
%>