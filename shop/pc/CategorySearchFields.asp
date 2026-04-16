<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
Dim pcv_intSFID, pcv_strSFNAME, pcv_strCValues, pcv_strSFVALUE, pcv_intSFCount
Dim pcv_strCSFilters, pcv_strCSFieldQuery

call openDb()

pcv_CurrentPageName = lcase(Session("pcStrPageName"))

'//////////////////////////////////////////////////////////////////////////////////////////////
'// START: CATEGORY SEARCH FIELDS
'//////////////////////////////////////////////////////////////////////////////////////////////
IF (pcv_CurrentPageName = "showsearchresults.asp" AND SRCH_CSFRON="1") OR (pcv_CurrentPageName = "viewcategories.asp" AND SRCH_CSFON="1") OR (instr(Request.ServerVariables("HTTP_REFERER"),"-c")>0 AND SRCH_CSFON="1") THEN

    '//////////////////////////////////////////////////////////////////////////////////////////
    '// START: Get Widget Query Parameters
    '//////////////////////////////////////////////////////////////////////////////////////////
	Dim pcPageStyleCSF
    pcPageStyleCSF = LCase(getUserInput(Request("pageStyle"),1))

	'// Check querystring saved to session by 404.asp
	if pcPageStyleCSF = "" then
		strSeoQueryString=lcase(session("strSeoQueryString"))
		if strSeoQueryString<>"" then
			if InStr(strSeoQueryString,"pagestyle")>0 then
				pcPageStyleCSF=left(replace(strSeoQueryString,"pagestyle=",""),1)
			end if
		end if
	end if
	
    if pcPageStyleCSF = "" then
		if Session("pStrPageStyle")<>"" then
        	pcPageStyleCSF = Session("pStrPageStyle")
			Session("pStrPageStyle")=""
		end if
    end if
    if isNULL(pcPageStyleCSF) OR trim(pcPageStyleCSF) = "" then
        pcPageStyleCSF = LCase(bType)
    end if
    if pcPageStyleCSF <> "h" and pcPageStyleCSF <> "l" and pcPageStyleCSF <> "m" and pcPageStyleCSF <> "p" then
        pcPageStyleCSF = LCase(bType)
    end if	

	'// SEO-START
	pcv_strCSFCatID=session("idCategoryRedirectSF")
	session("idCategoryRedirectSF")=""
	if pcv_strCSFCatID = "" then
		pcv_strCSFCatID=getUserInput(request("idCategory"),10)
	end if
	'// SEO-END
	
    if not validNum(pcv_strCSFCatID) then
        pcv_strCSFCatID=""
    end if
    pcv_strPage = getUserInput(Request("page"),10)
    if not validNum(pcv_strPage) then
        pcv_strPage=0
    end if
    
	If pcv_CurrentPageName = "showSearchResults.asp" Then
	
		SearchValues=getUserInput(Request("SearchValues"),0)
		pIdSupplier=getUserInput(request.querystring("idSupplier"),4)
		pPriceFrom=getUserInput(request.querystring("priceFrom"),20)
		pPriceUntil=getUserInput(request.querystring("priceUntil"),20)
		pSearchSKU=getUserInput(request.querystring("SKU"),150)
		IDBrand=getUserInput(request.querystring("IDBrand"),20)
		pKeywords=getUserInput(request.querystring("keyWord"),100)
		pcustomfield=getUserInput(request.querystring("customfield"),0)
		iPageSize=getUserInput(request("resultCnt"),10)
		strPrdOrd=getUserInput(request.querystring("order"),4)
		iPageCurrent=getUserInput(request.querystring("iPageCurrent"),4)
		pIdCategory=getUserInput(request.querystring("idCategory"),4)
		if NOT validNum(pIdCategory) or trim(pIdCategory)="" then
			pIdCategory=0
		end if
		tKeywords=pKeywords
		tIncludeSKU=getUserInput(request.querystring("includeSKU"),10)
		if tIncludeSKU = "" then
			tIncludeSKU = "true"
		end if
		
		if Instr(pPriceFrom,",")>Instr(pPriceFrom,".") then
			pPriceFrom=replace(pPriceFrom,",",".")
		end if
		if NOT isNumeric(pPriceFrom) then
			pPriceFrom=0
		end if
		if Instr(pPriceUntil,",")>Instr(pPriceUntil,".") then
			pPriceUntil=replace(pPriceUntil,",",".")
		end if
		if NOT isNumeric(pPriceUntil) then
			pPriceUntil=999999999
		end if
		if NOT validNum(pIdSupplier) or trim(pIdSupplier)="" then
			pIdSupplier=0
		end if
		pWithStock=getUserInput(request.querystring("withStock"),2)
		if NOT validNum(IDBrand) or trim(IDBrand)="" then
			IDBrand=0
		end if
		
		if NOT validNum(strPrdOrd) or trim(strPrdOrd)="" then
			strPrdOrd=3
		end if
		Select Case strPrdOrd
			Case "1": strORD1="A.idproduct ASC"
			Case "2": strORD1="A.sku ASC, A.idproduct DESC"
			Case "3": strORD1="A.description ASC"
			Case "4":
				If Session("customerType")=1 then
					if Ucase(scDB)="SQL" then
						strORD1 = "(CASE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE A.bToBPrice WHEN 0 THEN A.Price ELSE A.bToBPrice END) ELSE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN A.pcProd_BTODefaultPrice ELSE A.pcProd_BTODefaultWPrice END) END) ASC"
					else
						strORD1 = "(iif(iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),iif(IsNull(A.pcProd_BTODefaultPrice),0,A.pcProd_BTODefaultPrice),A.pcProd_BTODefaultWPrice)=0,iif(A.btoBPrice=0,A.Price,A.btoBPrice),iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),A.pcProd_BTODefaultPrice,A.pcProd_BTODefaultWPrice))) ASC"
					end if
				else
					if Ucase(scDB)="SQL" then
						strORD1 = "(CASE (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) WHEN 0 THEN A.Price ELSE A.pcProd_BTODefaultPrice END) ASC"
					else
						strORD1 = "(iif((A.pcProd_BTODefaultPrice=0) OR (IsNull(A.pcProd_BTODefaultPrice)),A.Price,A.pcProd_BTODefaultPrice)) ASC"
					end if
				End if
			Case "5": 
				If Session("customerType")=1 then
					if Ucase(scDB)="SQL" then
						strORD1 = "(CASE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE A.bToBPrice WHEN 0 THEN A.Price ELSE A.bToBPrice END) ELSE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN A.pcProd_BTODefaultPrice ELSE A.pcProd_BTODefaultWPrice END) END) DESC"
					else
						strORD1 = "(iif(iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),iif(IsNull(A.pcProd_BTODefaultPrice),0,A.pcProd_BTODefaultPrice),A.pcProd_BTODefaultWPrice)=0,iif(A.btoBPrice=0,A.Price,A.btoBPrice),iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),A.pcProd_BTODefaultPrice,A.pcProd_BTODefaultWPrice))) DESC"
					end if
				else
					if Ucase(scDB)="SQL" then
						strORD1 = "(CASE (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) WHEN 0 THEN A.Price ELSE A.pcProd_BTODefaultPrice END) DESC"
					else
						strORD1 = "(iif((A.pcProd_BTODefaultPrice=0) OR (IsNull(A.pcProd_BTODefaultPrice)),A.Price,A.pcProd_BTODefaultPrice)) DESC"
					end if
				End if
		End Select
		strORD=strPrdOrd
		
		intExact=getUserInput(request.querystring("exact"),4)
		if NOT validNum(intExact) or trim(intExact)="" then
			intExact=0
		end if
		
	End If '// If pcv_CurrentPageName = "search.asp" Then
    
    pcs_CSFSetVariables()
    Session("pcv_strCSFieldQuery") = pcf_CSFieldQuery()    
    '//////////////////////////////////////////////////////////////////////////////////////////
    '// END: Get Widget Query Parameters
    '//////////////////////////////////////////////////////////////////////////////////////////
    
    
    
    '//////////////////////////////////////////////////////////////////////////////////////////
    '// START: Disply Widget 
    '//////////////////////////////////////////////////////////////////////////////////////////
    
    If pcv_strCSFCatID<>"" Then  
	
		'// Get the list of products that are currently available
		'// Only run this block if the include is in the header
		if len(pcv_strCValues)>0 AND len(pcv_strCSFilters)=0 then
		
			tmpStrEx3=""
			pcv_HavingCount2 = 0
			tmpSValues3=split(pcv_strCValues,"||")
			For k=lbound(tmpSValues3) to ubound(tmpSValues3)	
				if tmpSValues3(k)<>"" then
					if pcv_HavingCount2=0 then
						tmpStrEx3 = tmpStrEx3 & ""& tmpSValues3(k)
					else
						tmpStrEx3 = tmpStrEx3 & ","& tmpSValues3(k)
					end if 					
					pcv_HavingCount2 = pcv_HavingCount2 + 1										
				end if
			Next
			
			queryCSF = "SELECT pcSearchFields_Products.idProduct "
			queryCSF = queryCSF & "FROM pcSearchFields_Products "
			queryCSF = queryCSF & "INNER JOIN products ON products.idProduct=pcSearchFields_Products.idProduct "
			queryCSF = queryCSF & "INNER JOIN categories_products ON products.idProduct=categories_products.idProduct "
			queryCSF = queryCSF & "WHERE pcSearchFields_Products.idSearchData in (" & tmpStrEx3 & ") "
			queryCSF = queryCSF & "AND categories_products.idCategory="& pcv_strCSFCatID &" AND active=-1 AND configOnly=0 AND removed=0 "
			queryCSF = queryCSF & "GROUP BY pcSearchFields_Products.idProduct "
			queryCSF = queryCSF & "HAVING COUNT(DISTINCT pcSearchFields_Products.idSearchData) = " & pcv_HavingCount2

			set rsCSF=Server.CreateObject("ADODB.Recordset")  
			set rsCSF=connTemp.execute(queryCSF)
			if NOT rsCSF.eof then
				ProductIdArray = pcf_ColumnToArray(rsCSF.getRows(),0)
				CartProductIdString = Join(ProductIdArray,",")
				pcv_strCSFilters = " AND (products.idProduct In ("& CartProductIdString &"))"
			else 
				pcv_strCSFilters = " AND (products.idProduct In (0))"
			end if
			set rsCSF = nothing

		end if

		pcv_strTmpCatID = pcv_strCSFCatID
		
		TmpCatList=""
		If pcv_CurrentPageName = "showSearchResults.asp" Then
		  if pIdCategory<>"0" then
			  if (schideCategory = "1") OR (SRCH_SUBS = "1") then	
			  	  TmpCatList=""				
				  call pcs_GetSubCats(pIdCategory) '// get sub cats
				  TmpCatList=pIdCategory&TmpCatList
				  pcv_strTmpCatID = TmpCatList
			  end if
		  end if
		End If
		
		query="SELECT DISTINCT idSearchData from pcSearchFields_Categories where idCategory IN (" & pcv_strTmpCatID & ") "
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=connTemp.execute(query)
        pcv_strSearchDataIDs = ""
        do while not rs.eof
            pcv_strSearchDataIDs=pcv_strSearchDataIDs & rs("idSearchData") & ","
            rs.movenext
        loop 
        set rs=nothing
        
        If pcv_strSearchDataIDs<>"" Then
        %>    
        <input type="hidden" name="customfield" id="customfield" value="0">  
        <input type="hidden" name="SearchValues" id="SearchValues" value="">   
        <%
        pcv_strTmpPageName = lcase(Session("pcStrPageName"))
        If NOT len(pcv_strTmpPageName)>0 OR instr(pcv_strTmpPageName,"404.asp")<>0  Then ' // SEO EDIT
            pcv_strTmpPageName="viewCategories.asp"
        End If
        %>
        <form name="CSF" id="CSF" action="<%=pcv_strTmpPageName%>" method="get"> 
        <div id="pcCSF">		
		<h4><%=dictLanguage.Item(Session("language")&"_categorysearchfield_1")%></h4>
		<span id="notice" name="notice" class="pcCSFNotice" style="display:none"><%=dictLanguage.Item(Session("language")&"_categorysearchfield_4")%></span>
        <span id="stable" name="stable" class="pcCSFStable"></span>
		
            <% If pcv_CurrentPageName = "showSearchResults.asp" Then %>
                <%' Required Search Parameters %>  
                <input type="hidden" name="SKU" id="SKU" value="<%=pSearchSKU%>"> 
                <input type="hidden" name="keyWord" id="keyWord" value="<%=pKeywords%>"> 
                <input type="hidden" name="SearchValues" id="SearchValues" value="<%=SearchValues%>"> 
                <input type="hidden" name="includeSKU" id="includeSKU" value="<%=tIncludeSKU%>"> 
                <input type="hidden" name="priceFrom" id="priceFrom" value="<%=pPriceFrom%>">
                <input type="hidden" name="priceUntil" id="priceUntil" value="<%=pPriceUntil%>">
                <input type="hidden" name="idSupplier" id="idSupplier" value="<%=pIdSupplier%>">
                <input type="hidden" name="withStock" id="withStock" value="<%=pWithStock%>">
                <input type="hidden" name="customfield" id="customfield" value="<%=pcustomfield%>">
                <input type="hidden" name="IDBrand" id="IDBrand" value="<%=IDBrand%>">
                <input type="hidden" name="order" id="order" value="<%=strPrdOrd%>">
                <input type="hidden" name="exact" id="exact" value="<%=intExact%>">
                <input type="hidden" name="resultCnt" id="resultCnt" value="<%=iPageSize%>">
                <input type="hidden" name="iPageSize" id="iPageSize" value="<%=iPageSize%>">
                <input type="hidden" name="iPageCurrent" id="iPageCurrent" value="<%=iPageCurrent%>">
    		<% end if %>
            
            <%' Required Category Parameters %>
            <input type="hidden" name="SFID" id="SFID" value="<%=pcv_intSFID%>">
            <input type="hidden" name="SFNAME" id="SFNAME" value="<%=pcv_strSFNAME%>">
            <input type="hidden" name="SFVID" id="SFVID" value="<%=pcv_strCValues%>">
            <input type="hidden" name="SFVALUE" id="SFVALUE" value="<%=pcv_strCValues%>">
            <input type="hidden" name="SFCount" id="SFCount" value="<%=pcv_intSFCount%>">         
            <input type="hidden" name="page" id="page" value="<%=pcv_strPage%>">
            <input type="hidden" name="pageStyle" id="pageStyle" value="<%=pcPageStyleCSF%>"> 
            <input type="hidden" name="idcategory" id="idcategory" value="<%=pcv_strCSFCatID%>"> 
	       	<div id="multiAccordion">		
            <%
            pcv_strSearchDataIDs=left(pcv_strSearchDataIDs,len(pcv_strSearchDataIDs)-1)
			query="SELECT DISTINCT idSearchField, pcSearchFieldName, pcSearchFieldShow, pcSearchFieldOrder "
			query=query&"FROM pcSearchFields "
			query=query&"WHERE idSearchField IN ("& pcv_strSearchDataIDs &") "
			query=query&"ORDER BY pcSearchFieldOrder ASC, pcSearchFieldName ASC;"
            set rs=connTemp.execute(query)
			if not rs.eof then			
                pcArray=rs.getRows()
                intCount=ubound(pcArray,2)
                set rs=nothing	
                pcv_ClearAll = 0
				pcv_strJQVars = 0
				pcv_intJQCount = 0
                For i=0 to intCount
                    pcv_strCatSearchName = pcArray(1,i)  
                    pcv_strCatSearchID = pcArray(0,i)
                    pcv_strFullView=Request("VS"&pcv_strCatSearchID) ' note: getUserInput cannot be used here
                    If NOT validNum(pcv_strFullView) Then
                        pcv_strFullView=0
                    End If
					If len(CartProductIdString)>0 Then
						query="SELECT DISTINCT pcSearchData.idSearchData, pcSearchData.pcSearchDataName, pcSearchData.idSearchField, pcSearchData.pcSearchDataOrder "
						query=query&"FROM pcSearchData "
						query=query&"INNER JOIN pcSearchFields_Products ON pcSearchFields_Products.idSearchData = pcSearchData.idSearchData "
						query=query&"WHERE pcSearchFields_Products.idProduct IN ("& CartProductIdString &") "
						query=query&"AND pcSearchData.idSearchField=" & pcArray(0,i) & " "
						query=query&"ORDER BY pcSearchDataOrder ASC,pcSearchDataName ASC;"
					Else
						query="SELECT DISTINCT A.idSearchData, A.pcSearchDataName, A.idSearchField, A.pcSearchDataOrder "
						query=query&"FROM pcSearchData A "
						query=query&"LEFT JOIN pcSearchFields_Products B on B.idSearchData=A.idSearchData "
						query=query&"LEFT JOIN  "
						query=query&"	( "
						query=query&"		SELECT D.idProduct, D.idCategory "
						query=query&"		FROM categories_products D "
						query=query&"		LEFT JOIN products E ON E.idProduct=D.idProduct "
						query=query&"		WHERE E.active=-1 AND E.configOnly=0 and E.removed=0 "
						query=query&"	) C on C.idProduct=B.idProduct "
						query=query&"WHERE C.idCategory in ("& pcv_strTmpCatID &") AND idSearchField=" & pcArray(0,i) & " "
						query=query&"ORDER BY A.pcSearchDataOrder ASC, A.pcSearchDataName ASC;"
					End If
                    set rs=Server.CreateObject("ADODB.Recordset")
                    set rs=connTemp.execute(query)
                    if not rs.eof then
                        tmpArr=rs.getRows()
                        LCount=ubound(tmpArr,2) 
                        set rs=nothing				
                        pcv_strGroupValues = pcf_ColumnToArray(tmpArr,0)
                        pcv_strGroupValues = Join(pcv_strGroupValues,"||") 
						if LCount>=0 then 
						pcv_intJQCount=pcv_intJQCount+1
                        %>
                                    <%
                                    Dim pcv_strNotInUse
                                    pcv_strInUse = 0
                                    pcv_CountGroup = 0
                                   	
                                    For j=0 to LCount
                                        pcv_strInUse = 0
                                        pcv_strDataID = tmpArr(0,j)
                                        pcv_strSFID = tmpArr(2,j)
                                        pcv_strCatSearchDataName = tmpArr(1,j)
										pcv_strInUse = pcf_InUse(pcv_strCValues,pcv_strDataID)	
                                        
                                         
										If pcv_strInUse=0 Then                                        
                                            if pcv_CurrentPageName = "showSearchResults.asp" then
												pcv_intPrdCount = pcf_CountSearchResults(pcv_strDataID, pcv_strGroupValues)
											else
												pcv_intPrdCount = pcf_CountResults(pcv_strDataID, pcv_strGroupValues)										
											end if	
										
                                            If pcv_intPrdCount>0 Then
                                                pcv_strSearchValue = pcv_strCValues & pcv_strDataID & "||"
                                                pcv_strCatSearchName = replace(pcv_strCatSearchName,"""","&quot;")
                                                pcv_strCatSearchNameTmp = pcf_SanitizeJava(pcv_strCatSearchName)											
                                                pcv_strCatSearchDataName = replace(pcv_strCatSearchDataName,"""","&quot;")
                                                pcv_strCatSearchDataNameTmp = pcf_SanitizeJava(pcv_strCatSearchDataName)
												pcv_ClearAll = pcv_ClearAll + 1 '// Count total filters
                                                pcv_CountGroup = pcv_CountGroup + 1 '// Count filters in group
												if pcv_CountGroup=1 then%>
												<h5><%=pcv_strCatSearchName%></h5>
					                            <div>
												<input type="hidden" name="VS<%=pcv_strCatSearchID%>" id="VS<%=pcv_strCatSearchID%>" value="<%=pcv_strFullView%>">
												<%end if%>
                                                <div class="pcCSFItem">
                                                    <a href="javascript:AddSF('<%=pcArray(0,i)%>','<%=pcv_strCatSearchNameTmp%>','<%=pcv_strDataID%>','<%=pcv_strCatSearchDataNameTmp%>',0);">
                                                        <img src="<%=pcf_getImagePath("../pc/images","plus.jpg")%>" alt="Add" border="0"> <%=pcv_strCatSearchDataName%>
                                                    </a>
                                                    <span class="pcCSFCount">(<%=pcv_intPrdCount%>)</span>
                                                </div>
                                                <%										
                                               
                                            End If
                                        End If
                                        if pcv_strFullView = 0 then
                                            if pcv_CountGroup > 5 then
                                                %>
                                                <div class="pcCSFItem">
                                                    <a href="javascript:ShowMore(<%=pcv_strCatSearchID%>)"><strong><%=dictLanguage.Item(Session("language")&"_categorysearchfield_2")%></strong></a>
                                                </div>
                                                <%
                                                exit for
                                            end if
                                        end if
                                    Next
									if pcv_strFullView = 1 then
										if pcv_CountGroup > 5 then
											%>
											<div class="pcCSFItem">
												<a href="javascript:ShowLess(<%=pcv_strCatSearchID%>)"><strong><%=dictLanguage.Item(Session("language")&"_categorysearchfield_3")%></strong></a>
											</div>
											<%
										end if
									end if
                                %>                            
						<%if pcv_CountGroup > 0 then%>
						</div>
                        <%end if%>
                        <%
						'// Create JavaScript Strings to Initialize the Panels
						pcv_strJQVars = pcv_strJQVars + 1
                        end if '// if LCount>0 then
    
                    end if '// if not rs.eof then
    
                Next '// For i=0 to intCount
                
            end if
            if pcv_ClearAll = 0 AND pcv_strSFNAME="" then
                %>
                <script type=text/javascript>
                   document.getElementById("pcCSF").style.display = 'none'; 
                </script>
                <%
            end if
            %>
        </div>
        </div>
        </form>
        <!-- search custom fields if any are defined -->
        <%
        tmpJSStr=""
        if pcv_intSFID<>"" then
            tmpJSStr=tmpJSStr & "var SFID=new Array(" & pcf_converToArray(pcv_intSFID) & ")" & vbcrlf
        else
            tmpJSStr=tmpJSStr & "var SFID=new Array();" & vbcrlf
        end if
        if pcv_strSFNAME<>"" then
            tmpJSStr=tmpJSStr & "var SFNAME=new Array(" & pcf_converToArray(pcf_SanitizeJava(pcv_strSFNAME)) & ")" & vbcrlf
        else
            tmpJSStr=tmpJSStr & "var SFNAME=new Array();" & vbcrlf	
        end if
        if pcv_strCValues<>"" then
            tmpJSStr=tmpJSStr & "var SFVID=new Array(" & pcf_converToArray(pcv_strCValues) & ")" & vbcrlf
        else
            tmpJSStr=tmpJSStr & "var SFVID=new Array();" & vbcrlf
        end if
        if pcv_strSFVALUE<>"" then
            tmpJSStr=tmpJSStr & "var SFVALUE=new Array(" & pcf_converToArray(pcf_SanitizeJava(pcv_strSFVALUE)) & ")" & vbcrlf
        else
            tmpJSStr=tmpJSStr & "var SFVALUE=new Array();" & vbcrlf
        end if
        if SFVORDER<>"" then
            tmpJSStr=tmpJSStr & "var SFVORDER=new Array(" & pcf_converToArray(SFVORDER) & ")" & vbcrlf
        else
            tmpJSStr=tmpJSStr & "var SFVORDER=new Array();" & vbcrlf
        end if
        if pcv_intSFCount<>"" then
            intCount=pcv_intSFCount
        else
            intCount=-1
        end if	        
        tmpJSStr=tmpJSStr & "var SFCount=" & intCount & ";" & vbcrlf
        %>
        <script type=text/javascript>
            <%=tmpJSStr%>
            function CreateTable(tmpRun)
            {
                var tmp1="";
                var tmp2="";
                var tmp3="";
                var i=0;
                var found=0;			
                tmp1='<div class="pcTable">';
                for (var i=0;i<=SFCount;i++)
                {
                    found=1;
                    tmp1=tmp1 + '<div class="pcTableRow"><div style="text-align: right; margin-bottom: 10px;"><a href="javascript:ClearSF(SFID['+i+']);"><img src="<%=pcf_getImagePath("../pc/images","minus.jpg")%>" alt="" border="0">&nbsp;</a></td><td style="text-align: left; width: 100%;">'+SFNAME[i]+': '+SFVALUE[i]+'</div></div>';
                    if (tmp2=="") tmp2=tmp2 + "||";
                    tmp2=tmp2 + SFID[i] + "||";
                    if (tmp3=="") tmp3=tmp3 + "||";
                    tmp3=tmp3 + SFVID[i] + "||";
                }
                tmp1=tmp1+'</div>';
                if (found==0) tmp1="";
                document.getElementById("stable").innerHTML=tmp1;
                if (tmp2=="") tmp2=0;
                document.getElementById("customfield").value=tmp2;
                document.getElementById("SearchValues").value=tmp3;
                if (tmp2==0)
                {
                    document.getElementById("customfield").value=0;
                    document.getElementById("SearchValues").value='';
                }
                if (tmpRun!=1) document.CSF.submit();
                
            }
    
            function ClearSF(tmpSFID)
            {
                var i=0;
                for (var i=0;i<=SFCount;i++)
                {
                    if (SFID[i]==tmpSFID)
                    {
                        removedArr = SFID.splice(i,1);
                        removedArr = SFNAME.splice(i,1);
                        removedArr = SFVID.splice(i,1);
                        removedArr = SFVALUE.splice(i,1);
                        removedArr = SFVORDER.splice(i,1);
                        SFCount--;
                        break;
                    }
                }
                document.getElementById("SFID").value=SFID.join("||");
                document.getElementById("SFNAME").value=SFNAME.join("||");
                document.getElementById("SFVID").value=SFVID.join("||");
                document.getElementById("SFVALUE").value=SFVALUE.join("||");
                document.getElementById("SFCount").value=SFCount;
                document.getElementById("notice").style.display = ''; 
                CreateTable(0);
            }
    
            function AddSF(tmpSFID,tmpSFName,tmpSVID,tmpSValue,tmpSOrder)
            {
                if ((tmpSVID!="") && (tmpSFID!="") && (tmpSVID!="0") && (tmpSFID!="0"))
                {
                    var i=0;
                    var found=0;
                    for (var i=0;i<=SFCount;i++)
                    {
                        if (SFID[i]==tmpSFID)
                        {
                            SFVID[i]=tmpSVID;
                            SFVALUE[i]=tmpSValue;
                            SFVORDER[i]=tmpSOrder;
                            found=1;
                            break;
                        }
                    }
                    if (found==0)
                    {
                        SFCount++;
                        SFID[SFCount]=tmpSFID;
                        SFNAME[SFCount]=tmpSFName;
                        SFVID[SFCount]=tmpSVID;
                        SFVALUE[SFCount]=tmpSValue;
                        SFVORDER[SFCount]=tmpSOrder;		
    
                    }
                    document.getElementById("SFID").value=SFID.join("||");
                    document.getElementById("SFNAME").value=SFNAME.join("||");
                    document.getElementById("SFVID").value=SFVID.join("||");
                    document.getElementById("SFVALUE").value=SFVALUE.join("||");
                    document.getElementById("SFCount").value=SFCount;
                    document.getElementById("notice").style.display = ''; 
                    CreateTable(0);
                }
            }  
    
            function ShowMore(VSID)
            {
                document.getElementById("VS"+VSID).value=1;
				document.getElementById("SFID").value=SFID.join("||");
				document.getElementById("SFNAME").value=SFNAME.join("||");
				document.getElementById("SFVID").value=SFVID.join("||");
				document.getElementById("SFVALUE").value=SFVALUE.join("||");
				document.getElementById("SFCount").value=SFCount;
				document.getElementById("notice").style.display = ''; 
				CreateTable(0);
            } 
    
            function ShowLess(VSID)
            {
                document.getElementById("VS"+VSID).value=0;
				document.getElementById("SFID").value=SFID.join("||");
				document.getElementById("SFNAME").value=SFNAME.join("||");
				document.getElementById("SFVID").value=SFVID.join("||");
				document.getElementById("SFVALUE").value=SFVALUE.join("||");
				document.getElementById("SFCount").value=SFCount;
				document.getElementById("notice").style.display = ''; 
				CreateTable(0);
            } 
    
            CreateTable(1);   
    
        </script>
        <%
        End If '// If pcv_strSearchDataIDs<>"" Then
        
    End If
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// END: Disply Widget 
	'//////////////////////////////////////////////////////////////////////////////////////////
	
END IF
'//////////////////////////////////////////////////////////////////////////////////////////////
'// END: CATEGORY SEARCH FIELDS
'//////////////////////////////////////////////////////////////////////////////////////////////

Session("pcv_strCSFilters") = pcv_strCSFilters

call closeDb()
%>
