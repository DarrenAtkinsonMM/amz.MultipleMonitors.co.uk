<% 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>

<%

'*******************************
' Settings
'*******************************

'// Set to "0" below to hide "View All Results" link.
pcv_ShowViewAllLink=1

iRecSize=10

'*******************************
' Generate Base Navigation Url
'*******************************
baseNavUrl = pcStrPageName
baseNavUrl = baseNavUrl & "?incSale=" & incSale 
baseNavUrl = baseNavUrl & "&IDSale=" & tmpIDSale 
baseNavUrl = baseNavUrl & "&ProdSort=" & ProdSort
baseNavUrl = baseNavUrl & "&PageStyle=" & pcPageStyle
baseNavUrl = baseNavUrl & "&customfield=" & pcustomfield
baseNavUrl = baseNavUrl & "&SearchValues=" & pCValues
baseNavUrl = baseNavUrl & "&exact=" & intExact
baseNavUrl = baseNavUrl & "&keyword=" & tKeywords
baseNavUrl = baseNavUrl & "&priceFrom=" & pPriceFrom
baseNavUrl = baseNavUrl & "&priceUntil=" & pPriceUntil
baseNavUrl = baseNavUrl & "&idCategory=" & pIdCategory
baseNavUrl = baseNavUrl & "&IdSupplier=" & IdSupplier
baseNavUrl = baseNavUrl & "&withStock=" & pWithStock
baseNavUrl = baseNavUrl & "&IDBrand=" & IDBrand
baseNavUrl = baseNavUrl & "&SKU=" & pSearchSKU
baseNavUrl = baseNavUrl & "&order=" & strORD
baseNavUrl = baseNavUrl & pcv_strCSFieldQuery

%>
        
<% If iPageCount>1 Then %>

	<div id="pcPagination<%= pcPageNavTopBottom %>" class="pcPagination">

	<% 'Page Number (x of y) %>
  <div class="pcPageResults">
  	<%= dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & iPageCount %>
      &nbsp;-&nbsp;
  </div>
  
	<% 'Navigation Links %>
		<% 
			If iPageCount>iRecSize Then
				If cint(iPageCurrent)>iRecSize Then
       		addUrl = ""
          addUrl = addUrl & "&VA=0"
          addUrl = addUrl & "&iPageCurrent=1"
          addUrl = addUrl & "&iPageSize=" & iPageSize
        	%>
            <a href="<%= Server.HtmlEncode(baseNavUrl & addUrl) %>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_1")%></a>
          
          	&nbsp;
        	<%
				End If
				
				If cint(iPageCurrent)>1 Then
        	If cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize Then
          	iPagePrev=cint(iPageCurrent)-1
          Else
           	iPagePrev=iRecSize
          End If
            
          addUrl = ""
          addUrl = addUrl & "&VA=0"
          addUrl = addUrl & "&iPageCurrent=" & (cint(iPageCurrent)-iPagePrev)
          addUrl = addUrl & "&iPageSize=" & iPageSize
          %>
            <a href="<%= Server.HtmlEncode(baseNavUrl & addUrl) %>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_2") & "&nbsp;" & iPagePrev & "&nbsp;" & dictLanguage.Item(Session("language")&"_PageNavigaion_3") %></a>
        	<%
				End If
				
        If cint(iPageCurrent)+1>1 Then
        	intPageNumber=cint(iPageCurrent)
        Else
        	intPageNumber=1
        End If
    	Else
        intPageNumber=1
    	End If

    	If (cint(iPageCount)-cint(iPageCurrent))<iRecSize Then
        iPageNext=cint(iPageCount)-cint(iPageCurrent)
    	Else
        iPageNext=iRecSize
    	End If

    	For pageNumber=intPageNumber To (cint(iPageCurrent) + (iPageNext))
        If Cint(pageNumber)=Cint(iPageCurrent) Then
					%>
						<b><%=pageNumber%></b> 
					<%
				Else
        	addUrl = ""
        	addUrl = addUrl & "&VA=0"
          addUrl = addUrl & "&iPageCurrent=" & pageNumber
          addUrl = addUrl & "&iPageSize=" & iPageSize
          %>
            <a href="<%= Server.HtmlEncode(baseNavUrl & addUrl) %>"><%=pageNumber%></a>
        	<%
				End If 
    	Next

    	If Not (cint(iPageNext)+cint(iPageCurrent))=iPageCount Then
        If iPageCount>(cint(iPageCurrent) + (iRecSize-1)) Then
          addUrl = ""
          addUrl = addUrl & "&VA=0"
          addUrl = addUrl & "&iPageCurrent=" & cint(intPageNumber)+iPageNext
          addUrl = addUrl & "&iPageSize=" & iPageSize
          %>
          	<a href="<%= Server.HtmlEncode(baseNavUrl & addUrl) %>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_4") & "&nbsp;" & iPageNext & "&nbsp;" & dictLanguage.Item(Session("language")&"_PageNavigaion_3") %></a>
        	<%
				End If

        If cint(iPageCount)>iRecSize AND (cint(iPageCurrent)<>cint(iPageCount)) Then
					addUrl = ""
					addUrl = addUrl & "&VA=0"
					addUrl = addUrl & "&iPageCurrent=" & cint(iPageCount)
					addUrl = addUrl & "&iPageSize=" & iPageSize
          %>
            <a href="<%= Server.HtmlEncode(baseNavUrl & addUrl) %>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_5")%></a>
        	<%
				End If 
      End If 

      If pcv_ShowViewAllLink=1 Then
				addUrl = "&VA=1"
				%> 
          <a href="<%= Server.HtmlEncode(baseNavUrl & addUrl) %>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_6")%></a>
      	<%
			End If 
		%>
	</div>
<% End If %>
