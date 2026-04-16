<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories
%>

<%

Dim iShow, pSQty, pCnt, pAddtoCart, pAllCnt

Sub pcShowProductsMTop(pArray, pCount, pPageSize)
	
	readCnt = pCount
	
	if pPageSize<>"" And Not IsNull(pPageSize) And pPageSize < pCount then
		readCnt = pPageSize
	end if
	
	'Check If customers are allowed to order products
	iShow=0
	pAddtoCart = 0
	
	If scOrderlevel=0 Then
		iShow=1
	End If
	If scOrderlevel=1 AND session("customerType")="1" Then
		iShow=1
	End If
	
	'reset count variables
	pCnt=Cint(0)
	pAllCnt=Cint(0)

	'// Run through the products to count all products, products with options, and BTO products
	do while (pCnt < readCnt)
		
		pidrelation=pArray(0,pCnt) '// rsCount("idProduct")
		pserviceSpec=pArray(6,pCnt) '// rsCount("serviceSpec")	
		pStock=pArray(10,pCnt) '// rsCount("stock")
		pNoStock=pArray(11,pCnt) '// rsCount("noStock")
		pcv_intBackOrder=pArray(13,pCnt) '// rs("pcProd_BackOrder")
		
		pCnt=pCnt+1
		
		' Check which items will have multi qty enabled,
		pcv_SkipCheckMinQty=-1 
		If pcf_AddToCart(pidrelation)=False Then
			pAllCnt=pAllCnt+1
		End If
		
	loop
	
	pcv_SkipCheckMinQty=0
	
	' If all items on the page are either BTO or have options,
	' do not show the quantity column or the Add to Cart button.						
	If cint(pAllCnt) <> cint(pCnt) Then 
		pAddtoCart = 1
	End If
	
	pCnt=Cint(0)
	pSQty=0
	pAllCnt=Cint(0)
%>
        
<!--Product List Start-->
<% if iShow=1 AND pAddtoCart = 1 then %> 
  <div class="pcShowProductsListDesc">
    <p><%= dictLanguage.Item(Session("language")&"_viewCat_P_12") %></p>
  </div>
<% end if %>
  
<div class="pcShowProductsMTable">
  <div class="pcShowProductsMHeader">
    <% If iShow=1 Then %> 
      <% If pAddtoCart = 1 Then %>
        <div class="pcShowProductQtyM">
          <%= dictLanguage.Item(Session("language")&"_viewCat_P_7") %>
        </div>
      <% End If %>
    <% End If %>
    <div class="pcShowProductImageM">&nbsp;</div>
    <% If pShowSku <> 0 Then %>
      <div class="pcShowProductSkuM">
        <%= dictLanguage.Item(Session("language")&"_viewCat_P_8") %>
      </div>
    <% End If %>
    <div class="pcShowProductNameM">
      <%= dictLanguage.Item(Session("language")&"_viewCat_P_9") %>
    </div>
    <div class="pcShowProductPriceM">
      <% If session("customerType")="1" Then %>
        <%= dictLanguage.Item(Session("language")&"_viewCat_P_11") %>
      <% Else %>
        <%= dictLanguage.Item(Session("language")&"_viewCat_P_10") %>
      <% End If %>
    </div>
  </div>
<%
End Sub

Sub pcShowProductsMBottom()
	'Show the Add to Cart button when
	' products can be added to the cart from this page.	
%>
  <input type="hidden" name="pCnt" value="<%=pCnt%>">
</div>
<!--Product List End-->
<%
End Sub

Sub pcShowProductsMBefore()
%>
	<form action="instPrd.asp" method="post" name="m" id="m" class="pcForms">
<%
End Sub

Sub pcShowProductsMAfter()
	If iShow=1 and clng(pSQty)<>0 Then %>
    <div class="pcShowProductAddToCart">
			<button class="pcButton pcButtonAddToCart" name="submit" id="submit">
				<img alt="<%= dictLanguage.Item(Session("language")&"_prdD1_3") %>" src="<%=pcf_getImagePath("",rslayout("addtocart"))%>" />
				<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_addtocart") %></span>
			</button>
    </div>
  <% 
	End If 
	%>
  </form>
<%
End Sub

Sub pcShowProductM(dblpcCC_Price)

    pIdCategoryTemp = getFirstCategoryID(pIdProduct, pIdCategory)
    
    atc_FlagM = "1"
	
	'// Call SEO Routine
	pcGenerateSeoLinks
	'//
	
	'// If product is "Not for Sale", should prices be hidden or shown?
	'// Set pcHidePricesIfNFS = 1 to hide, 0 to show.
	'// Here we leverage the "pcv_intHideBTOPrice" variable to change the behavior (a Control Panel setting could be added in the future)
	pcHidePricesIfNFS = 0
	if (pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0) and pcHidePricesIfNFS=1 then
		pcv_intHideBTOPrice=1
	end if
	
	pCnt=pCnt+1
%>

<div class="pcShowProductsM" itemscope itemtype="http://schema.org/Product">
	<% if iShow=1 AND pAddtoCart = 1 then %> 
  	<div class="pcShowProductQtyM">
			<% 
      '// Allow Multiple Qtys (the "pcf_AddToCart" function will not validate min qtys)
      pcv_SkipCheckMinQty=-1 
      %>
      <% If pcf_AddToCart(pIdProduct)=True Then %> 
          
        <%
        '//////////////////////////////////////////////////////////////////////
        '// Start: Validate for multiple of N
        '//////////////////////////////////////////////////////////////////////
        query="select pcprod_QtyValidate,pcprod_MinimumQty,pcProd_multiQty from products where idproduct=" & pidProduct 									
        set rs1=Server.CreateObject("ADODB.Recordset")
        set rs1=connTemp.execute(query)
        if err.number<>0 then
          call LogErrorToDatabase()
          set rs1=nothing
          call closedb()
          response.redirect "techErr.asp?err="&pcStrCustRefID
        end if
        pcv_intQtyValidate=rs1("pcprod_QtyValidate")
        if not pcv_intQtyValidate<>"" then
          pcv_intQtyValidate="0"
        end if			
        pcv_lngMinimumQty=rs1("pcprod_MinimumQty")
        if not pcv_lngMinimumQty<>"" then
          pcv_lngMinimumQty="0"
        end if
            pcv_lngMultiQty=rs1("pcProd_multiQty")
            if IsNull(pcv_lngMultiQty) or pcv_lngMultiQty="" then
              pcv_lngMultiQty="0"
            end if
        set rs1 = nothing
        pcv_lngQty = 1
        if pcv_intQtyValidate<>"1" then 
          pcv_lngQty=0
        end if
        '//////////////////////////////////////////////////////////////////////
        '// End: Validate for multiple of N
        '//////////////////////////////////////////////////////////////////////
        %>					
        <input name="idProduct<%=pCnt%>" type="hidden" value="<%=pidProduct%>">					
        <%
        pSQty=pSQty+1 '// "add to cart" button flag
        pcv_SkipCheckMinQty=0
        pcv_strOnBlur = "checkproqty(this,"&pcv_lngMinimumQty&","&pcv_lngQty&","&pcv_lngMultiQty&")"
        %>
        <input name="QtyM<%=pidProduct%>" type="text" value="0" size="2" maxlength="10" onBlur="<%=pcv_strOnBlur%>">
      <% Else %>
        <input name="idProduct<%=pCnt%>" type="hidden" value="<%=pidProduct%>">
		<meta itemprop="productID" content="<%=pidProduct%>" />
        <input type="hidden" name="QtyM<%=pidProduct%>" value="0">   
      <% End If %>
    </div>
  <% end if %>

	<div class="pcShowProductImageM">
		<%if pShowSmallImg <> 0 then%>
    	<%if pSmallImageUrl<>"" then%>
	      <a itemprop="url" href="<%=Server.HtmlEncode(pcStrPrdLink)%>" data-idProduct="<%= pIdProduct %>"><img itemprop="image" src="<%=pcf_getImagePath("catalog",pSmallImageUrl)%>" alt="<%=pDescription%>"></a>
      <% else %>
          &nbsp;
      <%end if %>
    <% end if %>
  </div>
        
	<%if pShowSKU <> 0 then%>
  	<div class="pcShowProductSkuM">
    	<span itemprop="sku" class="pcShowProductSku"><%=pSku%></span>
	  </div>
  <% end if %>
  
  <div class="pcShowProductNameM">
  	<div class="pcShowProductName">
  		<a itemprop="url" href="<%=Server.HtmlEncode(pcStrPrdLink)%>"><span itemprop="name"><%=pDescription%></span></a>
    </div>
    
		<% If pcf_HasHTMLContent(psDesc) Then %>
	    <div class="pcShowProductSDesc">
				<span itemprop="description"><%= pcf_FixHTMLContentPaths(psDesc) %></span>
      	<!--#include file="pcShowProductReview.asp" -->
      </div>
    <% End If %>
	</div>

	<div class="pcShowProductPriceM">
		<%ShowSaleIcon=0
  
    if UCase(scDB)="SQL" then	
    if pnoprices=0 then
    query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveIcon,pcSales_BackUp.pcSales_TargetPrice FROM (pcSales_Completed INNER JOIN Products ON pcSales_Completed.pcSC_ID=Products.pcSC_ID) INNER JOIN pcSales_BackUp ON pcSales_BackUp.pcSC_ID=pcSales_Completed.pcSC_ID WHERE Products.idproduct=" & pidproduct & " AND Products.pcSC_ID>0;"
    set rsS=Server.CreateObject("ADODB.Recordset")
    set rsS=conntemp.execute(query)
  
    if not rsS.eof then
      ShowSaleIcon=1
      pcSCID=rsS("pcSC_ID")
      pcSCName=rsS("pcSC_SaveName")
      pcSCIcon=rsS("pcSC_SaveIcon")
      pcTargetPrice=rsS("pcSales_TargetPrice")
    end if
    set rsS=nothing
    end if
    end if
    %>
    
    <% If (pPrice>"0") And (pcv_intHideBTOPrice<>"1") then %>
        <div class="pcShowProductPrice" itemprop="offers" itemscope itemtype="http://schema.org/Offer">
            <% If scCursign = "$" Then %><meta itemprop="priceCurrency" content="USD" /><% End If %>
            <meta itemprop="price" content="<%=pPrice%>" />
            <%= scCursign & money(pPrice)%>
            <% If (ShowSaleIcon=1) And (session("customerCategory")=0) And (pcTargetPrice="0") Then %>
                <span class="pcSaleIcon">
                    <a href="javascript:openbrowser('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="<%=pcf_getImagePath("catalog",pcSCIcon)%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a>
                </span>
            <% End If %>
            <!--#include file="pcShowQtyDiscIcon.asp" -->
        </div>
        <% If (pListPrice-pPrice)>0 And plistHidden<0 Then %>
            <div class="pcShowProductSavings">
                <%= dictLanguage.Item(Session("language")&"_prdD1_2") & scCursign & money(pListPrice-pPrice) & " (" & round(((pListPrice-pPrice)/pListPrice)*100) & "%)"%>
            </div>
        <% End If %>
        <% If session("customerCategory")<>0 and (session("customerType")="1") and (pcv_intHideBTOPrice<>"1") Then %>
        <% Else %>
        <input name="BTOTOTAL<%=pCnt%>" type="hidden" value="<%=pPrice%>">
        <% End If %>        
    <% End If %>  
      
    <%
    'if customer category type logged in - show pricing
    if session("customerCategory")<>0 and (dblpcCC_Price>"0") and (pcv_intHideBTOPrice<>"1") then %>
        <div class="pcShowProductPriceW">
      <%= session("customerCategoryDesc")& ": " & scCursign & money(dblpcCC_Price)%>
      <%if (ShowSaleIcon=1) AND (clng(session("customerCategory"))=clng(pcTargetPrice)) then%>
      <span class="pcSaleIcon"><a href="javascript:openbrowser('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="<%=pcf_getImagePath("catalog",pcSCIcon)%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
      <%end if%>
      <input name="BTOTOTAL<%=pCnt%>" type="hidden" value="<%=dblpcCC_Price%>">
        </div>
    <%else
      if (dblpcCC_Price>"0") and (session("customerType")="1") and (pcv_intHideBTOPrice<>"1") then %>
        <div class="pcShowProductPriceW">
        <%= dictLanguage.Item(Session("language")&"_prdD1_4")& " " & scCursign & money(dblpcCC_Price)%>
        <%if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=-1) then%>
        <span class="pcSaleIcon"><a href="javascript:openbrowser('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="<%=pcf_getImagePath("catalog",pcSCIcon)%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
        <%end if%>
        </div>
        <input name="BTOTOTAL<%=pCnt%>" type="hidden" value="<%=dblpcCC_Price%>">
      <%end if
    end if
 	 	%>
  </div>
</div>
<% 
End Sub 
%>
