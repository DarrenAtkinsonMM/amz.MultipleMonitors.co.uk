<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
  
'// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories
    pIdCategoryTemp = getFirstCategoryID(pIdProduct, pIdCategory)
    
	'// Call SEO Routine
	pcGenerateSeoLinks
	'//
	
	'// If product is "Not for Sale", should prices be hidden or shown?
	'// Set pcHidePricesIfNFS = 1 to hide, 0 to show.
	'// Here we leverage the "pcv_intHideBTOPrice" variable to change the behavior (a Control Panel setting could be added in the future)
	pcHidePricesIfNFS = 0
	If (pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0) and pcHidePricesIfNFS=1 then
		pcv_intHideBTOPrice=1
	end If
%>

<div class="pcShowProductsH <%= pcProductHover %>" itemscope itemtype="http://schema.org/Product">
	<div class="pcShowProductImageH">
		<% if pSmallImageUrl<>"" then %>
            <a itemprop="url" href="<%=Server.HtmlEncode(pcStrPrdLink)%>"><img itemprop="image" src="<%=pcf_getImagePath("catalog",pSmallImageUrl)%>" alt="<%=pDescription %>"></a>
        <% else %>
            <a itemprop="url" href="<%=Server.HtmlEncode(pcStrPrdLink)%>"><img itemprop="image" src="<%=pcf_getImagePath("catalog","no_image.gif")%>" alt="<%=pDescription %>"></a>
        <% end if %>    
		<%'QV-S%>
        <%      
        If scDisplayQuickView=1 And pcv_intSkipDetailsPage=0 Then
        %>
            <% pcf_QuickViewBtn pIdProduct %>
        <%
        End If
        %>
        <%'QV-E%>  
  </div>
  <div class="pcClear"></div>
  
  <div class="pcShowProductInfoH">
    <div class="pcShowProductName">
      <a href="<%=Server.HtmlEncode(pcStrPrdLink)%>"><span itemprop="name"><%=pDescription%></span></a>
    </div>
  <%ShowSaleIcon=0

    If pnoprices=0 then      

      query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveIcon,pcSales_BackUp.pcSales_TargetPrice FROM (pcSales_Completed INNER JOIN Products ON pcSales_Completed.pcSC_ID=Products.pcSC_ID) INNER JOIN pcSales_BackUp ON pcSales_BackUp.pcSC_ID=pcSales_Completed.pcSC_ID WHERE Products.idproduct=" & pidproduct & " AND Products.pcSC_ID>0;"

      set rsS=Server.CreateObject("ADODB.Recordset")
      set rsS=conntemp.execute(query)
      
      If not rsS.eof then
        ShowSaleIcon=1
        pcSCID=rsS("pcSC_ID")
        pcSCName=rsS("pcSC_SaveName")
        pcSCIcon=rsS("pcSC_SaveIcon")
        pcTargetPrice=rsS("pcSales_TargetPrice")
      end If
      set rsS=nothing
    end If

  %>
  
  <% If (pPrice>0) and (pcv_intHideBTOPrice<>"1") then %>
        
        <div class="pcShowProductPrice" itemprop="offers" itemscope itemtype="http://schema.org/Offer">
            <% If scCursign = "$" Then %><meta itemprop="priceCurrency" content="USD" /><% End If %>
            <meta itemprop="price" content="<%=pPrice%>" />
            <%=dictLanguage.Item(Session("language")&"_prdD1_1") & ": " %> <%=scCursign & money(pPrice) %>
        </div>
        
        <% If (ShowSaleIcon=1) And (session("customerCategory")=0) And (pcTargetPrice="0") Then %>
              <span class="pcSaleIcon"><a href="javascript:openbrowser('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="<%=pcf_getImagePath("catalog",pcSCIcon)%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
        <% End If %>
        
        <!--#include file="pcShowQtyDiscIcon.asp" -->
        
        <% If (pListPrice-pPrice)>0 And plistHidden<0 And session("customerType")<>1 Then %>
            <div class="pcShowProductListPrice">
                <%=dictLanguage.Item(Session("language")&"_viewPrd_20")%><%=scCursign & money(pListPrice)%>             </div>
            <div class="pcShowProductSavings">
                <%=dictLanguage.Item(Session("language")&"_prdD1_2") & scCursign & money(pListPrice-pPrice) & " (" & round(((pListPrice-pPrice)/pListPrice)*100) & "%)"%>
            </div>
        <% End If %>
        
  <% End If %>
      
  <% 'If customer category type logged in - show pricing
  If session("customerCategory")<>0 and (dblpcCC_Price>"0") and (pcv_intHideBTOPrice<>"1") then %>
    <span class="pcShowProductPriceW">
      <% response.write session("customerCategoryDesc")& " " & scCursign & money(dblpcCC_Price)%>
      <%If (ShowSaleIcon=1) AND (clng(session("customerCategory"))=clng(pcTargetPrice)) then%>
        <span class="pcSaleIcon"><a href="javascript:openbrowser('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="<%=pcf_getImagePath("catalog",pcSCIcon)%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
      <%end If%>
    </span>
  <% else %>
    <% If (dblpcCC_Price>"0") and (session("customerType")="1") and (pcv_intHideBTOPrice<>"1") then %>
      <span class="pcShowProductPriceW">
        <% response.write dictLanguage.Item(Session("language")&"_prdD1_4")& " " & scCursign & money(dblpcCC_Price)%>
        <%If (ShowSaleIcon=1) AND (clng(pcTargetPrice)=-1) then%>
        <span class="pcSaleIcon"><a href="javascript:openbrowser('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="<%=pcf_getImagePath("catalog",pcSCIcon)%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
        <%end If%>
      </span>
    <% end If 
  end If %>
  <!-- #include file="pcShowProductReview.asp" -->

  </div>

  <%
  'SB S
  Set objSB = New pcARBClass
  pSubscriptionID = objSB.getSubscriptionID(pIdProduct)
  If isNull(pSubscriptionID) OR pSubscriptionID="" then
      pSubscriptionID = "0"
  end If
  %>
  <!--#include file="../includes/pcSBDataInc.asp" --> 
  <% 
  'SB E	
  %>
    
  <div class="pcShowProductButtonsH">
    <a itemprop="url" href="<%=Server.HtmlEncode(pcStrPrdLink)%>" class="pcButton pcButtonMoreDetails">
      <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_morebtn") & " - " & pDescription %>">
      <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_morebtn") %></span>
    </a>

    <%
      If pcf_AddToCart(pIdProduct)=True Then
        addPrdLink = "instPrd.asp?idproduct=" & pIdProduct & "&pSubscriptionID=" & pSubscriptionID
        %>
          <a href="<%=Server.HtmlEncode(addPrdLink)%>" class="pcButton pcButtonAddToCartSmall" rel="nofollow">
            <img src="<%=pcf_getImagePath("",rslayout("add2"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_add2") & " " & pDescription %>">
            
            <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_add2") %></span>
          </a>
        <%
      End If
    %>
    <!--#include file="inc_addPinterest.asp"-->
  </div>
  
  <div class="pcClear"></div>
</div>
