<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/Or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories
pIdCategoryTemp = getFirstCategoryID(pIdProduct, pIdCategory)

'// Call SEO Routine
pcGenerateSeoLinks
'//

'// If product is "Not fOr Sale", should prices be hidden Or shown?
'// Set pcHidePricesIfNFS = 1 to hide, 0 to show.
'// Here we leverage the "pcv_intHideBTOPrice" variable to change the behaviOr (a Control Panel setting could be added in the future)
pcHidePricesIfNFS = 0
If (pFOrmQuantity="-1" and NotFOrSaleOverride(session("customerCategory"))=0) and pcHidePricesIfNFS=1 Then
	pcv_intHideBTOPrice=1
End If
%>
<div class="pcShowProductsP <%= pcProductHover %>" itemscope itemtype="http://schema.org/Product">
  <div class="pcShowProductImageP">
 		<% if pSmallImageUrl<>"" then %>
    	    <a itemprop="url" href="<%=Server.HtmlEncode(pcStrPrdLink)%>"><img itemprop="image" src="<%=pcf_getImagePath("catalog",pSmallImageUrl)%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_1")& pDescription %>"></a>
        <% Else %>
    	    <a itemprop="url" href="<%=Server.HtmlEncode(pcStrPrdLink)%>" class="pcShowProductNoImage" ><img itemprop="image" src="<%=pcf_getImagePath("catalog","no_image.gif")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_1")& pDescription %>"></a>
        <%End If%>    
		<%'QV-S%>
        <%
        tmpCanAddToCart=pcf_AddToCartQV(pIdProduct)
        If (scDisplayQuickView=1) And (tmpCanAddToCart=True) And (pcv_intSkipDetailsPage=0) Then
        %>
          <% pcf_QuickViewBtn pIdProduct %>
        <%
        End If
        %>
        <%'QV-E%>
  </div>  
  <div class="pcShowProductInfoP">
    <div class="pcShowProductName">
      <a itemprop="url" href="<%=Server.HtmlEncode(pcStrPrdLink)%>"><span itemprop="name"><%=pDescription%></span></a>
    </div>
      <%ShowSaleIcon=0

      If UCase(scDB)="SQL" Then	
          If pnoprices=0 Then
              query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveIcon,pcSales_BackUp.pcSales_TargetPrice FROM (pcSales_Completed INNER JOIN Products ON pcSales_Completed.pcSC_ID=Products.pcSC_ID) INNER JOIN pcSales_BackUp ON pcSales_BackUp.pcSC_ID=pcSales_Completed.pcSC_ID WHERE Products.idproduct=" & pidproduct & " AND Products.pcSC_ID>0;"
              set rsS=Server.CreateObject("ADODB.RecOrdset")
              set rsS=conntemp.execute(query)

              If not rsS.eof Then
                  ShowSaleIcon=1
                  pcSCID=rsS("pcSC_ID")
                  pcSCName=rsS("pcSC_SaveName")
                  pcSCIcon=rsS("pcSC_SaveIcon")
                  pcTargetPrice=rsS("pcSales_TargetPrice")
							End If
              set rsS=nothing
          End If
      End If
      %>
      
      <% If (pPrice>0) And (pcv_intHideBTOPrice<>"1") Then %>
        <div class="pcShowProductPrice" itemprop="offers" itemscope itemtype="http://schema.org/Offer">
            <% If scCursign = "$" Then %><meta itemprop="priceCurrency" content="USD" /><% End If %>
            <meta itemprop="price" content="<%=pPrice%>" />
            <%=dictLanguage.Item(Session("language")&"_prdD1_1") & ": " %> <%=scCursign & money(pPrice) %>
        </div>
        <% If (ShowSaleIcon=1) And (session("customerCategory")=0) And (pcTargetPrice="0") Then %>
            <span class="pcSaleIcon">
                <a href="javascript:openbrowser('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="<%=pcf_getImagePath("catalog",pcSCIcon)%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a>
            </span>
        <% End If %>
        <!--#include file="pcShowQtyDiscIcon.asp" -->
        <% If (pListPrice-pPrice)>0 And plistHidden<0 And session("customerType")<>1 Then %>
              <div class="pcShowProductListPrice">
                <%=dictLanguage.Item(Session("language")&"_viewPrd_20")%><%=scCursign & money(pListPrice)%>
              </div>
              <div class="pcShowProductSavings">
                <%=dictLanguage.Item(Session("language")&"_prdD1_2") & scCursign & money(pListPrice-pPrice) & " (" & round(((pListPrice-pPrice)/pListPrice)*100) & "%)"%>
              </div>
          <% End If %>
      <% End If %>
      
      <% 'If customer category type logged in - show pricing
      If session("customerCategory")<>0 and (dblpcCC_Price>"0") and (pcv_intHideBTOPrice<>"1") Then %>
          <span class="pcShowProductPriceW">
              <%= session("customerCategoryDesc")& " " & scCursign & money(dblpcCC_Price)%>
              <%If (ShowSaleIcon=1) AND (clng(session("customerCategory"))=clng(pcTargetPrice)) Then%>
                  <span class="pcSaleIcon"><a href="javascript:openbrowser('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="<%=pcf_getImagePath("catalog",pcSCIcon)%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
              <%End If%>
          </span>
      <% Else %>
          <% If (dblpcCC_Price>"0") and (session("customerType")="1") and (pcv_intHideBTOPrice<>"1") Then %>
              <span class="pcShowProductPriceW">
                  <%= dictLanguage.Item(Session("language")&"_prdD1_4")& " " & scCursign & money(dblpcCC_Price)%>
                  <%If (ShowSaleIcon=1) AND (clng(pcTargetPrice)=-1) Then%>
                  <span class="pcSaleIcon"><a href="javascript:openbrowser('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="<%=pcf_getImagePath("catalog",pcSCIcon)%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
                  <%End If%>
              </span>
          <% End If 
      End If %>

    <!--#include file="pcShowProductReview.asp" -->
    <%
    ' Show shOrt product description
    If pcf_HasHTMLContent(psDesc) Then
    %>
      <span itemprop="description" class="pcShowProductSDesc"><%= pcf_FixHTMLContentPaths(psDesc) %></span>
    <% 
    End If
  
    'SB S
    Set objSB = New pcARBClass
    pSubscriptionID = objSB.getSubscriptionID(pIdProduct)
    If isNull(pSubscriptionID) Or pSubscriptionID="" Then
        pSubscriptionID = "0"
    End If
    %>
    <!--#include file="../includes/pcSBDataInc.asp" --> 
    
    <% 
    'SB E
  
    ' Show product details page button %>
    <div class="pcShowProductButtonsP">   
      <a class="pcButton pcButtonMoreDetails" itemprop="url" href="<%=Server.HtmlEncode(pcStrPrdLink)%>">
        <img src="<%=pcf_getImagePath("",rslayout("morebtn"))%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_1")& pDescription %>">
        <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_morebtn") %></span>
      
      </a>
      <%
        If pcf_AddToCart(pIdProduct)=True Then
          addPrdLink = "instPrd.asp?idproduct=" & pIdProduct & "&pSubscriptionID=" & pSubscriptionID
          %>
            <a class="pcButton pcButtonAddToCartSmall" href="<%=Server.HtmlEncode(addPrdLink)%>" rel="nofollow">
              <img class="pcButtonImage" src="<%=pcf_getImagePath("",rslayout("add2"))%>" alt="<%=dictLanguage.Item(Session("language")&"_css_add2")& pDescription %>">
              <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_add2") %></span>
            </a>
          <%
        End If
      %>
      <!--#include file="inc_addPinterest.asp"-->
    </div>
  </div>
  <div class="pcClear"></div>
</div>
