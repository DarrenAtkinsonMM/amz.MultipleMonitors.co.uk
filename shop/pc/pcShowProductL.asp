<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>

<%
Sub pcShowProductsLTop()
%>
<!--Product List Start-->
<div class="pcShowProductsLTable">
  <div class="pcShowProductsLHeader">
    <% If pShowSmallImg <> 0 Then %>
      <div class="pcShowProductImageL">&nbsp;</div>
    <% End If %>
      
    <div class="pcShowProductNameL"><%= dictLanguage.Item(Session("language")&"_viewCat_P_9") %></div>
    
    <%If pShowSku <> 0 Then %>
      <div class="pcShowProductSkuL"><%= dictLanguage.Item(Session("language")&"_viewCat_P_8") %></div>
    <%End If %>
    
    <div class="pcShowProductPriceL"><%= dictLanguage.Item(Session("language")&"_viewCat_P_10") %></div>
    
  </div>
<%
End Sub

Sub pcShowProductsLBottom()
%>
	</div>
	<!--Product List End-->
<%
End Sub

Sub pcShowProductL(dblpcCC_Price)

    '// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories
    pIdCategoryTemp = getFirstCategoryID(pIdProduct, pIdCategory)

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
%>

<div class="pcShowProductsL" itemscope itemtype="http://schema.org/Product">
	
    <% if pShowSmallImg <> 0 then %>

        <div class="pcShowProductImageL">
            <% if pSmallImageUrl<>"" then %>
                <a itemprop="url" href="<%=Server.HtmlEncode(pcStrPrdLink)%>"><img itemprop="image" src="<%=pcf_getImagePath("catalog",pSmallImageUrl)%>" alt="<%=pDescription%>"></a>
            <% else %>
      	        &nbsp;
            <%end if %>
        </div>
        
    <% end if %>
        
	<div class="pcShowProductNameL">
    <div class="pcShowProductName">
      <a itemprop="url" href="<%=Server.HtmlEncode(pcStrPrdLink)%>"><span itemprop="name"><%=pDescription%></span></a>
    </div>
  
    <% If pcf_HasHTMLContent(psDesc) Then %>
      <div class="pcShowProductSDesc">
        <span itemprop="description"><%= pcf_FixHTMLContentPaths(psDesc) %></span>
        <!--#include file="pcShowProductReview.asp" -->
      </div>
    <% End If %>
    
		<div class="pcClear"></div>
  </div>
  
        <%if pShowSKU <> 0 then%>
        	<div class="pcShowProductSkuL">
          	<div class="pcShowProductSku" itemprop="sku">
							<%=pSku%>
            </div>
          </div>
				<%end if %>
  
        
        <div class="pcShowProductPriceL">
				<%
				
				ShowSaleIcon=0
        
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
        
		<% If (pPrice>0) And (pcv_intHideBTOPrice<>"1") Then %>
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
            <% If (pListPrice-pPrice)>0 And plistHidden<0 And session("customerType")<>1 Then %>
                <div class="pcShowProductSavings">
                    <%= dictLanguage.Item(Session("language")&"_prdD1_2") & scCursign & money(pListPrice-pPrice) & " (" & round(((pListPrice-pPrice)/pListPrice)*100) & "%)"%>
                </div>
            <% End If %>
        <% End If %>
        
        <% 'if customer category type logged in - show pricing
        if session("customerCategory")<>0 and (dblpcCC_Price>"0") and (pcv_intHideBTOPrice<>"1") then %>
            <div class="pcShowProductPriceW">
                <%= session("customerCategoryDesc")& " " & scCursign & money(dblpcCC_Price)%>
                <%if (ShowSaleIcon=1) AND (clng(session("customerCategory"))=clng(pcTargetPrice)) then%>
                <span class="pcSaleIcon"><a href="javascript:openbrowser('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="<%=pcf_getImagePath("catalog",pcSCIcon)%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
                <%end if%>
            </div>
        <% else %>
            <% if (dblpcCC_Price>"0") and (session("customerType")="1") and (pcv_intHideBTOPrice<>"1") then %>
                <div class="pcShowProductPriceW">
                    <%= dictLanguage.Item(Session("language")&"_prdD1_4")& " " & scCursign & money(dblpcCC_Price)%>
                    <%if (ShowSaleIcon=1) AND (clng(pcTargetPrice)=-1) then%>
                    <span class="pcSaleIcon"><a href="javascript:openbrowser('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="<%=pcf_getImagePath("catalog",pcSCIcon)%>" title="<%=pcSCName%>" alt="<%=pcSCName%>"></a></span>
                    <%end if%>
                </div>
            <% end if 
        end if %>
        
        </div>
		<%
        ' Detailed Product Reviews - START
        pcv_IDProduct = pIDProduct
        if pcStrPageNameOR="showRecentlyReviewed.asp" then %>
        <div>
          <!--#include file="prv_increviews.asp"-->
				</div>
        <% End If %>
  
</div>
<%
End Sub
%>
