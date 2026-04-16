<% 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>

<!--#include file="pcCheckPricingCats.asp"-->

<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Product of the Month
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  
Public Sub pcs_ProductOfTheMonth
  if pcIntHPFirst<>0 then %>
    <!-- Product of the Month -->
    <div class="pcPOMResult">
        <h2><%=dictLanguage.Item(Session("language")&"_mainIndex_12")%></h2>
        <div class="pcShowProducts">
    <%
    'Set the product count to zero
    count=0
    
    tCnt=Cint(0)
        
    do while (tCnt < pcv_intProductCount) and (count < 1)

        pidProduct=pcArray_Products(0,tCnt)
        pSku=pcArray_Products(1,tCnt)
        pDescription=pcArray_Products(2,tCnt)  
        pPrice=pcArray_Products(3,tCnt)
        pListHidden=pcArray_Products(4,tCnt)
        pListPrice=pcArray_Products(5,tCnt)              
        pserviceSpec=pcArray_Products(6,tCnt)
        pBtoBPrice=pcArray_Products(7,tCnt)   
        pSmallImageUrl=pcArray_Products(8,tCnt)   
        pnoprices=pcArray_Products(9,tCnt)
        if isNULL(pnoprices) OR pnoprices="" then
            pnoprices=0
        end if
        pStock=pcArray_Products(10,tCnt)
        pNoStock=pcArray_Products(11,tCnt)
        pcv_intHideBTOPrice=pcArray_Products(12,tCnt)
        if isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
            pcv_intHideBTOPrice="0"
        end if
        if pnoprices=2 then
            pcv_intHideBTOPrice=1
        end if
        pFormQuantity=pcArray_Products(14,tCnt)
        pcv_intBackOrder=pcArray_Products(15,tCnt)
        pidrelation=pcArray_Products(0,tCnt)            
        'SB S
        Dim objSB 
        Set objSB = New pcARBClass
        pSubscriptionID = objSB.getSubscriptionID(pidProduct)
        if isNull(pSubscriptionID) OR pSubscriptionID="" then
            pSubscriptionID = "0"
        end if          
        'SB E
      '// Get sDesc
      query="SELECT sDesc FROM products WHERE idProduct="&pidrelation&";"
      set rsDescObj=server.CreateObject("ADODB.RecordSet")
      set rsDescObj=conntemp.execute(query)
      psDesc=rsDescObj("sDesc")
      set rsDescObj=nothing
      
      if pcPageStyle = "m" then
        pCnt=pCnt+1
      end if
      tCnt=tCnt+1
      %>
      <!--#include file="pcGetPrdPrices.asp"-->
      <%
      '*******************************
      ' Show product information
      '*******************************
      %>            
      <!--#include file="pcShowProductP.asp" -->
      <% 
      count=count + 1
    loop %>
    </div>
  </div>
  <% 
  end if
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Product of the Month
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Featured Products
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  
Public Sub pcs_FeaturedProducts  
    If pcIntHPFeaturedCount > 0 Then 
        pIntProductColumns = scPrdRow
        if pcIntHPFirst<>0 then
            wstart=1
        else
            wstart=0
        end if
        If (pcv_intProductCount > 0) And (pcIntHPFeaturedCount > wstart)  Then
      %>
        <!-- Featured Products -->
        <div id="pcFeaturedProducts">
        <h2><%=dictLanguage.Item(Session("language")&"_mainIndex_7")%></h2>
        <%
        pMoreLinkTarget = "showfeatured.asp"
        pMoreLinkText = dictLanguage.Item(Session("language")&"_mainIndex_13") & dictLanguage.Item(Session("language")&"_mainIndex_7")
        pCnt = pcShowProducts(pcIntHPFeaturedCount, wstart)
        If pCnt = 0 Then
            %>
            <span class="pcFeaturedProductsMessage"><%= dictLanguage.Item(Session("language")&"_mainIndex_14") %></span>
            <%
        End If
        %>
        </div>
        <%
        End If
    End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END:  Featured Products
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  
Public Sub pcs_ShowSlideShow()
  call pcs_ShowSlideShowID(1)
End Sub

Public Sub pcs_ShowSlideShowMobile()
  call pcs_ShowSlideShowID(2)
End Sub

Dim pcv_intSlideCount
Public Sub pcs_ShowSlideShowID(settingID)
  pcv_intSlideCount=0
    
  '// Check if we need to use the default settings (Setting ID 1)
  query = "SELECT idSetting, useDefault FROM pcSlideShowSettings WHERE idSetting = " & settingID & ";"
  set rsSSS = conntemp.execute(query)
  if not rsSSS.eof then
    idSetting = rsSSS("idSetting")
    useDefault = rsSSS("useDefault")

    if useDefault = 1 then
      settingID = 1
    end if
  end if
  set rsSSS = nothing

  '// Get settings for the selected slideshow
  query = "SELECT effect, pauseTime, animSpeed FROM pcSlideShowSettings WHERE idSetting = " & settingID
  set rsSSS = conntemp.execute(query)
  if not rsSSS.eof then
    slideEffect = rsSSS("effect")
    slidePauseTime = rsSSS("pauseTime")
    slideAnimSpeed = rsSSS("animSpeed")
    
    if slidePauseTime < 500 or slidePauseTime > 120000 then
      slidePauseTime = 0
    end if
    
    if slideAnimSpeed < 0 or slideAnimSpeed > 5000 then
      slideAnimSpeed = 0
    end if
  end if
  set rsSSS = nothing
  
  '// Get slides
  query = "SELECT slideImage, slideCaption, slideUrl, slideAlt, slideNewWindow FROM pcSlideShow WHERE GETDATE() BETWEEN slideStart AND datediff(d,0, slideEnd) AND idSetting = " & settingID & " ORDER BY slideOrder, idSlide"
  set rsSlides=server.CreateObject("ADODB.Recordset")
  set rsSlides=conntemp.execute(query)
  if Err.number <> 0 then
        call LogErrorToDatabase()
        set rsSlides = Nothing
        call closeDb()
        response.redirect "techErr.asp?err="&pcStrCustRefID
  end if
  if NOT rsSlides.eof then
    pcArray_Slides = rsSlides.getRows()
    pcv_intSlideCount = UBound(pcArray_Slides,2)+1
  end if
  set rsSlides = nothing
%>
<% If pcv_intSlideCount>0 Then %>
  
  <script src="../includes/javascripts/flickity/flickity.pkgd.min.js"></script>

  <div class="slider-wrapper theme-productcart hero-carousel" data-flickity='{ "freeScroll": true, "wrapAround": true, "autoPlay": <%=slidePauseTime%>, "lazyLoad": true }' data-js="hero-carousel">
      
      <%
      for i = 0 to pcv_intSlideCount - 1
          'Get the slide information
          slideName = pcArray_Slides(0, i)
          slideCaption = pcArray_Slides(1, i)
          slideUrl = pcArray_Slides(2, i)
          slideAlt = pcArray_Slides(3, i)
          slideNewWindow = pcArray_Slides(4, i)

          'Check if the slide has an image'
          if len(slideAlt) < 1 then
            slideAlt = "Slideshow image " & i + 1
          end if

          'Check if the link should open in a new window'
          if slideNewWindow = 1 then
            slideNewWindow = " target=""_blank"""
          else
            slideNewWindow = ""
          end if

          'Create a numbered cell for the slide
          %><div class="hero-carousel__cell hero-carousel__cell--<%=i+1%>"><div class="hero-carousel__cell__content"><%
          
            'Check if the slide has a link and setup a link with the information we learned about linking in a new window.
            if len(slideUrl) > 0 then
              %><a href="<%= slideUrl %>" <%= slideNewWindow%>><%
            end if

            'Now the image.
            %><img src="<%=pcf_getImagePath("catalog",slideName)%>" alt="<%= slideAlt %>" title="<%= slideCaption %>" /><%

            'Close the link if we need to do such a thing.
            if len(slideUrl) > 0 then
              %></a><%
            end if

            'Close the div so we can have another.'
            %></div></div><%
        next
        %>
        </div>
    </div>
    
<% End If %> 
<%
End Sub


Public Sub pcs_ShowProductL(pIdProduct, pSku, pSmallImageUrl, pDescription)

        pIdCategoryTemp = getFirstCategoryID(pIdProduct, pIdCategory)
  
    '// Call SEO Routine
    pcGenerateSeoLinks

    if len(pDescription)>22 then
      pDescription=left(pDescription,19) & "..."
    end if
    
    If pSmallImageUrl = "" Then
      pSmallImageUrl = "no_image.gif"
    End If
  %>
  
    <li class="pcShowProductsUL">
      <% if pShowSmallImg <> 0 And pSmallImageUrl<>"" then%>
        <div class="pcShowProductImage">
          <a href="<%=Server.HtmlEncode(pcStrPrdLink)%>"><img src="<%=pcf_getImagePath("catalog",pSmallImageUrl)%>" alt="<%= pDescription %>"></a>
        </div>
      <% end if %>
      
      <div class="pcShowProductInfo">
        <div class="pcShowProductName">
            <a href="<%=Server.HtmlEncode(pcStrPrdLink)%>"><%=pDescription%></a>
        </div>
        <%if pShowSKU <> 0 then%>
          <div class="pcShowProductSku">
            <%=pSku%>
          </div>
        <% end if %>
        <%if pShowPrice <> 0 then%>
          <div class="pcShowProductPrice">
            <%= dictLanguage.Item(Session("language")&"_prdD1_1") & ": " & scCursign & money(pPrice) %>
          </div>
        <% end if %>
      </div>
      
      <div class="pcClear"></div>
    </li>
  <%
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START:  Best sellers, new arrivals, specials
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  
Public Sub pcs_ShowProducts
    on error resume next
    %>
    <div id="pcFeaturedProductsList" class="pcShowProducts">
  <%
  '// Specials
  If pcIntHPSpcCount > 0 Then

    Dim pcIntSpecialsNFS
    pcIntSpecialsNFS = 0 ' Not for sale items are shown
    pcIntSpecialsNFS = -1 ' Not for sale items are not shown
    if pcIntSpecialsNFS <> 0 and NotForSaleOverride(session("customerCategory"))=0 then
      queryNFS = "AND formQuantity=0 "
    else
      queryNFS = ""
    end if
    '*******************************
    ' GET sorting criteria
    '*******************************
    Dim querySort
    querySort = " ORDER BY products.description Asc"  
    '*******************************
    ' GET Specials from DB
    '*******************************
    if session("CustomerType")<>"1" then
      query1= " AND categories.pccats_RetailHide=0"
    else
      query1=""
    end if

    query="SELECT distinct products.idProduct,products.sku,products.description,products.smallImageUrl FROM products,categories_products,categories WHERE products.active=-1 AND products.hotdeal=-1 AND products.configOnly=0 AND products.removed=0 " & queryNFS & " AND categories_products.idProduct=products.idProduct AND categories.idCategory=categories_products.idCategory AND categories.iBTOhide=0 " & query1 & querySort 
    set rsProducts=server.CreateObject("ADODB.Recordset")
    set rsProducts=conntemp.execute(query)
    if err.number<>0 then
      call LogErrorToDatabase()
      set rsProducts=nothing
      call closedb()
      response.redirect "techErr.asp?err="&pcStrCustRefID
    end if
    pcv_intProductCount=-1
    if NOT rsProducts.eof then
      pcArray_Products = rsProducts.getRows()
      pcv_intProductCount = UBound(pcArray_Products,2)+1
    end if
    set rsProducts = nothing
    
    showVMLink = 0
        If (pcv_intProductCount > 0) And (pcIntHPSpcCount > 0) Then
            If  (pcv_intProductCount) <= (pcIntHPSpcCount) then
                gotoCount = pcv_intProductCount                
                showVMLink = 0
            Else
                gotoCount = pcIntHPSpcCount
                showVMLink = 1
            End If
        End If
    %>
      <div id="pcSpecials">
        <h2><%=dictLanguage.Item(Session("language")&"_mainIndex_4")%></h2>
        <%
        If gotoCount > 0 Then
          %>
          <ul>
            <%
            'Loop until the total number of products to show
            count=0
          
            tCnt=Cint(0)
          
            do while (tCnt < gotoCount)
              pIdProduct = pcArray_Products(0,tCnt)
              pSku = pcArray_Products(1,tCnt)                         
              pDescription = pcArray_Products(2,tCnt)
              pSmallImageUrl = Server.HtmlEncode(pcArray_Products(3,tCnt))
              pDesc = pDescription
            
              tCnt=tCnt+1
              if count < pcv_intProductCount then        
                  '// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories
                
                pcs_ShowProductL pIdProduct, pSku, pSmallImageUrl, pDescription
              end if
              iRecordsShown=iRecordsShown + 1
              count=count + 1
            loop
            %>
          </ul>
        
          <% if showVMLink = 1 then %>
            <a class="pcMoreLink" href="showspecials.asp"><%=dictLanguage.Item(Session("language")&"_mainIndex_13")%><%=dictLanguage.Item(Session("language")&"_mainIndex_4")%></a>
          <% end if %>
        <% Else %>
          <span class="pcFeaturedProductsMessage"><%= dictLanguage.Item(Session("language")&"_mainIndex_15") %></span>
        <%
        End If 'end gotoCount > 0
        %>
      </div>
    <%
    End If 'end pcIntHPSpcCount > 0


        '// New Arrivals 
    If pcIntHPNewCount > 0 Then

            query="SELECT pcNAS_NDays, pcNAS_NotForSale, pcNAS_OutOfStock FROM pcNewArrivalsSettings;"
            set rs=Server.CreateObject("ADODB.RecordSet")
            set rs=connTemp.execute(query)                                    
            if not rs.eof then
        pcNDays=rs("pcNAS_NDays")
                pcIntNewArrNFS=rs("pcNAS_NotForSale")
                pcIntNewArrInStock=rs("pcNAS_OutOfStock")
      end if
      set rs=nothing

      if isNULL(pcNDays) OR (pcNDays="0") OR (pcNDays="") then
        pcNDays=15
      end if
            if pcIntNewArrNFS <> 0 and NotForSaleOverride(session("customerCategory"))=0 then
        queryNFS = "((products.formQuantity)=0) AND "
            else
        queryNFS = " "
            end if
            '*******************************
            ' GET new arrivals from DB
            '*******************************
            pcTodayDate=Date()
            if SQL_Format="1" then
        pcTodayDate=Day(pcTodayDate)&"/"&Month(pcTodayDate)&"/"&Year(pcTodayDate)
      else
        pcTodayDate=Month(pcTodayDate)&"/"&Day(pcTodayDate)&"/"&Year(pcTodayDate)
      end if

      y="'"
      if session("CustomerType")<>"1" then
        query1= " AND ((categories.pccats_RetailHide)=0)"
      else
        query1=""
      end if
    
      if pcIntNewArrInStock <> 0 then
        query="SELECT distinct products.idProduct, products.sku, products.description, products.smallImageUrl,  products.formQuantity, products.pcprod_EnteredOn FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.stock)>0) AND " & queryNFS & "((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&cdate(pcTodayDate)&y&"-convert(datetime, [products].[pcprod_EnteredOn],101))<="& pcNDays &") AND ((categories.iBTOhide)=0) AND ((categories.pccats_RetailHide)=0)) OR (((products.noStock)=-1) AND " & queryNFS & "((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&pcTodayDate&y&"-[products].[pcprod_EnteredOn])<="& pcNDays &") AND ((categories.iBTOhide)=0) AND ((categories.pccats_RetailHide)=0)) ORDER BY products.pcprod_EnteredOn DESC;"
      else
        query="SELECT distinct products.idProduct, products.sku, products.description, products.smallImageUrl,  products.formQuantity, products.pcprod_EnteredOn FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE ("&queryNFS&"((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND (("&y&pcTodayDate&y&"-convert(datetime, [products].[pcprod_EnteredOn],101))<="& pcNDays &") AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.pcprod_EnteredOn DESC;"
      end if
      set rsProducts=server.CreateObject("ADODB.Recordset")
      set rsProducts=conntemp.execute(query)
      if err.number<>0 then
        call LogErrorToDatabase()
        set rsProducts=nothing
        call closedb()                
        response.redirect "techErr.asp?err="&pcStrCustRefID
      end if
      pcv_intProductCount=-1
      if NOT rsProducts.eof then
        pcArray_Products = rsProducts.getRows()
        pcv_intProductCount = UBound(pcArray_Products,2)+1
      end if
      set rsProducts = nothing
    
      showVMLink = 0
      gotoCount = pcv_intProductCount
      If  (pcIntHPNewCount) < (pcv_intProductCount) then
        gotoCount = pcIntHPNewCount
        showVMLink = 1
      End If
      %>
        <div id="pcNewArrivals">
          <h2><%=dictLanguage.Item(Session("language")&"_mainIndex_10")%></h2>
          <%
          If gotoCount > 0 Then
            %>
            <ul>
            <%
            'Loop until the total number of products to show
            count=0
            tCnt=0        
            do while (tCnt < gotoCount)           
              pIdProduct = pcArray_Products(0,tCnt)
              pSku = pcArray_Products(1,tCnt)                         
              pDescription = pcArray_Products(2,tCnt)
              pSmallImageUrl = Server.HtmlEncode(pcArray_Products(3,tCnt))
                            EnteredOnDate = pcArray_Products(5,tCnt)
              pDesc = pDescription
      
              tCnt=tCnt+1
              if count < cint(pcv_intProductCount) then        
                '// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories  
                pcs_ShowProductL pIdProduct, pSku, pSmallImageUrl, pDescription
              end if
              iRecordsShown=iRecordsShown + 1
              count=count + 1
            loop
            %>
            </ul>
            <% if showVMLink = 1 then %>
              <a class="pcMoreLink" href="shownewarrivals.asp"><%=dictLanguage.Item(Session("language")&"_mainIndex_13")%><%=dictLanguage.Item(Session("language")&"_mainIndex_10")%></a>
            <% end if %>
          <% Else %>
            <span class="pcFeaturedProductsMessage"><%= dictLanguage.Item(Session("language")&"_mainIndex_16") %></span>
          <% End If 'end gotoCount > 0 %>
        </div>
      <%
    End If 'end pcIntHPNewCount > 0

    If pcIntHPBestCount > 0 Then
      Dim pcIntBestSellNFS, pcIntBestSellInStock, pcIntBestSellSales
            pcIntBestSellSales=0

            query="SELECT pcBSS_BestSellCount,pcBSS_Style,pcBSS_PageDesc,pcBSS_NSold,pcBSS_NotForSale,pcBSS_OutOfStock,pcBSS_SKU,pcBSS_ShowImg FROM pcBestSellerSettings;"
            set rs=connTemp.execute(query)
            if not rs.eof then
                pcIntBestSellSales=rs("pcBSS_NSold")
                pcIntBestSellNFS=rs("pcBSS_NotForSale")
                pcIntBestSellInStock=rs("pcBSS_OutOfStock")
            end if
            set rs=nothing
            
            if isNULL(pcIntBestSellSales) or (pcIntBestSellSales="0") then
                 pcIntBestSellSales=2
            end if
            if pcIntBestSellNFS<> 0 and NotForSaleOverride(session("customerCategory"))=0 then
                queryNFS = " AND ((products.formQuantity)=0)"
            else
                queryNFS = " "
            end if
            '*******************************
            ' GET Best Sellers from DB
            '*******************************
            if session("CustomerType")<>"1" then
                query1= " AND ((categories.pccats_RetailHide)=0)"
            else
                query1=""
            end if

      if pcIntBestSellInStock<> 0 then
        query="SELECT distinct products.idProduct, products.sku, products.description, products.smallImageUrl, products.sales, products.formQuantity FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.stock)>0) AND ((products.sales)>="&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") OR (((products.noStock)=-1) AND ((products.sales)>"&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.sales DESC;"
      else
        query="SELECT distinct products.idProduct, products.sku, products.description, products.smallImageUrl, products.sales, products.formQuantity FROM (products INNER JOIN categories_products ON products.idProduct = categories_products.idProduct) INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE (((products.sales)>="&pcIntBestSellSales&")"&queryNFS&" AND ((products.active)=-1) AND ((products.configOnly)=0) AND ((products.removed)=0) AND ((categories.iBTOhide)=0)"&query1&") ORDER BY products.sales DESC;"
      end if
      set rsProducts=server.CreateObject("ADODB.Recordset")
      set rsProducts=conntemp.execute(query)
      if err.number<>0 then
        call LogErrorToDatabase()
        set rsProducts=nothing
        call closedb()
        response.redirect "techErr.asp?err="&pcStrCustRefID
      end if
      pcv_intProductCount=-1
      if NOT rsProducts.eof then
        pcArray_Products = rsProducts.getRows()
        pcv_intProductCount = UBound(pcArray_Products,2)+1
      end if
      set rsProducts = nothing
    
      showVMLink = 0
      gotoCount = pcv_intProductCount
      If  (pcIntHPBestCount) < (pcv_intProductCount) then
        gotoCount = pcIntHPBestCount
        showVMLink = 1
      End If
      %>
      <div id="pcBestSellers">
          <h2><%=dictLanguage.Item(Session("language")&"_mainIndex_6")%></h2>
        <%
        If gotoCount > 0 Then
            %>
                    <ul>
                        <%
                        'Loop until the total number of products to show
                        count=0                        
                        tCnt=Cint(0)                        
                        do while (tCnt < gotoCount)
                            pIdProduct = pcArray_Products(0,tCnt)
                            pSku = pcArray_Products(1,tCnt)                         
                            pDescription = pcArray_Products(2,tCnt)
                            pSmallImageUrl = Server.HtmlEncode(pcArray_Products(3,tCnt))
                            pDesc = pDescription
                            
                            tCnt=tCnt+1
                            if count < pcv_intProductCount then        
                                '// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories                                
                                pcs_ShowProductL pIdProduct, pSku, pSmallImageUrl, pDescription
                            end if
                            iRecordsShown=iRecordsShown + 1
                            count=count + 1
                        loop
                        %>
                    </ul>
                    <% if showVMLink = 1 then %>
                        <a class="pcMoreLink" href="showbestsellers.asp"><%=dictLanguage.Item(Session("language")&"_mainIndex_13")%><%=dictLanguage.Item(Session("language")&"_mainIndex_6")%></a>
                    <% end if %>
                <% Else %>
                    <span class="pcFeaturedProductsMessage"><%= dictLanguage.Item(Session("language")&"_mainIndex_17") %></span>
                <% End If 'end gotoCount > 0 %>
            </div>
            <%
        End If 'end pcIntHPBestCount > 0
        %>
        <div class="pcClear"></div>
  </div>
    <%
End Sub
%>
