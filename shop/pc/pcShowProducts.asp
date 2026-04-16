<%
Dim pIdProduct, pSku, pDescription, pStock, pNoStock, pSmallImageUrl, pFormQuantity, pserviceSpec, pcv_intBackOrder
Dim pPrice, pListHidden, pListPrice, pBtoBPrice, psDesc

Dim pMoreLinkTarget, pMoreLinkText

pMoreLinkTarget = ""
pMoreLinkText = ""
	
%>

<% If pcPageStyle = "l" Then %>
	<!--#include file="pcShowProductL.asp" -->
<% End If %>

<% If pcPageStyle = "m" Then %>
	<!--#include file="pcShowProductM.asp" -->
<% End If %>

<%
Function pcShowProducts(maxProducts, wstart)
	If pcPageStyle = "m" Then
		pcShowProductsMBefore
	End If

        '*******************************
        ' Add table headers for display
        ' styles L and M
        '*******************************
        %>        
        <% 
        If pcPageStyle = "l" Then
              pcShowProductsLTop	
        ElseIf pcPageStyle = "m" Then 								
              pcShowProductsMTop pcArray_Products, pcv_intProductCount, pagesize
        Else
        End If 
        %>        
        <%
        '*******************************
        ' End table headers
        '*******************************
      
        '*******************************
        ' Load product information
        ' Loop through the products
        '*******************************
        
        'Set the product count to zero
        count=0
    
        tCnt=Cint(0)
        maxProducts=cint(maxProducts)
    
        colCount = pIntProductColumns
        If maxProducts < colCount Then
          colCount = maxProducts
        End If

        If (pcv_intProductCount > 0) And (maxProducts > wstart)  Then
            
            For tCnt = wstart To pcv_intProductCount-1
        
                pidProduct=pcArray_Products(0,tCnt) '// rs("idProduct")
                pSku=pcArray_Products(1,tCnt) '// rs("sku")
                pDescription=pcArray_Products(2,tCnt) '// rs("description")   
                pPrice=pcArray_Products(3,tCnt) '// rs("price")
                pListHidden=pcArray_Products(4,tCnt) '// rs("listhidden")
                pListPrice=pcArray_Products(5,tCnt) '// rs("listprice")						   
                pserviceSpec=pcArray_Products(6,tCnt) '// rs("serviceSpec")
                pBtoBPrice=pcArray_Products(7,tCnt) '// rs("bToBPrice")   
                pSmallImageUrl=pcArray_Products(8,tCnt) '// rs("smallImageUrl")   
                pnoprices=pcArray_Products(9,tCnt) '// rs("noprices")
                If isNULL(pnoprices) OR pnoprices="" Then
                    pnoprices=0
                End If
                pStock=pcArray_Products(10,tCnt) '// rs("stock")
                pNoStock=pcArray_Products(11,tCnt) '// rs("noStock")
                pcv_intHideBTOPrice=pcArray_Products(12,tCnt) '// rs("pcprod_HideBTOPrice")
                If isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" Then
                    pcv_intHideBTOPrice="0"
                End If
                If pnoprices=2 Then
                    pcv_intHideBTOPrice=1
                End If
                pcv_intBackOrder=pcArray_Products(13,tCnt) '// rs("pcProd_BackOrder")
                pFormQuantity=pcArray_Products(14,tCnt) '// rs("FormQuantity")
                pidrelation=pcArray_Products(0,tCnt) '// rs("idProduct")						
                psDesc=Trim(pcArray_Products(16,tCnt)) '// rs("sDesc")
                pcv_intSkipDetailsPage=Trim(pcArray_Products(23,tCnt))  
                if isNull(pcv_intSkipDetailsPage) or pcv_intSkipDetailsPage="" then  
                    pcv_intSkipDetailsPage=0
                end if	
				pUrl=pcArray_Products(24,tCnt)
                %>
                <!--#include file="pcGetPrdPrices.asp"-->
                <%
              
                '*******************************
                ' Show product information
                ' depEnding on the page style
                '*******************************
    
                ' FIRST STYLE - Show products horizontally, with images
                If pcPageStyle = "h" Then
				
				'Determine which product listing styles we are using based on page we are on.
				Select Case Request.ServerVariables("PATH_INFO")
					Case "/shop/pc/CUSTOMCAT-bundles1.asp", "/shop/pc/CUSTOMCAT-bundles2.asp"
                    	%>
                        <!--#include file="pcShowProduct-Bundle.asp" -->
                    	<%
					Case "/shop/pc/CUSTOMCAT-bundles3.asp"
						   'DA EDit - Attempt to skip Pro PC if we have QHD Screen
							If not pSku = "MM-PRO1" then
								%>
								<!--#include file="pcShowProduct-BundlePC.asp" -->
								<%
							else 
								if not bunBCmid = 342 then
								%>
								<!--#include file="pcShowProduct-BundlePC.asp" -->
								<%
								end if
							end if
                    	%>
                        
                    	<%
					Case "/shop/pc/CUSTOMCAT-arrays1.asp", "/shop/pc/CUSTOMCAT-arrays2.asp"
                    	%>
                        <!--#include file="pcShowProduct-Array.asp" -->
                    	<%
					Case Else
                    	%>
                        <!--#include file="pcShowProduct-standard.asp" -->
                    	<%
				End Select
				
                    If ((CInt(tCnt+1) Mod colCount = 0) And wstart=0) OR ((CInt(tCnt) Mod colCount = 0) And wstart=1) Then
                    
                    '<div class="pcRowClear"></div>
                    
                    End If

                End If
              
                ' SECOND STYLE - Show products vertically, with images 
                If pcPageStyle = "p" Then	
                    %>
                    <div class="pcRow">
                        <!--#include file="pcShowProductP.asp" -->
                    </div>
                    <% 
                End If
                
                ' THIRD STYLE - Show a list of products, with a small image 
                If pcPageStyle = "l" Then	
                    call pcShowProductL(dblpcCC_Price)
                End If
                
                
                ' FOURTH STYLE - Show a list of products, with multiple add to cart 
                If pcPageStyle = "m" Then
                    call pcShowProductM(dblpcCC_Price)
                End If
                        
                If tCnt+1 >= maxProducts Then Exit For

            Next

        End If
                
        If pcPageStyle = "l" Then
            pcShowProductsLBottom
        ElseIf pcPageStyle = "m" Then
            pcShowProductsMBottom
        End If 

		 if (pMoreLinkTarget <> "") AND (Clng(pcv_intProductCount) > Clng(maxProducts)) then %>
            <div class="pcMoreLinkWrapper">
                <a class="pcMoreLink" href="<%= pMoreLinkTarget %>"><%= pMoreLinkText %></a>
            </div>
        <% end if %>

    <% 
    If pcPageStyle = "m" Then
        pcShowProductsMAfter
    End If
    pcShowProducts = tCnt + 1
End Function 
%>