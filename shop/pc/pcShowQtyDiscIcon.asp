<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

		' *******************************
		' * START quantity discounts
		' *******************************	
			if pnoprices=0 then
				' check for discount per quantity
				query="SELECT idDiscountperquantity FROM discountsperquantity WHERE idproduct=" & pidProduct
				if session("CustomerType")<>"1" then
					query=query & " and discountPerUnit<>0"
					else
					query=query & " and discountPerWUnit<>0"
				end if
				set rsDisc=Server.CreateObject("ADODB.Recordset")
				set rsDisc=conntemp.execute(query)
				
				if not rsDisc.eof then
					pDiscountPerQuantity=-1
					else
					pDiscountPerQuantity=0
				end if
			set rsDisc = nothing
			end if
				
			if pDiscountPerQuantity=-1 then %>
				<script type=text/javascript>
					function win(fileName)
					{
						myFloater=window.open('','myWindow','scrollbars=auto,status=no,width=300,height=250')
						myFloater.location.href=fileName;
					}
				</script>
				<div class="pcShowProductQtyDiscount">
					<a href="javascript:openbrowser('priceBreaks.asp?type=<%=Session("customerType")%>&amp;idproduct=<%=pidProduct%>&amp;type=1')">
						<img src="<%=pcf_getImagePath("",rsIconObj("discount"))%>" alt="<%response.write dictLanguage.Item(Session("language")&"_viewPrd_16")%>">
					</a>
				</div>
    		<% 
		else

			query="SELECT pcProd_Apparel FROM Products WHERE idproduct=" & pidProduct & " AND pcProd_Apparel=1;"
			set rsQ=connTemp.execute(query)
			pDiscountPerQuantity=0
			if not rsQ.eof then

				query="SELECT idDiscountperquantity FROM discountsperquantity INNER JOIN Products ON discountsperquantity.idProduct=Products.idProduct WHERE Products.pcProd_ParentPrd=" &pidProduct
				set rsQ2=server.CreateObject("ADODB.RecordSet")
				set rsQ2=conntemp.execute(query)
			
				if not rsQ2.eof then
					pDiscountPerQuantity=-1
				else
					pDiscountPerQuantity=0
				end if
				set rsQ2=nothing

			end if
			set rsQ=nothing

			if pDiscountPerQuantity=-1 then %>
				<script language="JavaScript">
					<!--
						function optwin(fileName)
						{
						myFloater = window.open('','myWindow','scrollbars=yes,status=no,width=400,height=300')
						myFloater.location.href = fileName;
						}
					//-->
				</script>
				<a href="javascript:openbrowser('app-subPrdDiscount.asp?idproduct=<%=pidProduct%>')"><img src="<%=pcf_getImagePath("",rsIconObj("discount"))%>" alt="<%response.write dictLanguage.Item(Session("language")&"_viewPrd_16")%>" style="vertical-align: middle"></a>
			<%end if
	
		end if
		
		' *******************************
		' * END quantity discounts
		' *******************************

		' *******************************
		' * START Promotion
		' *******************************	
			if pnoprices=0 then
					
					query="SELECT pcPrdPro_id,idproduct,pcPrdPro_QtyTrigger,pcPrdPro_DiscountType,pcPrdPro_DiscountValue,pcPrdPro_ApplyUnits,pcPrdPro_PromoMsg,pcPrdPro_ConfirmMsg,pcPrdPro_SDesc,pcPrdPro_IncExcCust,pcPrdPro_IncExcCPrice,pcPrdPro_RetailFlag,pcPrdPro_WholesaleFlag FROM pcPrdPromotions WHERE pcPrdPro_Inactive=0 AND idproduct=" & pidProduct & ";"
					set rsPromo=connTemp.execute(query)
					if not rsPromo.eof then
						PrdPromoArr=rsPromo.getRows()
						set rsPromo=nothing
						PrdPromoCount=ubound(PrdPromoArr,2)
						
						tmpIDCode=PrdPromoArr(0,0)
						tmpIDProduct=PrdPromoArr(1,0)
						tmpQtyTrigger=clng(PrdPromoArr(2,0))
						tmpDiscountType=PrdPromoArr(3,0)
						tmpDiscountValue=PrdPromoArr(4,0)
						tmpApplyUnits=PrdPromoArr(5,0)
						tmpConfirmMsg=PrdPromoArr(7,0)
						tmpDescMsg=PrdPromoArr(8,0)
						pcIncExcCust=PrdPromoArr(9,0)
						pcIncExcCPrice=PrdPromoArr(10,0)
						pcv_retail=PrdPromoArr(11,0)
						pcv_wholeSale=PrdPromoArr(12,0)
						
						pcv_Filters=0
						pcv_FResults=0
						'Filter by Customers
						pcv_CustFilter=0
						query="select IDCustomer from PcPPFCusts where pcPrdPro_id=" & tmpIDCode
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if		
						if not rs.eof then
							pcv_Filters=pcv_Filters+1
							pcv_CustFilter=1
						end if
						set rs=nothing
						
						if pcv_CustFilter=1 then
								
						query="select IDCustomer from PcPPFCusts where pcPrdPro_id=" & tmpIDCode & " and IDCustomer=" & session("IDCustomer")
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if		
						if not rs.eof then
							if (pcIncExcCust="0") then
								pcv_FResults=pcv_FResults+1
							end if
						else
							if (pcIncExcCust="1") then
								pcv_FResults=pcv_FResults+1
							end if
						end if
						set rs=nothing
						
						end if
						'End of Filter by Customers
						
						
						'Filter by Customer Categories
						pcv_CustCatFilter=0
						
						query="select idCustomerCategory from pcPPFCustPriceCats where pcPrdPro_id=" & tmpIDCode
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if		
						if not rs.eof then
							pcv_Filters=pcv_Filters+1
							pcv_CustCatFilter=1
						end if
						set rs=nothing
						
						if pcv_CustCatFilter=1 then
								
						query="select pcPPFCustPriceCats.idCustomerCategory from pcPPFCustPriceCats, Customers where pcPPFCustPriceCats.pcPrdPro_id=" & tmpIDCode & " and pcPPFCustPriceCats.idCustomerCategory = Customers.idCustomerCategory and Customers.idcustomer=" & session("IDCustomer")
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if		
						if not rs.eof then
							if (pcIncExcCPrice="0") then
								pcv_FResults=pcv_FResults+1
							end if
						else
							if (pcIncExcCPrice="1") then
								pcv_FResults=pcv_FResults+1
							end if
						end if
						set rs=nothing
						
						end if
						'End of Filter by Customer Categories
						
						' Check to see if promotion is filtered by reatil or wholesale.
						if (pcv_retail ="0" and pcv_wholeSale ="1") or (pcv_retail ="1" and pcv_wholeSale ="0") Then
							pcv_Filters=pcv_Filters+1
							if pcv_wholeSale = "1" and session("customertype") = 1 then
								pcv_FResults=pcv_FResults+1		
							end if 
							if pcv_retail = "1" and session("customertype") <> 1 Then
								pcv_FResults=pcv_FResults+1
							end if    
						end if
						
						if (pcv_Filters=pcv_FResults) AND PrdPromoArr(6,0)<>"" then%>
							<div class="pcShowProductPromoDiscount">
								<img src="<%=pcf_getImagePath("images","pc4_promo_icon_small.png")%>" alt="<%response.write ClearHTMLTags2(PrdPromoArr(6,0),0)%>" title="<%response.write ClearHTMLTags2(PrdPromoArr(6,0),0)%>">
							</div>
						<%end if
					end if
					set rsPromo=nothing
					
			end if
		
		' *******************************
		' * END PROMOTION
		' *******************************
%>