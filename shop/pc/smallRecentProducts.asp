<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'// START - Show Recently Viewed Products

'// Set maximum products to show
MaxShowRecentProducts=6

ViewedPrdList=getUserInput2(Request.Cookies("pcfront_visitedPrds"),0)
IF ViewedPrdList<>"" AND ViewedPrdList<>"*" THEN
	
	tmpViewedList=split(ViewedPrdList,"*")
	ViewedPrdList=""
	tmpIndex=0
	tmpIndex1=0
	pcv_ValidateList=0
	pcv_ValidateFailAll=0
	Do While (tmpIndex<ubound(tmpViewedList)) AND (tmpIndex1+1<=MaxShowRecentProducts)		
		pcv_EvalViewedPrd = tmpViewedList(tmpIndex)		
		if pcv_EvalViewedPrd="" OR validNum(pcv_EvalViewedPrd) then
			pcv_ValidateList=1
		else
			pcv_ValidateFailAll=1
		end if
		if tmpViewedList(tmpIndex)<>"" then
			if ViewedPrdList<>"" then
				ViewedPrdList=ViewedPrdList & ","
			end if
			ViewedPrdList=ViewedPrdList & tmpViewedList(tmpIndex)
			tmpIndex1=tmpIndex1+1
		end if
		tmpIndex=tmpIndex+1
	Loop
	
	tmpViewedList=split(ViewedPrdList,",")
		
	IF pcv_ValidateList=1 AND pcv_ValidateFailAll=0 AND len(ViewedPrdList)>0 THEN '// The cookie was NOT modified or corrupted
	
		Set connTemp6=Server.CreateObject("ADODB.Connection")
		connTemp6.Open scDSN
		query4="SELECT products.idproduct,products.description,products.sku,products.smallImageUrl,products.pcUrl FROM Products WHERE idproduct IN (" & ViewedPrdList & ");"
		set rs6=connTemp6.execute(query4)
		IF err.number<>0 THEN
			set rs6 = nothing
			set connTemp6=nothing
			Response.Write(ViewedPrdList)
		ELSE
		
			IF NOT rs6.eof THEN
			'DA - EDIT
			%>

					<%
					tmpVPrdArr=rs6.getRows()
					set rs6=nothing
					tmpVPrdCount=ubound(tmpVPrdArr,2)
					For tmpIndex2=0 to tmpIndex1-1
						For tmpIndex=0 to tmpVPrdCount
							if CLng(tmpVPrdArr(0,tmpIndex))=CLng(tmpViewedList(tmpIndex2)) then
							
								' Get product image and sku
								pcvStrSku = tmpVPrdArr(2,tmpIndex)
								pcvStrSmallImage = tmpVPrdArr(3,tmpIndex)

								'Show SKU?
								pcIntShowSKU = 0
								
								'Clean up product name
								pcvStrSPtitle=""
								pcvStrSPname=ClearHTMLTags2(tmpVPrdArr(1,tmpIndex),0)
								if len(pcvStrSPname)>41 then
									pcvStrSPtitle=pcvStrSPname
									pcvStrSPname=left(pcvStrSPname,40) & "..."
								end if

								'prdLink = pcGenerateSeoProductLink(tmpVPrdArr(1,tmpIndex), "", tmpVPrdArr(0,tmpIndex)) 
								prdLink = "/products/" & tmpVPrdArr(4,tmpIndex) & "/"
								
							%>
							<li>              
								<a href="<%= Server.HtmlEncode(prdLink) %>" title="<%= pcvStrSPtitle %>">
                                    <%= pcvStrSPname %></a>
							</li>
								<%exit for
							end if
						Next
					Next%>

			<%
			END IF ' Empty recordset
		END IF ' Any errors
		set rs6=nothing
		set connTemp6=nothing
	
	END IF ' Valid cookie
	
END IF ' Product list exists

'// END - Show Recently Viewed Products
%>
