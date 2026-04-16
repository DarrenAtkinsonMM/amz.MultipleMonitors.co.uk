<%'APP-S
	call opendb()
	
	query="SELECT Customers.customerType FROM Customers INNER JOIN Orders ON Customers.idcustomer=Orders.idcustomer WHERE Orders.idOrder=" & qryID & ";"
	set rsQ=connTemp.execute(query)
	
	tmpCustomerType=0
	
	if not rsQ.eof then
		tmpCustomerType=rsQ("customerType")
		if tmpCustomerType<>"1" OR IsNull(tmpCustomerType) OR tmpCustomerType="" then
			tmpCustomerType=0
		end if
	end if
	set rsQ=nothing
	
	query="SELECT Products.idproduct,Products.sku,ProductsOrdered.quantity,ProductsOrdered.unitPrice,ProductsOrdered.QDiscounts,Products.pcprod_ParentPrd FROM ProductsOrdered INNER JOIN Products ON ProductsOrdered.idproduct=Products.idproduct WHERE ProductsOrdered.idOrder=" & qryID & " AND Products.pcprod_ParentPrd>0;"
	set rsQ=connTemp.execute(query)
	IF NOT rsQ.eof then
	
	pcArrSub=rsQ.getRows()
	set rsQ=nothing
	intCountSub=ubound(pcArrSub,2)

	pcv_ParentPrd=0
	tmpSaveParent="||"
	for f=0 to intCountSub
				pcv_ParentPrd=pcArrSub(5,f)
				if instr(tmpSaveParent,"||" & pcv_ParentPrd & "||")=0 then
					tmpSaveParent=tmpSaveParent & pcv_ParentPrd & "||"
					tmpQuantity=0
					For k=0 to intCountSub
						if clng(pcArrSub(5,k))=pcv_ParentPrd then
							tmpQuantity=tmpQuantity+cLng(pcArrSub(2,k))
						end if
					Next
	
					disTotalQuantity=tmpQuantity
					
					query="SELECT discountPerUnit,discountPerWUnit,percentage,baseproductonly FROM discountsPerQuantity WHERE idProduct=" & pcv_ParentPrd & ";"
					set rsQ1=server.CreateObject("ADODB.RecordSet")
					set rsQ1=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rsQ1=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					if not rsQ1.eof then
						For k=0 to intCountSub
							if clng(pcArrSub(5,k))=pcv_ParentPrd then
								pcArrSub(4,k)=0
							end if
						Next
					end if
					
					query="SELECT discountPerUnit,discountPerWUnit,percentage,baseproductonly FROM discountsPerQuantity WHERE idProduct=" & pcv_ParentPrd & " AND quantityFrom<=" &disTotalQuantity& " AND quantityUntil>=" &disTotalQuantity
					set rsQ1=server.CreateObject("ADODB.RecordSet")
					set rsQ1=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rsQ1=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
					if not rsQ1.eof and err.number<>9 then
						' there are quantity discounts defined for that quantity 
						pDiscountPerUnit=rsQ1("discountPerUnit")
						pDiscountPerWUnit=rsQ1("discountPerWUnit")
						pPercentage=rsQ1("percentage")
						pbaseproductonly=rsQ1("baseproductonly")
		
						set rsQ1=nothing
						
						For k=0 to intCountSub
							if clng(pcArrSub(5,k))=pcv_ParentPrd then
		
								pOrigPrice=pcArrSub(3,k)
								pcArrSub(4,k)=0
								pTotalQuantity=pcArrSub(2,k)
								'reset price for apparel sub-product
								if tmpCustomerType<>"1" then  'customer is a normal user
									if pPercentage="0" then 
										pcArrSub(4,k)=pcArrSub(4,k) + (pDiscountPerUnit * pTotalQuantity)  'running total of discounts
									else
										pcArrSub(4,k)=pcArrSub(4,k) + (((pDiscountPerUnit/100) * pOrigPrice) * pTotalQuantity)
									end if
								else  'customer is a wholesale customer
									if pPercentage="0" then 
										pcArrSub(4,k)=pcArrSub(4,k) + (pDiscountPerWUnit * pTotalQuantity)
									else
										pcArrSub(4,k)=pcArrSub(4,k) + (((pDiscountPerWUnit/100) * pOrigPrice)* pTotalQuantity)
									end if
								end if
							end if
						Next
					end if
					set rsQ1=nothing
				end if
	next
	
	for f=0 to intCountSub
		query="UPDATE ProductsOrdered SET QDiscounts=" & pcArrSub(4,f) & " WHERE idproduct=" & pcArrSub(0,f) & ";"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
	next
	
	END IF
	set rsQ=nothing
'APP-E%>