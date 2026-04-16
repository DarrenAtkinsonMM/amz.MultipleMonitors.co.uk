<%'APP-S
	call opendb()
	pcv_ParentPrd=0
	tmpSaveParent="||"
	for f=1 to ppcCartIndex
		if pcCartArray(f,10)=0 then
			if pcCartArray(f,32)<>"" then

				pcv_ParentPrd=pcf_ProductIdFromArray(pcCartArray, f)
                
				if instr(tmpSaveParent,"||" & pcv_ParentPrd & "||")=0 then
					tmpSaveParent=tmpSaveParent & pcv_ParentPrd & "||"
					tmpQuantity=0
					For k=1 to ppcCartIndex
						if pcCartArray(k,10)=0 then
							if pcCartArray(k,32)<>"" then
								tmp1=pcf_ProductIdFromArray(pcCartArray, k)
								if tmp1=pcv_ParentPrd then
									tmpQuantity=tmpQuantity+cLng(pcCartArray(k,2))
								end if
							end if
						end if
					Next
	
					disTotalQuantity=tmpQuantity
					
					query="SELECT discountPerUnit,discountPerWUnit,percentage,baseproductonly FROM discountsPerQuantity WHERE idProduct=" & pcv_ParentPrd & ";"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					if not rstemp.eof then
						For k=1 to ppcCartIndex
							if pcCartArray(k,10)=0 then
								if pcCartArray(k,32)<>"" then

                                    tmp1=pcf_ProductIdFromArray(pcCartArray, k)
									if tmp1=pcv_ParentPrd then
										pcCartArray(k,15)=0
										pcCartArray(k,3)=pcCartArray(k,17)
									end if
                                    
								end if
							end if
						Next
					end if
					
					query="SELECT discountPerUnit,discountPerWUnit,percentage,baseproductonly FROM discountsPerQuantity WHERE idProduct=" & pcv_ParentPrd & " AND quantityFrom<=" &disTotalQuantity& " AND quantityUntil>=" &disTotalQuantity
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
					if not rstemp.eof and err.number<>9 then
						' there are quantity discounts defined for that quantity 
						pDiscountPerUnit=rstemp("discountPerUnit")
						pDiscountPerWUnit=rstemp("discountPerWUnit")
						pPercentage=rstemp("percentage")
						pbaseproductonly=rstemp("baseproductonly")
		
						set rstemp=nothing
						
						For k=1 to ppcCartIndex
						if pcCartArray(k,10)=0 then
						if pcCartArray(k,32)<>"" then

							tmp1 = pcf_ProductIdFromArray(pcCartArray, k)
							if tmp1=pcv_ParentPrd then
								if session("customerType")=1 then
									pcCartArray(k,18)=1
								else
									pcCartArray(k,18)=0
								end if
		
								pOrigPrice=pcCartArray(k,17)
								pcCartArray(k,15)=0
								pTotalQuantity=pcCartArray(k,2)
								'reset price for apparel sub-product
								pcCartArray(k,3)=pOrigPrice
								if session("customerType")<>1 then  'customer is a normal user
									if pPercentage="0" then 
										pcCartArray(k,3)=pcCartArray(k,3) - pDiscountPerUnit  'Price - discount per unit
										pcCartArray(k,15)=pcCartArray(k,15) + (pDiscountPerUnit * pTotalQuantity)  'running total of discounts
									else
										if pbaseproductonly="-1" then
											pcCartArray(k,3)=pcCartArray(k,3) - ((pDiscountPerUnit/100) * pcCartArray(k,17))
										else
											pcCartArray(k,3)=pcCartArray(k,3) - ((pDiscountPerUnit/100) * (pcCartArray(k,17)+pcCartArray(k,5)))
										end if
										if pbaseproductonly="-1" then
											pcCartArray(k,15)=pcCartArray(k,15) + (((pDiscountPerUnit/100) * pOrigPrice) * pTotalQuantity)
										else
											pcCartArray(k,15)=pcCartArray(k,15) + (((pDiscountPerUnit/100) * (pOrigPrice+pcCartArray(k,5))) * pTotalQuantity)
										end if
									end if
								else  'customer is a wholesale customer
									if pPercentage="0" then 
										pcCartArray(k,3)=pcCartArray(k,3) - pDiscountPerWUnit
										pcCartArray(k,15)=pcCartArray(k,15) + (pDiscountPerWUnit * pTotalQuantity)
									else
										if pbaseproductonly="-1" then
											pcCartArray(k,3)=pcCartArray(k,3) - ((pDiscountPerWUnit/100) * pcCartArray(k,17))
										else
											pcCartArray(k,3)=pcCartArray(k,3) - ((pDiscountPerWUnit/100) * (pcCartArray(k,17)+pcCartArray(k,5)))
										end if
										if pbaseproductonly="-1" then
											pcCartArray(k,15)=pcCartArray(k,15) + (((pDiscountPerWUnit/100) * pOrigPrice)* pTotalQuantity)
										else
											pcCartArray(k,15)=pcCartArray(k,15) + (((pDiscountPerWUnit/100) * (pOrigPrice+pcCartArray(k,5))) * pTotalQuantity)
										end if
									end if
								end if
							end if
						end if
						end if
						Next
					end if
					set rstemp=nothing
				end if
			end if
		end if
	next
'APP-E%>