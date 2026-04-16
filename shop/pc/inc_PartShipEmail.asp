<%if IsNull(pcv_PackageID) or pcv_PackageID="" or (not isNumeric(pcv_PackageID)) then
	pcv_PackageID=0
end if

IF pcv_PackageID<>0 THEN

	query="SELECT pcPackageInfo_UPSServiceCode, pcPackageInfo_UPSPackageType,  pcPackageInfo_ShipMethod, pcPackageInfo_ShippedDate, pcPackageInfo_TrackingNumber, pcPackageInfo_Comments, pcPackageInfo_MethodFlag FROM pcPackageInfo WHERE pcPackageInfo_ID=" & pcv_PackageID & ";"
	set rsQ1=connTemp.execute(query)
	
	if not rsQ1.eof then
		pcv_PK_UPSServiceCode=rsQ1("pcPackageInfo_UPSServiceCode")
		pcv_PK_UPSPackageType=rsQ1("pcPackageInfo_UPSPackageType")
		pcv_PK_ShipMethod=rsQ1("pcPackageInfo_ShipMethod")
		pcv_PK_ShippedDate=rsQ1("pcPackageInfo_ShippedDate")
		pcv_PK_TrackingNumber=rsQ1("pcPackageInfo_TrackingNumber")
		pcv_PK_Comments=rsQ1("pcPackageInfo_Comments")
		pcv_PK_MethodFlag=rsQ1("pcPackageInfo_MethodFlag")
		if pcv_PK_MethodFlag="2" AND pcv_PK_ShipMethod="" then
			select case pcv_PK_UPSServiceCode
				case "01"
					pcv_PK_ShipMethod="UPS Next Day Air"
				case "02"
					pcv_PK_ShipMethod="UPS 2nd Day Air"
				case "03"
					pcv_PK_ShipMethod="UPS Ground"
				case "07"
					pcv_PK_ShipMethod="UPS Worldwide Express"
				case "08"
					pcv_PK_ShipMethod="UPS Worldwide Expedited"
				case "11"
					pcv_PK_ShipMethod="UPS Standard To Canada"
				case "12"
					pcv_PK_ShipMethod="UPS 3 Day Select"
				case "13"
					pcv_PK_ShipMethod="UPS Next Day Air Saver"
				case "14"
					pcv_PK_ShipMethod="UPS Next Day Air"
				case "54"
					pcv_PK_ShipMethod="UPS Worldwide Express Plus"
				case "59"
					pcv_PK_ShipMethod="UPS 2nd Day Air A.M."
				case "65"
					pcv_PK_ShipMethod="UPS Express Saver"
			end select
		end if	
			
		set rsQ1=nothing
		
		query="SELECT idorder FROM ProductsOrdered WHERE pcPackageInfo_ID=" & pcv_PackageID & ";"
		set rsQ1=connTemp.execute(query)
		if not rsQ1.eof then
			qry_ID=rsQ1("idorder")
		end if
		set rsQ1=nothing
		
		'Get Additional Comments
		query="SELECT pcACom_Comments FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=2 AND pcPackageInfo_ID=" & pcv_PackageID
		set rsQ1=connTemp.execute(query)

		pcv_AdmComments=""
		if not rsQ1.eof then
			pcv_AdmComments=rsQ1("pcACom_Comments")
		end if
		set rsQ1=nothing
		'End of Get Additional Comments
		
		'Start Create Product List
		pcv_PrdList=""
		query="SELECT products.idproduct,products.sku, products.description, ProductsOrdered.pcSC_ID, quantity, unitPrice, xfdetails"
		'CONFIGURATOR ADDON-S
		if scBTO=1 then
			query=query&" ,idconfigSession"
		end if
		'CONFIGURATOR ADDON-E
		query=query&", ProductsOrdered.QDiscounts, ProductsOrdered.ItemsDiscounts, ProductsOrdered.pcPrdOrd_SelectedOptions, ProductsOrdered.pcPrdOrd_OptionsPriceArray, ProductsOrdered.pcPrdOrd_OptionsArray, ProductsOrdered.pcPO_GWOpt, ProductsOrdered.pcPO_GWNote, ProductsOrdered.pcPO_GWPrice, ProductsOrdered.pcPrdOrd_BundledDisc FROM products, ProductsOrdered WHERE ProductsOrdered.idproduct=products.idproduct AND ProductsOrdered.idorder=" & qry_ID & " AND ProductsOrdered.pcPackageInfo_ID=" & pcv_PackageID & " AND ProductsOrdered.pcPrdOrd_Shipped=1;"
		
		set rsOrderDetails=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsOrderDetails=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rsOrderDetails.EOF then
			pcv_PrdList=pcv_PrdList & ship_dictLanguage.Item(Session("language")&"_partship_msg_2") & "<br>" & vbcrlf & "<br>" & vbcrlf
			pcv_PrdList=pcv_PrdList & FixedField(10, "L", dictLanguage.Item(Session("language")&"_adminMail_16"))
			pcv_PrdList=pcv_PrdList & FixedField(40, "L", dictLanguage.Item(Session("language")&"_adminMail_18"))
			pcv_PrdList=pcv_PrdList & FixedField(15, "R", "")
			pcv_PrdList=pcv_PrdList & FixedField(15, "R", "")
			pcv_PrdList=pcv_PrdList & "<br>" & vbcrlf
			pcv_PrdList=pcv_PrdList & FixedField(80, "R", "====================================================================================================")
			pcv_PrdList=pcv_PrdList & "<br>" & vbcrlf
				
			Do While Not rsOrderDetails.EOF
				pidProduct=rsOrderDetails("idproduct")
				psku=rsOrderDetails("sku")
				pdescription=rsOrderDetails("description")
				pdescription=ClearHTMLTags2(pdescription,0)
				pcSCID=rsOrderDetails("pcSC_ID")
				if IsNull(pcSCID) OR len(pcSCID)=0 then
					pcSCID=0
				end if
				pqty=rsOrderDetails("quantity")
				pPrice=rsOrderDetails("unitPrice")
				xfdetails=replace(rsOrderDetails("xfdetails"),"&lt;BR&gt;","<br>")
				'xfdetails=replace(xfdetails,"<BR>",vbcrlf)
				if scBTO=1 then
					pIdConfigSession=rsOrderDetails("idconfigSession")
				end if
				QDiscounts=rsOrderDetails("QDiscounts")
				ItemsDiscounts=rsOrderDetails("ItemsDiscounts")	
				
				'// Product Options Arrays
				pcv_strSelectedOptions = rsOrderDetails("pcPrdOrd_SelectedOptions") ' Column 11
				pcv_strOptionsPriceArray = rsOrderDetails("pcPrdOrd_OptionsPriceArray") ' Column 25
				pcv_strOptionsArray = rsOrderDetails("pcPrdOrd_OptionsArray") ' Column 4
				
				'GGG Add-on start	
				pGWOpt=rsOrderDetails("pcPO_GWOpt")
				if pGWOpt<>"" then
				else
				pGWOpt="0"
				end if 
				pGWText=rsOrderDetails("pcPO_GWNote")
				pGWPrice=rsOrderDetails("pcPO_GWPrice")
				if pGWPrice<>"" then
				else
				pGWPrice="0"
				end if
				'GGG Add-on end
				pcPrdOrd_BundledDisc=rsOrderDetails("pcPrdOrd_BundledDisc")
			
				pExtendedPrice=pPrice*pqty
				pcv_PrdList=pcv_PrdList & FixedField(10, "L", pqty)
				dispStr = replace(pdescription & " (" & psku & ")","&quot;", chr(34))
				tStr = dispStr
				wrapPos=40
				if len(dispStr) > 40 then
					tStr = WrapString(40, dispStr)
				end if
				pcv_PrdList=pcv_PrdList & FixedField(40, "L", tStr)
				pcv_PrdList=pcv_PrdList & "<br>" & vbcrlf
				dispStrLen = len(dispStr)-wrapPos
				do while dispStrLen > 40
					dispStr = right(dispStr,dispStrLen)
					tStr = WrapString(40, dispStr)
					pcv_PrdList=pcv_PrdList & FixedField(10, "L", "")
					pcv_PrdList=pcv_PrdList  & FixedField(40, "L", tStr)
					pcv_PrdList=pcv_PrdList  & "<br>" & vbcrlf					
					dispStrLen = dispStrLen-wrapPos	
				loop 
				if dispStrLen > 0 then
					dispStr = right(dispStr,dispStrLen)
					pcv_PrdList=pcv_PrdList  & FixedField(10, "L", "")
					pcv_PrdList=pcv_PrdList  & FixedField(40, "L", dispStr)
					pcv_PrdList=pcv_PrdList  & "<br>" & vbcrlf
				end if
				
				'CONFIGURATOR ADDON-S
				TotalUnit=0
				if scBTO=1 then
					if pIdConfigSession<>"0" then
						query="SELECT * FROM configSessions WHERE idconfigSession=" & pIdConfigSession
						set rsConfigObj=conntemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rsConfigObj=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
						pcv_PrdList=pcv_PrdList & FixedField(10, "L","") & FixedField(40, "L", dictLanguage.Item(Session("language")&"_adminMail_34A"))
						pcv_PrdList=pcv_PrdList & FixedField(15, "R", " ")
						pcv_PrdList=pcv_PrdList & FixedField(15, "R", " ") & "<br>" & vbcrlf
						stringProducts=rsConfigObj("stringProducts")
						stringValues=rsConfigObj("stringValues")
						stringCategories=rsConfigObj("stringCategories")
						stringQuantity=rsConfigObj("stringQuantity")
						stringPrice=rsConfigObj("stringPrice")
						ArrProduct=Split(stringProducts, ",")
						ArrValue=Split(stringValues, ",")
						ArrCategory=Split(stringCategories, ",")
						ArrQuantity=Split(stringQuantity, ",")
						ArrPrice=Split(stringPrice, ",")
						for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
							
							If statusAPP="1" Then
			
								query="SELECT products.pcProd_ParentPrd FROM products WHERE products.idProduct="&ArrProduct(i)&";" 
								set rsConfigObj=conntemp.execute(query)
								If Not rsConfigObj.Eof Then
									pcv_intIdProduct=rsConfigObj("pcProd_ParentPrd")
								End If
								Set rsConfigObj = Nothing
			
								if pcv_intIdProduct>"0" then
								else
									pcv_intIdProduct=ArrProduct(i)
								end if
			
							Else
			
								pcv_intIdProduct = ArrProduct(i)
							
							End If
							
							query="SELECT displayQF FROM configSpec_Products WHERE configProduct="& pcv_intIdProduct &" AND specProduct=" & pidProduct & " AND configProductCategory=" & ArrCategory(i) & ";"
							set rsQ=server.CreateObject("ADODB.RecordSet") 
							set rsQ=conntemp.execute(query)
							if not rsQ.eof then	
								btDisplayQF=rsQ("displayQF")
							end if
							set rsQ=nothing
														
							query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & pcv_intIdProduct & ";"
							set rsQ=connTemp.execute(query)
							tmpMinQty=1
							if not rsQ.eof then
								tmpMinQty=rsQ("pcprod_minimumqty")
								if IsNull(tmpMinQty) or tmpMinQty="" then
									tmpMinQty=1
								else
									if tmpMinQty="0" then
										tmpMinQty=1
									end if
								end if
							end if
							set rsQ=nothing
							tmpDefault=0
							query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pidProduct & " AND configProduct=" & pcv_intIdProduct & " AND cdefault<>0" & " AND configProductCategory=" & ArrCategory(i) & ";"
							set rsQ=connTemp.execute(query)
							if not rsQ.eof then
								tmpDefault=rsQ("cdefault")
								if IsNull(tmpDefault) or tmpDefault="" then
									tmpDefault=0
								else
									if tmpDefault<>"0" then
										tmpDefault=1
									end if
								end if
							end if
							set rsQ=nothing
							
							query="SELECT products.sku, categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
							set rsConfigObj=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsConfigObj=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							pcv_strBtoItemSku = rsConfigObj("sku")
							pcv_strBtoItemSku=ClearHTMLTags2(pcv_strBtoItemSku,0)
							pcv_strBtoItemName = rsConfigObj("description")
							pcv_strBtoItemName=ClearHTMLTags2(pcv_strBtoItemName,0)
							pcv_strBtoItemCat=rsConfigObj("categoryDesc")
							pcv_strBtoItemCat=ClearHTMLTags2(pcv_strBtoItemCat,0)
							pcv_PrdList=pcv_PrdList & FixedField(10, "L", "")
							dispStr = ""
							dispStr = pcv_strBtoItemCat &": "& pcv_strBtoItemName
							dispStr = dispStr & " - SKU: " & pcv_strBtoItemSku
							if btDisplayQF=True then
								if clng(ArrQuantity(i))>1 then
									dispStr = dispStr & " - QTY: " & ArrQuantity(i)
								end if
							end if
							dispStr = replace(dispStr,"&quot;", chr(34))
							tStr = dispStr
							wrapPos=40
							if len(dispStr) > 40 then
								tStr = WrapString(40, dispStr)
							end if
							pcv_PrdList=pcv_PrdList & FixedField(40, "L", tStr)
			
							if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
								if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
									if tmpDefault=1 then
										UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
									else
										UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
									end if
								else
									UPrice=0
								end if
								TotalUnit=TotalUnit+cdbl((ArrValue(i)+UPrice)*pQty)
								pcv_PrdList=pcv_PrdList & FixedField(30, "R", "")
							else
								if tmpDefault=1 then
									pcv_PrdList=pcv_PrdList & FixedField(30, "R", dictLanguage.Item(Session("language")&"_defaultnotice_1"))
								end if
							end if
							pcv_PrdList=pcv_PrdList & "<br>" & vbcrlf
							dispStrLen = len(dispStr)-wrapPos
							do while dispStrLen > 40
								dispStr = right(dispStr,dispStrLen)
								tStr = WrapString(40, dispStr)
								pcv_PrdList=pcv_PrdList & FixedField(10, "L", "")
								pcv_PrdList=pcv_PrdList  & FixedField(40, "L", tStr)
								pcv_PrdList=pcv_PrdList  & "<br>" & vbcrlf					
								dispStrLen = dispStrLen-wrapPos	
							loop 
							if dispStrLen > 0 then
								dispStr = right(dispStr,dispStrLen)
								pcv_PrdList=pcv_PrdList  & FixedField(10, "L", "")
								pcv_PrdList=pcv_PrdList  & FixedField(40, "L", dispStr)
								pcv_PrdList=pcv_PrdList  & "<br>" & vbcrlf
							end if
							set rsConfigObj=nothing
						next
					end if
				end if
				'CONFIGURATOR ADDON-E
				
				
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START: Add first 40 characters of options on a separate line
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
				if isNull(pcv_strSelectedOptions) or pcv_strSelectedOptions="NULL" then
					pcv_strSelectedOptions = ""
				end if
				
				If len(pcv_strSelectedOptions)>0 Then
						'// Add the header "OPTIONS"		
						pcv_PrdList=pcv_PrdList & FixedField(10, "L","") & FixedField(40, "L","OPTIONS") & "<br>" & vbcrlf
						
						'#####################
						' START LOOP
						'#####################				
						'// Generate Our Local Arrays from our Stored Arrays  
						
						' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers	
						pcArray_strSelectedOptions = ""					
						pcArray_strSelectedOptions = Split(pcv_strSelectedOptions,chr(124))
						
						' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
						pcArray_strOptionsPrice = ""
						pcArray_strOptionsPrice = Split(pcv_strOptionsPriceArray,chr(124))
						
						' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
						pcArray_strOptions = ""
						pcArray_strOptions = Split(pcv_strOptionsArray,chr(124))
						
						' Get Our Loop Size
						pcv_intOptionLoopSize = 0
						pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)
			
						' Display Our Options
						For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize
							dispStr = ""
										
							'//There isnt a header after the first one so we indent
							pcv_PrdList=pcv_PrdList & FixedField(10, "L", " ")
						
							dispStr = pcArray_strOptions(pcv_intOptionLoopCounter)
							dispStr = replace(dispStr,"&quot;", chr(34))
							tStr = dispStr
							wrapPos=40
							if len(dispStr) > 40 then
								tStr = WrapString(40, dispStr)
							end if
							pcv_PrdList=pcv_PrdList & FixedField(40, "L", tStr)
			
							tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
																
							if tempPrice="" or tempPrice=0 then
								pcv_PrdList=pcv_PrdList & FixedField(30, "R", " ")
								pcv_PrdList=pcv_PrdList & "<br>" & vbcrlf
							else 
								pcv_PrdList=pcv_PrdList & FixedField(30, "R", "")
								pcv_PrdList=pcv_PrdList & "<br>" & vbcrlf
							end if
							dispStrLen = len(dispStr)-wrapPos
							do while dispStrLen > 40
								dispStr = right(dispStr,dispStrLen)
								tStr = WrapString(40, dispStr)
								pcv_PrdList=pcv_PrdList & FixedField(10, "L", "")
								pcv_PrdList=pcv_PrdList  & FixedField(40, "L", tStr)
								pcv_PrdList=pcv_PrdList  & "<br>" & vbcrlf					
								dispStrLen = dispStrLen-wrapPos	
							loop 
							if dispStrLen > 0 then
								dispStr = right(dispStr,dispStrLen)
								pcv_PrdList=pcv_PrdList  & FixedField(10, "L", "")
								pcv_PrdList=pcv_PrdList  & FixedField(40, "L", dispStr)
								pcv_PrdList=pcv_PrdList  & "<br>" & vbcrlf
							end if
							
						Next
						'#####################
						' END LOOP
						'#####################	
						
						pcv_PrdList=pcv_PrdList & "<br>" & vbcrlf
						
				End If
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END: Add first 40 characters of options on a separate line
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
				
				If len(xfdetails)>3 then
					pcv_PrdList=pcv_PrdList & "<br>" & vbcrlf
					xfarray=split(xfdetails,"|")
					for q=lbound(xfarray) to ubound(xfarray)
						pcv_PrdList=pcv_PrdList & FixedField(10, "L", "")
						dispStr = replace(xfarray(q),"&quot;", chr(34))
						tStr = dispStr
						wrapPos=40
						if len(dispStr) > 40 then
							tStr = WrapString(40, dispStr)
						end if
						if Instr(tStr,"<BR>") then
							tStr=replace(tStr,"<BR>","<br>" & vbcrlf & FixedField(10, "L", ""))
							pcv_PrdList=pcv_PrdList & tStr & "<br>" & vbcrlf
						else
							pcv_PrdList=pcv_PrdList & FixedField(40, "L", tStr) & "<br>" & vbcrlf
						end if
						dispStrLen = len(dispStr)-wrapPos
						if inStr(dispStr,"<BR>")>0 then
						  if dispStrLen > 0 then
							  dispStr = right(dispStr,dispStrLen)
							end if
							dispStr = FixedField(10, "L", "") & replace(dispStr,"<BR>","<br>" & vbcrlf & FixedField(10, "L", ""))
							pcv_PrdList=pcv_PrdList & dispStr & "<br>" & vbcrlf
						else
							do while dispStrLen > 40
								dispStr = right(dispStr,dispStrLen)
								response.write dispStr & "<br>"
								tStr = WrapString(40, dispStr)
								response.write tStr & "<br>"
								if Instr(tStr,"<BR>") then
									tStr=replace(tStr,"<BR>","<br>" & vbcrlf & FixedField(10, "L", ""))
								end if
								pcv_PrdList=pcv_PrdList & FixedField(10, "L", "")
								pcv_PrdList=pcv_PrdList  & FixedField(40, "L", tStr)
								pcv_PrdList=pcv_PrdList  & "<br>" & vbcrlf					
								dispStrLen = dispStrLen-wrapPos	
							loop 
							if dispStrLen > 0 then
								dispStr = right(dispStr,dispStrLen)
								if Instr(dispStr,"<BR>") then
									dispStr=replace(dispStr,"<BR>","<br>" & vbcrlf & FixedField(10, "L", ""))
								end if
								pcv_PrdList=pcv_PrdList  & FixedField(10, "L", "")
								pcv_PrdList=pcv_PrdList  & FixedField(40, "L", dispStr)
								pcv_PrdList=pcv_PrdList  & "<br>" & vbcrlf
							end if
						end if
					next
				End If
				
				pPrice1=pPrice
				pExtendedPrice1=pExtendedPrice
				
				if TotalUnit>0 then
					pExtendedPrice1=pExtendedPrice1-TotalUnit
					pPrice1=Round(pExtendedPrice1/pqty,2)
				end if	
			
				tmpText1=""
				if money(pPrice1)=money(pExtendedPrice1) then
					tmpText1=tmpText1 & FixedField(15, "R","")
				else
					tmpText1=tmpText1 & FixedField(15, "R", scCurSign & money(pPrice1))
				end if
				tmpText1=tmpText1 & FixedField(15, "R", scCurSign & money(pExtendedPrice1)) & "<br>" & vbcrlf	
				pcv_PrdList=pcv_PrdList & "<br>" & vbcrlf
			
				'CONFIGURATOR ADDON-S
				Charges=0
				if scBTO=1 then
					if pIdConfigSession<>"0" then
					
					'BTO Additional Charges
					'Add customizations if there are any
					if pIdConfigSession<>"0" then
						query="SELECT * FROM configSessions WHERE idconfigSession=" & pIdConfigSession
						set rsConfigObj=conntemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rsConfigObj=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
						stringCProducts=rsConfigObj("stringCProducts")
						stringCValues=rsConfigObj("stringCValues")
						stringCCategories=rsConfigObj("stringCCategories")
						ArrCProduct=Split(stringCProducts, ",")
						ArrCValue=Split(stringCValues, ",")
						ArrCCategory=Split(stringCCategories, ",")
						if ArrCProduct(0)<>"na" then
						for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
							Charges=Charges+Cdbl(ArrCValue(i))
						next
						
						pcv_PrdList=pcv_PrdList & FixedField(10, "L","") & FixedField(40, "L", dictLanguage.Item(Session("language")&"_adminMail_34"))
						pcv_PrdList=pcv_PrdList & FixedField(15, "R", " ")
						pcv_PrdList=pcv_PrdList & FixedField(15, "R", " ") & "<br>" & vbcrlf
						
						for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
							query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
							set rsConfigObj=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsConfigObj=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							pcv_PrdList=pcv_PrdList & FixedField(10, "L", "")
							dispStr = rsConfigObj("categoryDesc")&": "&rsConfigObj("description")
							dispStr = replace(dispStr,"&quot;", chr(34))
							tStr = dispStr
							wrapPos=40
							if len(dispStr) > 40 then
								tStr = WrapString(40, dispStr)
							end if
							pcv_PrdList=pcv_PrdList & FixedField(40, "L", tStr)
			
							if ArrCValue(i)<>0 then
								pcv_PrdList=pcv_PrdList & FixedField(30, "R", "")
							end if
							pcv_PrdList=pcv_PrdList & "<br>" & vbcrlf
			
							dispStrLen = len(dispStr)-wrapPos
							do while dispStrLen > 40
								dispStr = right(dispStr,dispStrLen)
								tStr = WrapString(40, dispStr)
								pcv_PrdList=pcv_PrdList & FixedField(10, "L", "")
								pcv_PrdList=pcv_PrdList  & FixedField(40, "L", tStr)
								pcv_PrdList=pcv_PrdList  & "<br>" & vbcrlf					
								dispStrLen = dispStrLen-wrapPos	
							loop 
							if dispStrLen > 0 then
								dispStr = right(dispStr,dispStrLen)
								pcv_PrdList=pcv_PrdList  & FixedField(10, "L", "")
								pcv_PrdList=pcv_PrdList  & FixedField(40, "L", dispStr)
								pcv_PrdList=pcv_PrdList  & "<br>" & vbcrlf
							end if
							set rsConfigObj=nothing
						next
						end if
					end if
						'BTO Additional Charges
						iSubTotal=iSubtotal + (pPrice*pqty)-cdbl(ItemsDiscounts)+cdbl(Charges)-cdbl(pcPrdOrd_BundledDisc)
					else
						iSubTotal=iSubtotal + (pPrice*pqty)-cdbl(pcPrdOrd_BundledDisc)
				end if
				else
					iSubTotal=iSubtotal + (pPrice*pqty)-cdbl(pcPrdOrd_BundledDisc)
				end if
				'======================================
				iSubTotal=iSubtotal-cdbl(QDiscounts)
			
				cdblCmprTmp1=(pPrice*pqty)
				cdblCmprTmp2=(pPrice*pqty)-cdbl(QDiscounts)-cdbl(ItemsDiscounts)+cdbl(Charges)
			
				if cdblCmprTmp2<>cdblCmprTmp1 then
					pcv_PrdList=pcv_PrdList & FixedField(10, "L","") & FixedField(40, "L", dictLanguage.Item(Session("language")&"_adminMail_33"))
					pcv_PrdList=pcv_PrdList & FixedField(15, "R", " ")
					pcv_PrdList=pcv_PrdList & FixedField(15, "R", "") & "<br>" & vbcrlf
				end if
				
						
				pcv_PrdList=pcv_PrdList & "<br>" & vbcrlf & "<br>" & vbcrlf
					
				rsOrderDetails.MoveNext
			loop
		end if
		set rsOrderDetails=nothing
		'End of Create Product List
			
		'Create Link
		strPathInfo=""
		strPath=Request.ServerVariables("PATH_INFO")
		iCnt=0
		do while iCnt<2
			if mid(strPath,len(strPath),1)="/" then
				iCnt=iCnt+1
			end if
			if iCnt<2 then
				strPath=mid(strPath,1,len(strPath)-1)
			end if
		loop
	
		strPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & strPath
					
		if Right(strPathInfo,1)="/" then
		else
			strPathInfo=strPathInfo & "/"
		end if
			
		strPathInfo=strPathInfo & scAdminFolderName & "/OrdDetails.asp?id=" & qry_ID
		'End of Create Link
		
		if pcv_LastShip="1" then ' This is the last (or only) shipment
			pcv_DropShipperSbj=ship_dictLanguage.Item(Session("language")&"_partship_sbj_9")
			pcv_DropShipperSbj=replace(pcv_DropShipperSbj,"<ORDER_ID>",(scpre + int(qry_ID)))	
			'pcv_DropShipperMsg=pcv_DropShipperSbj & "<br>" & vbcrlf & "<br>" & vbcrlf
			pcv_DropShipperMsg=""
			'pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_8") & "<br>" & vbcrlf & "<br>" & vbcrlf
			'pcv_DropShipperMsg1=ship_dictLanguage.Item(Session("language")&"_partship_msg_8a") & "<br>" & vbcrlf
		else
			pcv_DropShipperSbj=ship_dictLanguage.Item(Session("language")&"_partship_sbj_1")
			pcv_DropShipperSbj=replace(pcv_DropShipperSbj,"<ORDER_ID>",(scpre + int(qry_ID)))	
			pcv_DropShipperMsg=pcv_DropShipperSbj & "<br>" & vbcrlf & "<br>" & vbcrlf
			pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_1") & "<br>" & vbcrlf & "<br>" & vbcrlf
			pcv_DropShipperMsg1=ship_dictLanguage.Item(Session("language")&"_partship_msg_1a") & "<br>" & vbcrlf & "<br>" & vbcrlf
		end if
		if pcv_AdmComments<>"" then
			pcv_DropShipperMsg=pcv_DropShipperMsg & vbcrlf & replace(pcv_AdmComments,"''","'") & "<br>" & vbcrlf & "<br>" & vbcrlf
		end if
		pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_3") & "<br>" & vbcrlf
		pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_3") & "<br>" & vbcrlf
		if pcv_PK_ShipMethod<>"" then
			pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_4") & pcv_PK_ShipMethod & "<br>" & vbcrlf
			pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_4") & pcv_PK_ShipMethod & "<br>" & vbcrlf
			pcv_DropShipperMsg=pcv_DropShipperMsg & "Courier: DPD Local<br>" & vbcrlf
			pcv_DropShipperMsg1=pcv_DropShipperMsg1 & "Courier: DPD Local<br>" & vbcrlf
		end if
		if pcv_PK_TrackingNumber<>"" then	
			pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_5") & "<a href=""https://www.dpdlocal.co.uk/apps/tracking/?reference=" & pcv_PK_TrackingNumber & """>" & pcv_PK_TrackingNumber & "</a><br>" & vbcrlf
			pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_5") & "<a href=""https://www.dpdlocal.co.uk/apps/tracking/?reference=" & pcv_PK_TrackingNumber & """>" & pcv_PK_TrackingNumber & "</a><br>" & vbcrlf
		end if
		if not IsNull(pcv_PK_ShippedDate) then
			pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_6") & ShowDateFrmt(pcv_PK_ShippedDate) & "<br>" & vbcrlf
			pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_6") & ShowDateFrmt(pcv_PK_ShippedDate) & "<br>" & vbcrlf
		end if
		pcv_DropShipperMsg=pcv_DropShipperMsg & "<br>" & vbcrlf & pcv_PrdList & "<br>" & vbcrlf
		pcv_DropShipperMsg1=pcv_DropShipperMsg1 & "<br>" & vbcrlf & pcv_PrdList & "<br>" & vbcrlf
		if pcv_PK_TrackingNumber<>"" then		
			'//  Start: Tracking Link
			pcv_strTempLink=""
			if instr(ucase(pcv_PK_ShipMethod),"UPS:") then
				pcv_strTempLink = scStoreURL & "/" & scPcFolder & "/pc/custUPSTracking.asp?itracknumber=" & pcv_PK_TrackingNumber & "<br>" & vbcrlf & "<br>" & vbcrlf
				pcv_strTempUPSLink=replace(pcv_strTempLink,"//","/")
				pcv_strTempLink=replace(pcv_strTempLink,"http:/","http://")
				pcv_strTempLink=replace(pcv_strTempLink,"https:/","https://")
			elseif instr(ucase(pcv_PK_ShipMethod),"FEDEX:") then
				pcv_strTempLink = "http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=" & pcv_PK_TrackingNumber
			end if	
			if pcv_strTempLink<>"" then
				pcv_DropShipperMsg=pcv_DropShipperMsg & ship_dictLanguage.Item(Session("language")&"_partship_msg_9") & pcv_strTempLink & "<br>" & vbcrlf
				pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_9") & pcv_strTempLink & "<br>" & vbcrlf
			end if
			'//  End: Tracking Link	
		end if
		if pcv_PK_Comments<>"" then
			pcv_DropShipperMsg1=pcv_DropShipperMsg1 & ship_dictLanguage.Item(Session("language")&"_partship_msg_7") & "<br>" & vbcrlf & pcv_PK_Comments & "<br>" & vbcrlf
		end if
		
		if (pcv_PrdList="") and (pcv_PK_Comments<>"") then
			pcv_DropShipperMsg1=ship_dictLanguage.Item(Session("language")&"_partship_msg_7") & "<br>" & vbcrlf & pcv_PK_Comments & "<br>" & vbcrlf
		end if
		
		pcv_DropShipperMsg1=pcv_DropShipperMsg1 & "<br>" & vbcrlf & strPathInfo & "<br>" & vbcrlf
			
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"''",chr(39))
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"//","/")
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"http:/","http://")
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"https:/","https://")
		pcv_DropShipperMsg=replace(pcv_DropShipperMsg,"<ORDER_ID>",(scpre + int(qry_ID)))
		
		pcv_DropShipperMsg1=replace(pcv_DropShipperMsg1,"''",chr(39))
		pcv_DropShipperMsg1=replace(pcv_DropShipperMsg1,"//","/")
		pcv_DropShipperMsg1=replace(pcv_DropShipperMsg1,"http:/","http://")
		pcv_DropShipperMsg1=replace(pcv_DropShipperMsg1,"https:/","https://")
		pcv_DropShipperMsg1=replace(pcv_DropShipperMsg1,"<ORDER_ID>",(scpre + int(qry_ID)))
		
		if pcv_ResendShip="1" then
		
			query="SELECT pcACom_Comments FROM pcAdminComments WHERE idorder=" & qry_ID & " AND pcACom_ComType=3;"
			set rsQ1=connTemp.execute(query)
			
			pcv_AdmComments1=""
			if not rsQ1.eof then
				pcv_AdmComments1=rsQ1("pcACom_Comments")
			end if
			set rsQ1=nothing
			
			if pcv_AdmComments1<>"" then
				pcv_DropShipperMsg=replace(pcv_AdmComments1,"''","'") & "<br>" & vbcrlf & "------------------------------------------------" & "<br>" & vbcrlf  & "<br>" & vbcrlf & pcv_DropShipperMsg 
			end if
		end if
		
		session("News_MsgType")="1"
			
		if (pcv_SendCust="1") and (pcv_PrdList<>"") then
			query="SELECT Customers.email, Orders.pcOrd_ShippingEmail FROM Customers INNER JOIN Orders ON Customers.idcustomer = Orders.idCustomer WHERE Orders.idOrder=" & qry_ID & ";"
			set rsQ1=connTemp.execute(query)
			pEmail=rsQ1("email")
			pShippingEmail=rsQ1("pcOrd_ShippingEmail")
			
			set rsQ1=nothing
			pcv_DropShipperMsg=pcv_DropShipperMsg & "<br>" & vbcrlf & dictLanguage.Item(Session("language")&"_sendMail_36") & scCompanyName & "." & "<br>" & vbcrlf & "<br>" & vbcrlf
			
			'************************************************************************************************************************
			' START: Shipper Information
			' The e-mail is sent so that the shipment appears to be coming from your store
			' If you want to include the drop-shipper's information instead, uncomment the following 4 commented lines of code
			'************************************************************************************************************************

			'if pcv_UseDropShipperInfo="1" then
			'	call sendmail (pcv_DS_Name, pcv_DS_Email, pEmail, pcv_DropShipperSbj, replace(pcv_DropShipperMsg, "&quot;", chr(34)))
			'else
				pcv_DropShipperMsg=pcf_HtmlEmailWrapper(pcv_DropShipperMsg, pcv_HTMLEmailFontFamily)
				call sendmail (scCompanyName, scEmail, pEmail, pcv_DropShipperSbj, pcv_DropShipperMsg)
				call pcs_hookOrderPartShippedEmailSent(pEmail)
				
				'//Send email to shipping email if it is different and exist
				if trim(pShippingEmail)<>"" AND trim(pShippingEmail)<>trim(pEmail) then
					call sendmail (scCompanyName, scEmail, pShippingEmail, pcv_DropShipperSbj, pcv_DropShipperMsg)
					call pcs_hookOrderPartShippedEmailSent(pShippingEmail)
				end if
			'end if
			
			'************************************************************************************************************************
			' END: Shipper Information
			'************************************************************************************************************************
			
		end if
		if pcv_SendAdmin="1" then
			if (pcv_PrdList="") and (pcv_PK_Comments<>"") then
				pcv_DropShipperSbj=ship_dictLanguage.Item(Session("language")&"_partship_sbj_1a")
				pcv_DropShipperSbj=replace(pcv_DropShipperSbj,"<ORDER_ID>",(scpre + int(qry_ID)))
			end if
			pcv_DropShipperMsg1=pcf_HtmlEmailWrapper(pcv_DropShipperMsg1, pcv_HTMLEmailFontFamily)
			if pcv_UseDropShipperInfo="1" then
				call sendmail (pcv_DS_Name, pcv_DS_Email, scFrmEmail, pcv_DropShipperSbj, pcv_DropShipperMsg1)
			else
				call sendmail (scCompanyName, scEmail, scFrmEmail, pcv_DropShipperSbj, pcv_DropShipperMsg1)
			end if
		end if
		
	End if 'Have Package Info
	set rsQ1=nothing
	
END IF%>
