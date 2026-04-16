<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<%
grCode=getUserInput(request("grCode"),0)
geID=getUserInput(request("geID"),0)

gparentPrdId=getUserInput(request("gparentPrdId"),0)
if NOT isNumeric(gparentPrdId) OR gparentPrdId="" then
	gparentPrdId = 0
end if


if grCode="" then
	response.redirect "msg.asp?message=98"
end if

if geID="" then
	response.redirect "msg.asp?message=99"
end if

query="select pcEv_IDEvent,pcEv_Name,pcEv_Date,pcEv_Type,pcEv_IncGcs from pcEvents where pcEv_Code='" & grCode & "' and pcEv_Active=1"
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if rstemp.eof then
	call closedb()
    response.redirect "msg.asp?message=98"
else
	gIDEvent=rstemp("pcEv_IDEvent")
	geName=rstemp("pcEv_Name")
	geDate=rstemp("pcEv_Date")
	if gedate<>"" then
		if scDateFrmt="DD/MM/YY" then
			gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
		else
			gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
		end if
	end if
	gType=rstemp("pcEv_Type")
	if gType<>"" then
	else
		gType=""
	end if
	geincGc=rstemp("pcEv_IncGcs")
	if geincGc<>"" then
	else
		geincGc="0"
	end if
end if
set rstemp=nothing

query="select products.idproduct,products.sku,products.description,products.imageUrl,products.largeImageURL,products.details,products.sdesc,pcEvProducts.pcEP_Price,pcEvProducts.pcEP_IDConfig, pcEvProducts.pcEP_OptionsArray , pcEvProducts.pcEP_xdetails from products,pcEvProducts where pcEvProducts.pcEP_IDEvent=" & gIDEvent & " and products.idproduct=pcEvProducts.pcEP_IDProduct and pcEvProducts.pcEP_ID=" & geID
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rstemp.eof then
	response.redirect "msg.asp?message=99"
else
	pIDProduct=rstemp("idproduct")
	pSku=rstemp("sku")
	pname=rstemp("description")
	pImageURL=rstemp("imageUrl")
	pLgimageURL=rstemp("largeImageURL")

	'// Check if this is an apparel product and get parent's descriptions and images
	if gparentPrdId = 0 then
			pdetails=rstemp("details")
			psdesc=rstemp("sdesc")
	else

			query="SELECT products.imageUrl,products.largeImageURL,products.details,products.sdesc FROM products WHERE idProduct = "& gparentPrdId
			set rsTempParentInfo=Server.CreateObject("ADODB.Recordset")
			set rsTempParentInfo=connTemp.execute(query)

			pdetails=rsTempParentInfo("details")
			psdesc=rsTempParentInfo("sdesc")
			pImageTempURL=rsTempParentInfo("imageUrl")
			pLgimageTempURL=rsTempParentInfo("largeImageURL")
			if trim(pImageURL)="" then
				pImageURL=pImageTempURL
			end if
			if trim(pLgimageURL)="" then
				pLgimageURL=pLgimageTempURL
			end if
			set rsTempParentInfo=nothing

	end if

	pPrice=rstemp("pcEP_Price")
	pIDConfig=rstemp("pcEP_IDConfig")
	if pIDConfig<>"" then
	else
		pIDConfig="0"
	end if
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Product Options
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcv_strSelectedOptions=""
	pcv_strSelectedOptions = rstemp("pcEP_OptionsArray")
	pcv_strSelectedOptions=pcv_strSelectedOptions&""		
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Product Options
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	
	pxdetails=rstemp("pcEP_xdetails")
end if
set rstemp=nothing

%> 
<!--#include file="header_wrapper.asp"-->
<!--#include file="../includes/javascripts/pcWindowsViewPrd.asp"-->

<div id="pcMain">
	<div class="pcMainContent">
		<div id="pcViewProductEP">
			<div class="pcViewProductLeft">
				<div class="pcFormItem">
					<div class="pcFormItemFull"><strong><%=pname%></strong></div>
					<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_viewEP_1")%><%=pSKU%></div>
				</div>

				<div class="pcSpacer"></div>

				<div class="pcFormItem">
					<div class="pcFormItemFull">
					<% if trim(psDesc) <> "" then 
 						response.Write psDesc & " <a href='#details'>" & dictLanguage.Item(Session("language")&"_viewEP_7") & "</a>"
 					else
 						response.Write pDetails
					end if %>
					</div>
				</div>
		
				<div class="pcSpacer"></div>
		
				<div class="pcFormItem">
					<div class="pcFormItemFull">
						<strong><%= dictLanguage.Item(Session("language")&"_viewEP_2")%></strong><%=scCurSign & money(pPrice)%>
					</div>
				</div>

					<%IF pIDConfig<>"0" then%>
					<%query="SELECT * FROM configSessions WHERE idconfigSession=" & pIDConfig
					set rs=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
							
					stringProducts=rs("stringProducts")
					stringValues=rs("stringValues")
					stringCategories=rs("stringCategories")
					ArrProduct=Split(stringProducts, ",")
					ArrValue=Split(stringValues, ",")
					ArrCategory=Split(stringCategories, ",")
					Qstring=rs("stringQuantity")
					ArrQuantity=Split(Qstring,",")
					Pstring=rs("stringPrice")
					ArrPrice=split(Pstring,",")
					stringCProducts=rs("stringCProducts")
					stringCValues=rs("stringCValues")
					stringCCategories=rs("stringCCategories")
					ArrCProduct=Split(stringCProducts, ",")
					ArrCValue=Split(stringCValues, ",")
					ArrCCategory=Split(stringCCategories, ",")
		
					set rs=nothing

					if ArrProduct(0)="na" then
					else%>
						<%tempCat=""%>
						<div class="pcShowBTOconfiguration">
						<div class="pcTableRowFull"><b><%= dictLanguage.Item(Session("language")&"_viewEP_3")%></b></div>
						<%for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
							query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
							set rsObj=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsObj=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i) & " and specProduct=" & pIDProduct 
							set rsObj1=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsObj1=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if%>
							<div class="pcTableRowFull">
								<div class="pcTableColumn30">
									<%if tempCat<>rsObj("categoryDesc") then
										tempCat=rsObj("categoryDesc")%>
										<%=tempCat%>:
									<%else%>
										&nbsp;
									<%end if%>
								</div>
								<div class="pcTableColumn70">		
									<%=rsObj("description")%>
									<%if rsObj1("displayQF")=True then%>
										- <%= dictLanguage.Item(Session("language")&"_viewEP_9")%><%=ArrQuantity(i)%>
									<%end if%>
								</div>
							</div>
							<%set rsObj=nothing
							set rsObj1=nothing
						next%>
						</div>
					<%end if 'End of Configuration
	
					if ArrCProduct(0)="na" then 'Additional Charges
					else%>
						<%tempCat=""%>
						<div class="pcShowBTOconfiguration">
						<div class="pcTableRowFull"><b><%= dictLanguage.Item(Session("language")&"_viewEP_4")%></b></div>
						<%for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
							query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
							set rsObj=conntemp.execute(query)
							if err.number<>0 then
								call LogErrorToDatabase()
								set rsObj=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if%>
							<div class="pcTableRowFull">
								<div class="pcTableColumn30">
									<%if tempCat<>rsObj("categoryDesc") then
										tempCat=rsObj("categoryDesc")%>
										<%=tempCat%>:
									<%else%>
										&nbsp;
									<%end if%>
								</div>
								<div class="pcTableColumn70">				
									<%=rsObj("description")%>
								</div>
							</div>
							<% set rsObj=nothing
						next%>
						</div>					
					<%end if 'End of Additional Charges%>
				<br>
				<% END IF 'Have BTO %>
	
	
				<%
				'*************************************************************************************************
				' START: GET OPTIONS
				'*************************************************************************************************
				Dim pPriceToAdd, pOptionDescrip, pOptionGroupDesc, pcv_strSelectedOptions
				Dim pcArray_SelectedOptions, pcv_strOptionsArray, cCounter, xOptionsArrayCount
				Dim pcv_strOptionsPriceArray, pcv_strOptionsPriceArrayCur, pcv_strOptionsPriceTotal
	
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START:  Get the Options for the item
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				IF len(pcv_strSelectedOptions)>0 AND pcv_strSelectedOptions<>"NULL" THEN
	
					pcArray_SelectedOptions = Split(pcv_strSelectedOptions,chr(124))
		
					pcv_strOptionsArray = ""
					pcv_strOptionsPriceArray = ""
					pcv_strOptionsPriceArrayCur = ""
					pcv_strOptionsPriceTotal = 0
					xOptionsArrayCount = 0
		
					For cCounter = LBound(pcArray_SelectedOptions) TO UBound(pcArray_SelectedOptions)
			
						' SELECT DATA SET
						' TABLES: optionsGroups, options, options_optionsGroups
						query = 		"SELECT optionsGroups.optionGroupDesc, options.optionDescrip, options_optionsGroups.price, options_optionsGroups.Wprice "
						query = query & "FROM optionsGroups, options, options_optionsGroups "
						query = query & "WHERE idoptoptgrp=" & pcArray_SelectedOptions(cCounter) & " "
						query = query & "AND options_optionsGroups.idOption=options.idoption "
						query = query & "AND options_optionsGroups.idOptionGroup=optionsGroups.idoptiongroup "	
			
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=conntemp.execute(query)
						if err.number<>0 then
							'//Logs error to the database
							call LogErrorToDatabase()
							'//clear any objects
							set rs=nothing
							'//close any connections
							call closedb()
							'//redirect to error page
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if					
			
						if Not rs.eof then 
				
							xOptionsArrayCount = xOptionsArrayCount + 1
				
							pOptionDescrip=""
							pOptionGroupDesc=""
							pPriceToAdd=""
							pOptionDescrip=rs("optiondescrip")
							pOptionGroupDesc=rs("optionGroupDesc")
				
							If Session("customerType")=1 Then
								pPriceToAdd=rs("Wprice")
								If rs("Wprice")=0 then
									pPriceToAdd=rs("price")
								End If
							Else
								pPriceToAdd=rs("price")
							End If	
				
							'// Generate Our Strings
							if xOptionsArrayCount > 1 then
								pcv_strOptionsArray = pcv_strOptionsArray & chr(124)
								pcv_strOptionsPriceArray = pcv_strOptionsPriceArray & chr(124)
								pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & chr(124)
							end if
							'// Column 4) This is the Array of Product "option groups: options"
							pcv_strOptionsArray = pcv_strOptionsArray & pOptionGroupDesc & ": " & pOptionDescrip
							'// Column 25) This is the Array of Individual Options Prices
							pcv_strOptionsPriceArray = pcv_strOptionsPriceArray & pPriceToAdd
							'// Column 26) This is the Array of Individual Options Prices, but stored as currency "scCurSign & money(pcv_strOptionsPriceTotal) "
							pcv_strOptionsPriceArrayCur = pcv_strOptionsPriceArrayCur & scCurSign & money(pPriceToAdd)
							'// Column 5) This is the total of all option prices
							pcv_strOptionsPriceTotal = pcv_strOptionsPriceTotal + pPriceToAdd
				
						end if
			
						set rs=nothing
					Next
		
				END IF	
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END:  Get the Options for the item
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			

				
				'*************************************************************************************************
				' END: GET OPTIONS
				'*************************************************************************************************
				%>
	
				<%
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' START: SHOW PRODUCT OPTIONS
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				Dim pcArray_strOptionsPrice, pcArray_strOptions, pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice, tAprice
	
				if len(pcv_strOptionsArray)>0 then 
				%>
	
	
				<%'response.write dictLanguage.Item(Session("language")&"_Custwlview_15")%>
	
				<div class="pcSpacer"></div>

				<div class="pcFormItem">
					<%
					'#####################
					' START LOOP
					'#####################	
					'// Generate Our Local Arrays from our Stored Arrays
		
					' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers	
					pcArray_strSelectedOptions = ""					
					pcArray_strSelectedOptions = Split(trim(pcv_strSelectedOptions),chr(124))
		
					' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
					pcArray_strOptionsPrice = ""
					pcArray_strOptionsPrice = Split(trim(pcv_strOptionsPriceArray),chr(124))
		
					' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
					pcArray_strOptions = ""
					pcArray_strOptions = Split(trim(pcv_strOptionsArray),chr(124))
		
					' Get Our Loop Size
					pcv_intOptionLoopSize = 0
					pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)
		
					' Start in Position One
					pcv_intOptionLoopCounter = 0
		
					' Display Our Options
					For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize
					%>
					<div class="pcFormItemFull">
			
							<%= pcArray_strOptions(pcv_intOptionLoopCounter)%>
			
													
							<% 
							tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
				
							if tempPrice="" or tempPrice=0 then
								response.write "&nbsp;"
							else 
								response.write " (" & scCurSign&money(tempPrice) & ")"
							end if 
							%>			
			
					</div>
					<%
					Next
					'#####################
					' END LOOP
					'#####################
					%>
				</div>
	
				<% 
				End if
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' END: SHOW PRODUCT OPTIONS
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				%>					
	
				<%if pxdetails<>"" then%>
					<%=pxdetails%><br>
				<%end if%>
			</div>

			<!-- Show Product Images -->
			<div class="pcViewProductRight">
				<div id="mainimgdiv" class="pcShowMainImage">
					<%if pImageUrl<>"" then%>
						<img src="<%=pcf_getImagePath("catalog",pImageUrl)%>" alt="<%= pname %>"> 
						<% ' show link to detail view image if it exists
						if pLgimageURL<>"" then%>
							<br>
						<% end if
					'if no image, show no_image.gif
					else%>
						<img src="<%=pcf_getImagePath("catalog","no_image.gif")%>" alt="Product image not available">
					<% end if%>
				</div>
				
				<%if pImageUrl<>"" then%>
					<div class="pcShowAdditionalZoom">
						<a href="javascript:enlrge('catalog/<%=pLgimageURL%>')">
							<img src="<%=pcf_getImagePath("",rsIconObj("zoom"))%>" alt="Zoom">
						</a>
					</div>
				<%end if%>
			</div>
			<!-- End of Show Product Images -->

			<div class="pcClear"></div>

			<% if trim(psDesc) <> "" then %>
				<div class="pcFormItem">
					<div class="pcFormItemFull"><hr /></div>
				</div>

				<div class="pcFormItem">
					<a name="details"></a>
					<div class="pcFormItemFull">
						<strong><%= dictLanguage.Item(Session("language")&"_viewEP_8")%></strong>
					</div>
					<div class="pcFormItemFull">
						<%=pDetails%>
					</div>
				</div>
			<% end if %>
			
			<div class="pcSpacer"></div>

			<div class="pcFormButtons">
				<a class="pcButton pcButtonReturnToRegistry" href="javascript:history.go(-1);">
					<img src="<%=pcf_getImagePath("",rslayout("RetRegistry"))%>" alt="<%= dictLanguage.Item(Session("language")&"_viewEP_6") %>" />
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_retregistry") %></span>
				</a>
			</div>

		</div>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->