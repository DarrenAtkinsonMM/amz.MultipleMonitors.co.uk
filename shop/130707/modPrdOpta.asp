<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_UpdateDates.asp" -->
<% 
dim f, pIdProduct

'***********************************************************************
' START: ON POST BACK
'***********************************************************************
If request.form("Submit2")<>"" then
	pCnt=request.Form("oCnt")

	pIdProduct=Request("idProduct")
	app_UpdPriceSub=request("updprice1")
	if app_UpdPriceSub="" then
		app_UpdPriceSub=0
	end if
	app_UpdPriceSubCC=request("updprice2")
	if app_UpdPriceSubCC="" then
		app_UpdPriceSubCC=0
	end if

	for i=1 to pCnt
		Uprice=request.form("price"&i)
		If Uprice="" then
		  Uprice="0"
		End If
		UWprice=request.form("Wprice"&i)
		If UWprice="" then
			UWprice="0"
		End If
		if scDecSign="," then
			Uprice=replacecomma(Uprice)
			UWprice=replacecomma(UWprice)
		else
			Uprice=replace(Uprice,",","")
			UWprice=replace(UWprice,",","")
		end if
		Uid=request.form("id"&i)
		USortOrder=request.form("sortOrder"&i)
		If USortOrder="" then
			USortOrder="0"
		End If
		OptActive=request.form("OptActive"&i)
		If OptActive="" then
			OptInActive="1"
		else
			OptInActive="0"	
		End If		
		query="UPDATE options_optionsGroups SET price="& Uprice &", Wprice="& UWprice &", SortOrder="& USortOrder &",InActive=" & OptInActive & " WHERE idoptoptgrp="& Uid
		set rstemp=conntemp.execute(query)				
	next
	
	
	pCnt2=request.Form("yCnt")
	for i=1 to pCnt2
		pRequired=request.Form("Required"&i)
		If pRequired<>"1" then
			pRequired="0"
		End If	
		catSort=request.Form("catSort"&i)
		If catSort="" then
			catSort="0"
		End If
		query="UPDATE pcProductsOptions SET pcProdOpt_Required="& pRequired &", pcProdOpt_Order="& catSort &" WHERE pcProdOpt_ID="& request.form("OptionGroupID"&i)
		set rstemp=conntemp.execute(query)			
	next
	set rstemp = nothing
	
	call updPrdEditedDate(request.form("idProduct"))


	'// UPDATE SUB-PRODUCTS INFORMATION
	If statusAPP="1" OR scAPP=1 Then

		query="SELECT description,sku,pcprod_Apparel,price,btoBPrice FROM Products WHERE idproduct=" & pIdProduct
		set rstemp=connTemp.execute(query)
		
		pcv_Apparel=0
		
		if not rstemp.eof then
			PR_Name=rstemp("description")
			PR_SKU=rstemp("sku")
			pcv_Apparel=rstemp("pcprod_Apparel")
			parent_price=rstemp("price")
			parent_wprice=rstemp("btoBprice")
			if parent_wprice=0 then
				parent_wprice=parent_price
			end if
		end if
		set rstemp = nothing
	
		Dim pcArr(100,2)
		
		IF pcv_Apparel=1 THEN
		for i=1 to pCnt2
			pcArr(i,0)=i	
			catSort=request.Form("catSort"&i)
			If catSort="" then
				catSort="0"
			End If
			pcArr(i,1)=catSort
		next
		
		for i=1 to pCnt2
			for j=i+1 to pCnt2
				if pcArr(i,1)>pcArr(j,1) then
					tmp1=pcArr(i,0)
					tmp2=pcArr(i,1)
					pcArr(i,0)=pcArr(j,0)
					pcArr(i,1)=pcArr(j,1)
					pcArr(j,0)=tmp1
					pcArr(j,1)=tmp2
				end if
			next
		next
		
		query="SELECT idproduct,pcprod_Relationship FROM Products WHERE pcprod_ParentPrd=" & pIdProduct & " AND removed=0"
		set rstemp=connTemp.execute(query)
		
		if not rstemp.eof then
			tmpArr=rstemp.GetRows()
			intCount=ubound(tmpArr,2)
			set rstemp=nothing
			
			For i=0 to intCount
				sub_price=parent_price
				sub_Wprice=parent_wprice
				sub_addprice=0
				sub_addWprice=0
				sub_ID=tmpArr(0,i)
				sub_Relation=tmpArr(1,i)
				tmp1=split(sub_Relation,"_")
				tmp2=tmp1
				sub_Relation=tmp2(0)
				sub_Name=PR_Name & " ("
				sub_SKU=PR_SKU
				for j=1 to pCnt2
					tmp2(j)=tmp1(pcArr(j,0))
					sub_Relation=sub_Relation & "_" & tmp2(j)
					query="SELECT options.optionDescrip,options.pcOpt_Code,options_optionsGroups.price,options_optionsGroups.wprice FROM options INNER JOIN options_optionsGroups ON options.idOption=options_optionsGroups.idOption WHERE options_optionsGroups.idoptoptgrp=" & tmp2(j)
					set rstemp=connTemp.execute(query)
					if not rstemp.eof then
						if j>1 then
							sub_Name=sub_Name & " - "
						end if
						sub_Name=sub_Name & rstemp("optionDescrip")
						sub_SKU=sub_SKU & rstemp("pcOpt_Code")
						sub_addprice1=rstemp("price")
						sub_addWprice1=rstemp("wprice")
						if sub_addWprice1=0 then
							sub_addWprice1=sub_addprice1
						end if
						sub_addprice=sub_addprice+sub_addprice1
						sub_addWprice=sub_addWprice+sub_addWprice1
					end if
					set rstemp=nothing
				next
				
				sub_Name=sub_Name & ")"
				
				sub_price=sub_price+sub_addprice
				sub_wprice=sub_wprice+sub_addWprice
				if sub_wprice=sub_price then
					sub_wprice=0
				end if
				
				tmp_Str=""
				if app_UpdPriceSub="1" then
					tmp_Str="price=" & sub_price & ",btoBPrice=" & sub_wprice & ",pcprod_AddPrice=" & sub_addprice & ",pcprod_AddWPrice=" & sub_addWprice & ","
				end if
				
				sub_Name=replace(sub_Name,"'","''")
				sub_Name=replace(sub_Name,"""","""""")
				
				query="UPDATE Products SET " & tmp_Str & "description=N'" & sub_Name & "',pcprod_Relationship='" & sub_Relation & "' WHERE idproduct=" & sub_ID
				set rstemp=connTemp.execute(query)
				set rstemp=nothing
				
				call pcs_hookProductModified(sub_ID, "")
						
				'// START: Re-Calculate all Customer Category Pricing for this Product
				SP_id=sub_ID
				SP_AddPrice=sub_addprice
				SP_AddWPrice=sub_addWprice
				pcv_ParentProduct=pIdProduct
				
				if app_UpdPriceSubCC="1" then
				
					query="SELECT pcCC_Pricing.idCC_Price, pcCC_Pricing.idcustomerCategory FROM pcCC_Pricing WHERE idProduct="&SP_id&";"
					SET rsPBPObj=Server.CreateObject("ADODB.RecordSet")
					SET rsPBPObj=conntemp.execute(query)
					if NOT rsPBPObj.eof then
						Do while NOT rsPBPObj.eof
							
							intIdcustomerCategory=rsPBPObj("idcustomerCategory")
							
							'// Get information about this Customer Category
							query="SELECT pcCustomerCategories.pcCC_CategoryType, pcCustomerCategories.pcCC_ATB_Off, pcCustomerCategories.pcCC_ATB_Percentage FROM pcCustomerCategories WHERE pcCustomerCategories.idcustomerCategory="&intIdcustomerCategory&";"
							SET rsCObj=Server.CreateObject("ADODB.RecordSet")
							SET rsCObj=conntemp.execute(query)
							if not rsCObj.eof then	
													
									pcv_strCategoryType=rsCObj("pcCC_CategoryType")
									pcv_strCCATBOff=rsCObj("pcCC_ATB_Off")
									pcv_strCCATBPercentage=rsCObj("pcCC_ATB_Percentage")
									
									'// Choose Wholesale or Retail Differential									
									If pcv_strCCATBOff="Retail" then
										pcv_Addprice=SP_AddPrice
									Else
										pcv_Addprice=SP_AddWPrice
									End If										
									if pcv_Addprice<>"" then
									else
									pcv_Addprice=0
									end if								
									
									'// If ATB Make Adjustments to Differential
									if pcv_strCategoryType="ATB" then
										'// Apply the ATB to the differential	
										pcv_Addprice=pcv_Addprice-(pcf_Round(pcv_Addprice*(cdbl(pcv_strCCATBPercentage)/100),2))
									end if
									
									'// Get the Base Customer Category Price
									query="SELECT pcCC_Pricing.idcustomerCategory, pcCC_Pricing.idProduct, pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE (((pcCC_Pricing.idcustomerCategory)="&intIdcustomerCategory&") AND ((pcCC_Pricing.idProduct)="&pcv_ParentProduct&"));"
									SET rsPriceObj=server.CreateObject("ADODB.RecordSet")
									SET rsPriceObj=conntemp.execute(query)
									if rsPriceObj.eof then
										dblpcCC_Price=0
									else
										dblpcCC_Price=rsPriceObj("pcCC_Price")
									end if
									SET rsPriceObj=nothing
									
									intpcCC_Price=dblpcCC_Price
									
									'// Calculate the New Sub-Product Customer Category Price (base price + new differential - ATB Discounts)			
									pcv_PPrice=cdbl(intpcCC_Price)+cdbl(pcv_Addprice)												
	
									'// Update the Sub-Product Customer Category Price
									query="Update pcCC_Pricing set pcCC_Pricing.pcCC_Price="&pcv_PPrice&" WHERE (((pcCC_Pricing.idcustomerCategory)="&intIdcustomerCategory&") AND ((pcCC_Pricing.idProduct)="&SP_id&"));"
									SET rsIObj=Server.CreateObject("ADODB.RecordSet")
									SET rsIObj=conntemp.execute(query)
									SET rsIObj=nothing
	
							end if
							SET rsCObj=nothing
							
						rsPBPObj.movenext
						Loop
					end if
					SET rsPBPObj=nothing			
						
				end if
				'// END: Re-Calculate all Customer Category Pricing for this Product
			Next
		end if
		END IF

	End If '// If statusAPP="1" OR scAPP=1 Then

	call closeDb()
	response.redirect "modPrdOpta.asp?s=1&msg="&Server.Urlencode("You have successfully updated your product attributes.")&"&idProduct="& request.form("idProduct")
	response.end

End If
'***********************************************************************
' END: ON POST BACK
'***********************************************************************



'***********************************************************************
' START: ON LOAD
'***********************************************************************
'// Form parameter 
pIdProduct=Request("idProduct")
if not validNum(pidproduct) then
   call closeDb()
response.redirect "msg.asp?message=2"
end if

'// Get item details from db

query="SELECT idProduct, description, pcprod_Apparel  FROM products WHERE products.idProduct=" & pIdProduct
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)
if err.number <> 0 then
	set rstemp=nothing
	
    call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in modPrdOpta.asp: "&Err.Description) 
end if

'// set data into local variables
pIdProduct=rstemp("idProduct")
pDescription=rstemp("description")

pcv_Apparel=rstemp("pcprod_Apparel")

pcv_HaveSubPrds=0
if pcv_Apparel="1" then

	query="SELECT idproduct FROM Products WHERE pcProd_ParentPrd=" & pIdProduct & " and removed=0;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcv_HaveSubPrds=1
	end if
	set rs=nothing

end if

' SELECT DATA SET
' TABLES: products, pcProductsOptions, optionsgroups, ptions_optionsGroups
query = 		"SELECT DISTINCT optionsGroups.OptionGroupDesc, pcProductsOptions.pcProdOpt_ID, pcProductsOptions.idOptionGroup, pcProductsOptions.pcProdOpt_Required, pcProductsOptions.pcProdOpt_Order "
query = query & "FROM products "
query = query & "INNER JOIN ( "
query = query & "pcProductsOptions INNER JOIN ( "
query = query & "optionsgroups "
query = query & "INNER JOIN options_optionsGroups "
query = query & "ON optionsgroups.idOptionGroup = options_optionsGroups.idOptionGroup "
query = query & ") ON optionsGroups.idOptionGroup = pcProductsOptions.idOptionGroup "
query = query & ") ON products.idProduct = pcProductsOptions.idProduct "
query = query & "WHERE products.idProduct=" & pidProduct &" "
query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
query = query & "ORDER BY pcProductsOptions.pcProdOpt_Order, optionsGroups.OptionGroupDesc;"
set rs=server.createobject("adodb.recordset")
set rs=conntemp.execute(query)	
if err.number<>0 then
	'//Logs error to the database
	'call LogErrorToDatabase()
	'//clear any objects
	'set rs=nothing
	'//close any connections
	'
	'//redirect to error page
	'response.redirect "techErr.asp?err="&pcStrCustRefID
	set rs=nothing
	
	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in modPrdOpta.asp: "&Err.Description)
end if
	
'***********************************************************************
' END: ON LOAD
'***********************************************************************


'***********************************************************************
' START: MODE DELETE
'***********************************************************************
If Request("mode")="DEL" then

	idoptoptgrp=Request("id")
	
	If statusAPP="1" OR scAPP=1 Then
		query="UPDATE Products SET removed=-1,active=0 WHERE pcprod_ParentPrd="& pIdProduct & " AND removed=0 AND ((pcprod_Relationship like '%[_]" & idoptoptgrp & "[_]%') OR (pcprod_Relationship like '%[_]" & idoptoptgrp & "'))"
		set rstemp=conntemp.execute(query)
		set rstemp=nothing
	End If
	
	'// Check the Option Group Number
	strSQL="SELECT idOptionGroup FROM options_optionsGroups WHERE idoptoptgrp="& idoptoptgrp &";"
	set rstemp=conntemp.execute(strSQL)	
	pIdOptionGroup = rstemp("idOptionGroup")
	
	'// Delete this option
	query="DELETE FROM options_optionsGroups WHERE idoptoptgrp="& idoptoptgrp
	set rstemp=conntemp.execute(query)
	
	'// Check if all options have been removed.
	strSQL="SELECT * FROM options_optionsGroups WHERE idproduct="& pIdProduct &" AND idoptionGroup="& pIdOptionGroup &";"
	set rstemp=conntemp.execute(strSQL)							
	if rstemp.eof then
		'// It is NOT related
		contgo=1
	end if	
	
	'// If all Options have been removed then delete the corrisponding record in pcProductOptions
	if contgo=1 then				
		strSQL="DELETE FROM pcProductsOptions WHERE idproduct="& pIdProduct &" AND idoptionGroup="& pIdOptionGroup &";"
		set rstemp=conntemp.execute(strSQL)
	end if	
	
	set rstemp=nothing
	
	call updPrdEditedDate(pIdProduct)
	
	
	call closeDb()
response.redirect "modPrdOpta.asp?s=1&msg="&Server.Urlencode("Your deletion was successful.")&"&idProduct="& pIdProduct
	response.end
End If
'***********************************************************************
' END: MODE DELETE
'***********************************************************************

pageTitle="Modify Product Options for: <strong>" & pDescription & "</strong>"
%>
<!--#include file="AdminHeader.asp"-->
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
<form method="post" name="modifyProduct" action="modPrdOpta.asp" class="pcForms">                 
<input type="hidden" name="idproduct" value="<%=pidProduct%>">
		<table class="pcCPcontent">
			<tr>
				<td colspan="6">
					<div class="cpOtherLinks"><a href="modPrdOpta1.asp?idproduct=<%=pIdProduct%>">Add New Option Group</a> | <a href="FindProductType.asp?id=<%=pIdProduct%>">Edit Product</a> | <a href="../pc/viewPrd.asp?idProduct=<%=pIdProduct%>&adminPreview=1" target="_blank">Preview</a></div>				
					<% If statusAPP="1" OR scAPP=1 Then %>
						NOTE: 
						<% if pcv_Apparel="1" then %>
	                    This is an <u>apparel</u> product. You can <a href="app-subPrdsMngAll.asp?idproduct=<%=pIdProduct%>">create &amp; manage sub-products</a> by combining the options listed below.
	                    <% else %>
	                    This is <u>not</u> an apparel product. You cannot create sub-products. If you need to change this setting, <a href="modifyProduct.asp?idproduct=<%=pIdProduct%>">edit</a> the product details.
	                    <% end if %>
					<% End If %>
				</td>
			</tr>
            <tr>
            	<td colspan="6" class="pcCPspacer"></td>
            </tr>
			<%									
			' If we have data	
			If NOT rs.eof Then
				pcv_intOptionGroupCount = 0 '// keeps count of the number of options
				xOptionsCnt = 0 '// keeps count of the number of required options
				oCnt = 0
				yCnt = 0
				
				Do until rs.eof				
				yCnt = yCnt + 1	
					
					'// Get the Group Name
					pcv_strOptionGroupDesc=rs("OptionGroupDesc")
					'// Get the Group ID
					pcv_strOptionGroupID=rs("idOptionGroup")
					'// Is it required
					pcv_strOptionRequired=rs("pcProdOpt_Required")			
					'// Primary Key
					pcv_strProdOpt_ID=rs("pcProdOpt_ID")
					'// Sort Order
					strCatSort=rs("pcProdOpt_Order")
					'// Start: Do Option Count
					pcv_intOptionGroupCount = pcv_intOptionGroupCount + 1 
					'// End: Do Option Count
					
					'// Get the number of the Option Group
					pcv_strOptionGroupCount = pcv_intOptionGroupCount
					
					'// Start: Do Required Option Count
					if IsNull(pcv_strOptionRequired) OR pcv_strOptionRequired="" then
							pcv_strOptionRequired=0 '// not required // else it is "1"
					end if			
					if pcv_strOptionRequired=1 then							
						' Keep Tally
						xOptionsCnt = xOptionsCnt + 1
					end if
					'// End: Do Required Option Count
				
					'// Add Table Here
					%>
					<tr bgcolor="#e5e5e5"> 
						<td <% if pcv_Apparel="0" then %>colspan="4"<% else %>colspan="3"<% end if %>> 
							<span style="font-size:14px; font-weight: bold;"><%=pcv_strOptionGroupDesc%></span> 
						</td>
						<td colspan="2" align="right">
							Order: <input type="text" name="catSort<%=yCnt%>" size="1" maxlength="3" value="<%=strCatSort%>" style="text-align: right; font-size: 8pt; font-weight: bold; color: #000000; background-color: #99CCFF">
						</td>
					</tr>	
					<tr>
						<th nowrap>&nbsp;</th>	
						<th nowrap>Option Group - Option Attribute</th>							

							<%if pcv_Apparel="0" then%>							
							<th nowrap><input type="checkbox" name="A<%=yCnt%>" value="1" onClick="javascript:RunCheck<%=yCnt%>(this.checked);" class="clearBorder">&nbsp;Active</th>
							<%end if%>

						<th nowrap>Price</th>
						<th nowrap>Wholesale Price</th>
						<th nowrap>Order</th>
					</tr>	
                    <tr>
                        <td colspan="6" class="pcCPspacer"></td>
                    </tr>
					<%
					' SELECT DATA SET
					' TABLES: options_optionsGroups, options
					query = 		"SELECT options_optionsGroups.InActive, options_optionsGroups.price, options_optionsGroups.Wprice, "
					query = query & "options_optionsGroups.idoptoptgrp, options_optionsGroups.sortOrder, options.idoption, options.optiondescrip "
					query = query & "FROM options_optionsGroups "
					query = query & "INNER JOIN options "
					query = query & "ON options_optionsGroups.idOption = options.idOption "
					query = query & "WHERE options_optionsGroups.idOptionGroup=" & pcv_strOptionGroupID &" "
					query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
					query = query & "ORDER BY options_optionsGroups.sortOrder, options.optiondescrip;"	
					set rs2=server.createobject("adodb.recordset")
					set rs2=conntemp.execute(query)	

				
					' If we have data
					if NOT rs2.eof then

						'// clean up the option group description
						if pcv_strOptionGroupDesc<>"" then
							pcv_strOptionGroupDesc=replace(pcv_strOptionGroupDesc,"""","&quot;")
						end if 							
											
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' Start Loop
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						tmp_start=0
						tmp_end=0
						do until rs2.eof			
						oCnt = oCnt + 1
						if tmp_start=0 then
							tmp_start=oCnt
						end if
						OptInActive=rs2("InActive") ' Is it active?
						if IsNull(OptInActive) OR OptInActive="" then
							OptInActive="0"
						end if
						
						dblOptPrice=rs2("price") '// Price
						dblOptWPrice=rs2("Wprice") '// WPrice
						intIdOptOptGrp=rs2("idoptoptgrp") '// The Id of the Option Group
						intIdOption=rs2("idoption") '// The Id of the Option
						strOptionDescrip=rs2("optiondescrip") '// A description of the Option
						pcv_strSortOrder=rs2("sortorder")
						If statusAPP="1" OR scAPP=1 Then

							query="SELECT idproduct FROM Products WHERE pcprod_ParentPrd=" & pIdProduct & " AND ((pcprod_Relationship like '%[_]" & intIdOptOptGrp & "[_]%') OR (pcprod_Relationship like '%[_]" & intIdOptOptGrp & "'));"
							set rsQ=connTemp.execute(query)
							pcv_HaveSub=0
							if not rsQ.eof then
								pcv_HaveSub=1
							end if
							set rsQ=nothing

						End If
				
						'**************************************************************************************************
						' START: Dispay the Options
						'**************************************************************************************************
						%>
						<tr>                               
							<td width="6%">  
							<% If statusAPP="1" OR scAPP=1 Then %> 
								<img src="images/delete2.gif" width="23" height="18" border="0" alt="Remove" onMouseOver="document.body.style.cursor = 'hand';" onMouseOut="document.body.style.cursor = 'default';" onClick="javascript:<%if pcv_HaveSub=1 then%>if (confirm('This attribute was assigned to one or more sub-products. When you delete it, all related sub-products will be also removed. Are you sure that you want to delete it?'))<%end if%> location='modPrdOpta.asp?mode=DEL&id=<%=intIdOptOptGrp%>&idproduct=<%=pIdProduct%>';">
							<% Else %>
								<a href="modPrdOpta.asp?mode=DEL&id=<%=intIdOptOptGrp%>&idproduct=<%=pIdProduct%>">
									<img src="images/delete2.gif" width="23" height="18" border="0" alt="Remove">
								</a>
							<% End If %>
							</td>
							<td width="60%"> 												
								<%=pcv_strOptionGroupDesc%> -  <b><%=strOptionDescrip%></b>
							</td>

							<% if pcv_Apparel="0" then %>
							<td nowrap>
								<input name="OptActive<%=oCnt%>" type="checkbox" value="1" <%if (OptInActive<>"") and (OptInActive="1") then%><%else%>checked<%end if%> class="clearBorder">
							</td>
							<% end if %>

							<td nowrap>
								<%=scCurSign%>
								<input type="text" name="price<%=oCnt%>" value="<%=money(dblOptPrice)%>" size="6" maxlength="10">
							</td>
							<td nowrap>
								<%=scCurSign%> 
								<input type="text" name="Wprice<%=oCnt%>" value="<%=money(dblOptWPrice)%>" size="6" maxlength="10">
								<input type="hidden" name="id<%=oCnt%>" value="<%=intIdOptOptGrp%>">
							
								<% if pcv_Apparel<>"0" then %>
									<input name="OptActive<%=oCnt%>" type="hidden" value="1">
								<% end if %>
							
							</td>
							<td nowrap>          
								<input name="sortOrder<%=oCnt%>" type="text" size="2" value="<%=pcv_strSortOrder%>">
							</td>
						</tr>										
						<% 
						'**************************************************************************************************
						' END: Dispay the Options
						'**************************************************************************************************
					rs2.movenext 
					loop
					if tmp_start>0 then
						tmp_end=oCnt
					end if
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END Loop
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					set rs2=nothing	
				end if
				%>				
						
				<tr>                               
					<td colspan="6" nowrap style="border-top: 1px solid #CCC;">
				
					<% if pcv_Apparel="0" then %>
						<% if pcv_strOptionRequired=1 then %>
							<input type="checkbox" name="Required<%=yCnt%>" value="1" checked class="clearBorder">
						<% else %>
							<input type="checkbox" name="Required<%=yCnt%>" value="1" class="clearBorder">
						<% end If %> <b>Required Option
						&nbsp;<%=chr(124)%>&nbsp;						
					<% else %>
						<input type="hidden" name="Required<%=yCnt%>" value="1">
					<% end if %>
				
					<a href="modPrdOpta3.asp?idproduct=<%=pidProduct%>&IdOptionGroup=<%=pcv_strOptionGroupID%>">Add More Attributes</a></b>
					<input type="hidden" name="OptionGroupID<%=yCnt%>" value="<%=pcv_strProdOpt_ID%>">
					</td>
				</tr>                            
				<tr> 								  
					<td colspan="6">
						<script type=text/javascript>
							function RunCheck<%=yCnt%>(tstatus)
							{
								CheckUncheckBoxes(<%=tmp_start%>,<%=tmp_end%>,tstatus);
							}
						</script>
					&nbsp;</td>
				</tr>					
				<%
				rs.movenext
			Loop			
			set rs=nothing
			%>
			<tr> 
				<td colspan="6"><hr></td>
			</tr>
	
			<% if pcv_Apparel<>"0" then %>

				<tr> 
					<td colspan="6"><input type="checkbox" name="updprice1" value="1" class="clearBorder"> Update sub-product prices using current attribute prices</td>
				</tr>
				<% 
				query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories ORDER BY pcCC_Name ASC;"
				SET rs=Server.CreateObject("ADODB.RecordSet")
				SET rs=conntemp.execute(query)
				if not rs.eof then
					Dim pcv_strIsCustomerCategories
					pcv_strIsCustomerCategories=True
				end if	
				if pcv_strIsCustomerCategories=True then 
					%>
					<tr> 
						<td colspan="6"><input type="checkbox" name="updprice2" value="1" class="clearBorder"> Re-Calculate Sub-Product Customer Category prices using current attribute prices</td>
					</tr>
					<% 'end if %>
					<tr> 
						<td colspan="6" class="pcSpacer"><hr></td>
					</tr>
				<% end if %>

			<% End If %>
			
			<tr> 
				<td colspan="6" align="center">
					<input type="submit" name="Submit2" value="Update" class="btn btn-primary">
					&nbsp;
					<input type="button" class="btn btn-default"  name="Clone" value="Copy to other products" onClick="location.href='ApplyOptionsMulti2.asp?action=add&prdlist=<%=pIdProduct%>'">
					&nbsp;
					<input type="button" class="btn btn-default"  name="Button" value="Manage Options" onClick="location.href='manageOptions.asp'">
					&nbsp;
			
					<% if pcv_Apparel="1" then %>
						<input type="button" name="Button" value="Manage Sub-Products" onClick="location.href='app-subPrdsMngAll.asp?idproduct=<%=pIdProduct%>'">
						&nbsp;
					<% end if %>

					<input type="button" class="btn btn-default" name="Button" value="Locate Another Product" onClick="location.href='LocateProducts.asp?cptype=0'">
					<script type=text/javascript>
						function CheckUncheckBoxes(tmp_start,tmp_end,tstatus)
						{
							for (var i=tmp_start;i<=tmp_end;i++)
							eval("document.modifyProduct.OptActive" + i).checked=tstatus;
						}
					</script>
				</td>
			</tr>	
				
			<% Else %>	
														
                <tr> 								  
                    <td colspan="6">
                        <div class="pcCPmessage">No option group has been added to this product.</div>
                    </td>
                </tr>
                <tr> 
                    <td colspan="6"><hr></td>
                </tr>
                <tr> 
                    <td colspan="6" align="center">
                    
                        <input type="button" class="btn btn-default"  name="Button" value="Add New Option Group" onClick="location.href='modPrdOpta1.asp?idproduct=<%=pIdProduct%>'" class="btn btn-primary">
                        &nbsp;
                        <input type="button" class="btn btn-default"  name="Button" value="Manage Options" onClick="location.href='manageOptions.asp'">
                        &nbsp;
                        <input type="button" class="btn btn-default"  name="Button" value="Locate Another Product" onClick="location.href='LocateProducts.asp?cptype=0'">
                        
                    </td>
                </tr>
				
			    <%	
			End If
			%>												
		</table>
	<input type="hidden" name="oCnt" value="<%=oCnt%>">
	<input type="hidden" name="yCnt" value="<%=yCnt%>">
</form>
<!--#include file="AdminFooter.asp"-->