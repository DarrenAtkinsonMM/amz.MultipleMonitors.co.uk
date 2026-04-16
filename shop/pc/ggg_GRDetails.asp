<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="CustLIv.asp"-->
<%
pIdCustomer=session("idCustomer")
gIDEvent=getUserInput(request("IDEvent"),0)

if gIDEvent<>"" then
	query="select pcEv_IDEvent,pcEv_Name,pcEv_Date,pcEv_Type,pcEv_IncGcs from pcEvents where pcEv_IDCustomer=" & pIDCustomer & " and pcEv_IDEvent=" & gIDEvent
	set rstemp=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rstemp.eof then
		response.redirect "ggg_manageGRs.asp"
	else
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
		if request("addgc")="1" then
			if geincgc="1" then
				query="select IDProduct from Products where pcprod_GC=1 and removed=0 and active<>0"
				set rstemp=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rstemp=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				do while not rstemp.eof
					IDProduct=rstemp("IDProduct")
					query="select pcEP_IDProduct from pcEvProducts where pcEP_IDEvent=" & gIDEvent & " and pcEP_IDProduct=" & IDProduct & " and pcEP_GC=1"
					set rs1=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs1=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					if rs1.eof then
						query="insert into pcEvProducts (pcEP_IDEvent,pcEP_IDProduct,pcEP_GC) values (" & gIDEvent & "," & IDProduct & ",1)"
						set rs1=conntemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rs1=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
					end if
					set rs1=nothing
					rstemp.MoveNext
				loop
				set rs=nothing
			end if
		end if
	end if
	set rstemp=nothing
end if

if request("action")="update" then
	Count=getUserInput(request("Count"),0)
	Count1=getUserInput(request("Count1"),0)
	
	For k=1 to Count1
		geID=getUserInput(request("geID" & k),0)
		geadd=getUserInput(request("add" & k),0)
		if geadd="" then
			geadd="0"
		end if
		if geID<>"" then
			query="update pcEvProducts set pcEP_Qty=pcEP_Qty+" & geadd & " where pcEP_IDEvent=" & gIDEvent & " and pcEP_ID=" & geID
			set rstemp=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		end if
	Next
	For k=1 to Count
		geID=getUserInput(request("geID" & k),0)
		geadd=getUserInput(request("add" & k),0)
		if geadd="" then
			geadd="0"
		end if
		gedel=getUserInput(request("del" & k),0)
		if gedel="" then
			gedel="0"
		end if
		if (geID<>"") and (gedel="1") then
			query="delete from pcEvProducts where pcEP_IDEvent=" & gIDEvent & " and pcEP_ID=" & geID
			set rstemp=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rstemp=nothing
		end if
	Next
	msg=dictLanguage.Item(Session("language")&"_GRDetails_11")
end if	

%>
<!--#include file="header_wrapper.asp"--> 

<%query="SELECT products.sku,products.description,pcEvProducts.pcEP_ID,pcEvProducts.pcEP_Qty,pcEvProducts.pcEP_HQty,pcEvProducts.pcEP_GC,pcEP_OptionsArray FROM products,pcEvProducts WHERE pcEvProducts.pcEP_IDEvent=" & gIDEvent & " AND products.idproduct=pcEvProducts.pcEP_IDProduct AND products.removed=0 AND ((products.active<>0) OR ((products.active=0) AND (products.pcProd_ParentPrd>0))) ORDER BY products.Description ASC,pcEvProducts.pcEP_GC ASC"
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if%>
<div id="pcMain">
	<div class="pcMainContent">
		<form method="post" name="Form1" action="ggg_GRDetails.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
			<h1><%= dictLanguage.Item(Session("language")&"_GRDetails_1")%>"<%=geName%>"<%= dictLanguage.Item(Session("language")&"_GRDetails_1a")%></h1>

			<div class="pcFormItem">
				<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_GRDetails_1c")%><strong><%=geName%></strong></div>
				<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_GRDetails_1b")%><%=geDate%></div>
				<%if gType<>"" then%>
					<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_GRDetails_1d")%><%=gType%></div>
				<%end if%>
			</div>

			<% If msg<>"" then %>
				<div class="pcErrorMessage"><%=msg%></div>
			<% end if %>	
	
			<%IF rstemp.eof then%>
				<div class="pcSpacer"></div>

				<div class="pcFormItem">
					<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_GRDetails_10")%></div>
				</div>
			
				<div class="pcSpacer"></div>

				<div class="pcFormButtons">
					<a class="pcButton pcButtonBack" href="ggg_manageGRs.asp">
						<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
					</a>
				</div>

			<%ELSE%>
				<div class="pcSpacer"></div>

				<%
					col_SKUClass		= "pcCol-2 pcGRDetails_SKU"
					col_NameClass		= "pcCol-5 pcGRDetails_Name"
					col_ItemsClass	= "pcCol-1 pcGRDetails_Items"
					col_AvailClass	= "pcCol-1 pcGRDetails_Avail"
					col_AddClass		= "pcCol-1 pcGRDetails_Add"
					col_RemoveClass	= "pcCol-1 pcGRDetails_Remove"
				%>
				<div id="pcTableGRDetails" class="pcTable">
					<div class="pcTableHeader">
						<div class="<%= col_SKUClass %>"><%= dictLanguage.Item(Session("language")&"_GRDetails_2")%></div>
						<div class="<%= col_NameClass %>"><%= dictLanguage.Item(Session("language")&"_GRDetails_3")%></div>
						<div class="<%= col_ItemsClass %>"><%= dictLanguage.Item(Session("language")&"_GRDetails_4")%></div>
						<div class="<%= col_AvailClass %>"><%= dictLanguage.Item(Session("language")&"_GRDetails_5")%></div>
						<div class="<%= col_AddClass %>"><%= dictLanguage.Item(Session("language")&"_GRDetails_6")%></div>
						<div class="<%= col_RemoveClass %>"><%= dictLanguage.Item(Session("language")&"_GRDetails_7")%></div>
					</div>
			
					<div class="pcSpacer"></div>

				<%
				Count=0
				do while not rstemp.eof
					gsku=rstemp("sku")
					gname=rstemp("description")
					geID=rstemp("pcEP_ID")
					gQty=rstemp("pcEP_Qty")
					if gQty<>"" then
					else
						gQty="0"
					end if
					gHQty=rstemp("pcEP_HQty")
					if gHQty<>"" then
					else
						gHQty="0"
					end if
					gGC=rstemp("pcEP_GC")
	
					pcv_strSelectedOptions=""
					pcv_strSelectedOptions = rstemp("pcEP_OptionsArray")
					pcv_strSelectedOptions=pcv_strSelectedOptions&""
	
					if gGC<>"1" then
						Count=Count+1%>
						<div class="pcTableRow">
							<div class="<%= col_SKUClass %>"><%=gsku%></div>
							<div class="<%= col_NameClass %>"><%=gname%>
								<%Dim pPriceToAdd, pOptionDescrip, pOptionGroupDesc
								Dim pcArray_SelectedOptions, pcv_strOptionsArray, cCounter, xOptionsArrayCount
								Dim pcv_strOptionsPriceArray, pcv_strOptionsPriceArrayCur, pcv_strOptionsPriceTotal
	
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
			
									set rsQ=server.CreateObject("ADODB.RecordSet")
									set rsQ=conntemp.execute(query)
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
			
									if Not rsQ.eof then 
				
									xOptionsArrayCount = xOptionsArrayCount + 1
				
									pOptionDescrip=""
									pOptionGroupDesc=""
									pPriceToAdd=""
									pOptionDescrip=rsQ("optiondescrip")
									pOptionGroupDesc=rsQ("optionGroupDesc")
				
									If Session("customerType")=1 Then
										pPriceToAdd=rsQ("Wprice")
										If rsQ("Wprice")=0 then
											pPriceToAdd=rsQ("price")
										End If
									Else
										pPriceToAdd=rsQ("price")
									End If	
				
									'// Generate Our Strings
									%>
										<br />
										<span class="small ident"><%= pOptionGroupDesc & ": " & pOptionDescrip %></span>
									<%
				
									end if
			
									set rsQ=nothing
								Next
		
								END IF%>
							</div>

							<div class="<%= col_ItemsClass %>"><%=gQty%></div>
							<div class="<%= col_AvailClass %>"><%if clng(GQty)-clng(gHQty)>=0 then%><%=clng(GQty)-clng(gHQty)%><%else%>0<%end if%></div>
							<div class="<%= col_AddClass %>">
								<input type="hidden" name="geID<%=Count%>" value="<%=geID%>">
								<input type="hidden" name="remainqty<%=Count%>" value="<%=(clng(GQty)-clng(gHQty))*(-1)%>">

								<input name="add<%=Count%>" value="0" size="3">
							</div>
							<div class="<%= col_RemoveClass %>">
								<%if gHQty="0" then%>
									<input type="checkbox" name="del<%=Count%>" value="1" class="clearBorder">
								<%end if%>
							</div>
						</div>
					<%
					end if
					rstemp.MoveNext
				loop
				set rstemp=nothing

				Count1=Count

				query="select products.sku,products.description,pcEvProducts.pcEP_ID,pcEvProducts.pcEP_Qty,pcEvProducts.pcEP_HQty,pcEvProducts.pcEP_GC from products,pcEvProducts where pcEvProducts.pcEP_IDEvent=" & gIDEvent & " and products.idproduct=pcEvProducts.pcEP_IDProduct and pcEvProducts.pcEP_GC=1 and products.removed=0 order by products.Description asc,pcEvProducts.pcEP_GC asc"
				set rstemp=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rstemp=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if

				If Not rstemp.eof Then %>
				<div class="pcSpacer"></div>
				<div class="pcTableHeader">
					<div class="pcTableRowFull"><%= dictLanguage.Item(Session("language")&"_GRDetails_8")%></div>
				</div>
				<div class="pcSpacer"></div>
				<%do while not rstemp.eof
					gsku=rstemp("sku")
					gname=rstemp("description")
					geID=rstemp("pcEP_ID")
					gQty=rstemp("pcEP_Qty")
					if gQty<>"" then
					else
						gQty="0"
					end if
					gHQty=rstemp("pcEP_HQty")
					if gHQty<>"" then
					else
						gHQty="0"
					end if
					gGC=rstemp("pcEP_GC")
					if ((gGC="1") and (geincgc="1")) or (clng(gHQty)>0) then
						Count=Count+1%>
						<div class="pcTableRow">
							<div class="<%= col_SKUClass %>"><%=gsku%></div>
							<div class="<%= col_NameClass %>"><%=gsku%><%=gname%></div>
							<div class="<%= col_ItemsClass %>">&nbsp;</div>
							<div class="<%= col_AvailClass %>">&nbsp;</div>
							<div class="<%= col_AddClass %>"><input type=hidden name="geID<%=Count%>" value="<%=geID%>"></div>
							<div class="<%= col_RemoveClass %>">
								<%if gHQty="0" then%>
								<input type=checkbox name="del<%=Count%>" value="1">
								<%end if%>
							</div>
						</div>
					<%
					end if
					rstemp.MoveNext
				loop
				END IF 'Have GCs
				set rstemp=nothing%>
				<%if geincgc="1" then
				query="select IDProduct from Products where pcprod_GC=1 and removed=0 and active<>0"
				set rstemp=connTemp.execute(query)
				if not rstemp.eof then%>
				<div class="pcTableRowFull">
					<a href="ggg_GRDetails.asp?IDEvent=<%=gIDEvent%>&addgc=1"><%= dictLanguage.Item(Session("language")&"_GRDetails_12")%></a>
				</div>
				<%end if
				end if%>

				</div>

				<div class="pcSpacer"></div>

				<div class="pcFormItem">
					<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_GRDetails_13")%></div>
				</div>

				<div class="pcSpacer"></div>

				<div class="pcFormButtons">
					<button class="pcButton pcButtonUpdateRegistry" id="submit" name="submit" value="<%= dictLanguage.Item(Session("language")&"_instGR_12")%>">
						<img src="<%=pcf_getImagePath("",rslayout("UpdRegistry"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_updregistry") %>">
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_updregistry") %></span>
					</button>

					<a class="pcButton pcButtonBack" href="ggg_manageGRs.asp">
						<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
					</a>

					<input type=hidden name="IDEvent" value="<%=gIDEvent%>">
					<input type=hidden name="Count" value="<%=Count%>">
					<input type=hidden name="Count1" value="<%=Count1%>">
				</div>
				<%END IF 'Have products%>
		</form>
	</div>
</div>

<script type=text/javascript>
function isDigit(s)
{
	var test=""+s;
	if(test=="-"||test=="+"||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
	{
		return(true) ;
	}
	return(false);
}
	
function allDigit(s)
{
	var test=""+s ;
	for (var k=0; k <test.length; k++)
	{
		var c=test.substring(k,k+1);
		if (isDigit(c)==false)
		{
			return (false);
		}
	}
	return (true);
}
	
function Form1_Validator(theForm)
{
	<%For k=1 to Count1%>
		if ((theForm.add<%=k%>.value != "") && (theForm.add<%=k%>.value != "0"))
	  	{
			if (allDigit(theForm.add<%=k%>.value) == false)
			{
				alert("Please enter a valid number for this field.");
				theForm.add<%=k%>.focus();
			    return (false);
			}
			if (eval(theForm.add<%=k%>.value) < eval(theForm.remainqty<%=k%>.value))
			{
				alert("Please enter a valid number greater than or equal to " + eval(theForm.remainqty<%=k%>.value));
				theForm.add<%=k%>.focus();
			    return (false);
			}
		}
	<%Next%>

	return (true);
}
</script>
<!--#include file="footer_wrapper.asp"-->
