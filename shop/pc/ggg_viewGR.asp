<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="ggg_inc_chkEPPrices.asp"-->
<!--#include file="pcStartSession.asp"-->
<%
grCode=getUserInput(request("grCode"),0)
gOrder=getUserInput(request("gOrder"),0)
gSort=getUserInput(request("gSort"),0)
if gOrder="" then
	gOrder="products.Description"
end if
if gSort="" then
	gSort="ASC"
end if

mOrder=" ORDER by " & gOrder & " " & gSort

if grCode="" then
	response.redirect "msg.asp?message=98"
end if

query="select pcEv_IDEvent,pcEv_Name,pcEv_Date,pcEv_Type,pcEv_IncGcs from pcEvents where pcEv_Code='" & grCode & "' and pcEv_Active=1;"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rstemp.eof then
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

'Update Products Price
query="SELECT products.idproduct, pcEvProducts.pcEP_ID, pcEvProducts.pcEP_IDProduct, pcEvProducts.pcEP_OptionsArray, pcEvProducts.pcEP_IDConfig FROM products, pcEvProducts WHERE pcEvProducts.pcEP_IDEvent=" & gIDEvent & " AND products.idproduct=pcEvProducts.pcEP_IDProduct AND products.removed=0 AND ((products.active<>0) OR ((products.active=0) AND (products.pcProd_ParentPrd>0))) ORDER BY products.Description ASC, pcEvProducts.pcEP_GC ASC"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
do while not rstemp.eof
	geID=rstemp("pcEP_ID")
	gIDProduct=rstemp("pcEP_IDProduct")
	
	pcv_strOptionsArray=rstemp("pcEP_OptionsArray")

	gIDConfig=rstemp("pcEP_IDConfig")
	if gIDConfig<>"" then
	else
		gIDConfig="0"
	end if

	gnewPrice=updPrices(gIDProduct,gIDConfig,pcv_strOptionsArray,0)

	query="update pcEvProducts set pcEP_Price=" & gnewPrice & " where pcEP_ID=" & geID
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	set rs=nothing

	rstemp.MoveNext
loop

set rstemp=nothing

'End of Update Product Prices

%>
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<%query="SELECT products.sku,products.description,products.imageUrl,products.smallImageUrl,products.stock,products.nostock,products.ServiceSpec,products.pcProd_ParentPrd,pcEvProducts.pcEP_ID,pcEvProducts.pcEP_Qty,pcEvProducts.pcEP_HQty,pcEvProducts.pcEP_GC,pcEvProducts.pcEP_Price,pcEP_OptionsArray FROM products,pcEvProducts WHERE pcEvProducts.pcEP_IDEvent=" & gIDEvent & " AND products.idproduct=pcEvProducts.pcEP_IDProduct AND products.removed=0 AND ((products.active<>0) OR ((products.active=0) AND (products.pcProd_ParentPrd>0))) " & mOrder
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if%>
<div id="pcMain">
	<div class="pcMainContent">
		<form method="post" name="Form1" action="ggg_addEPtocart.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
		
			<h1>"<%=geName%>"<%= dictLanguage.Item(Session("language")&"_viewGR_1")%></h1>

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
					<div class="pcFormItemFull">
						<%= dictLanguage.Item(Session("language")&"_viewGR_11")%>
					</div>
				</div>

				<div class="pcSpacer"></div>

				<div class="pcFormButtons">
					<a class="pcButton pcButtonBack" href="javascript:history.go(-1);">
						<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
					</a>
				</div>
			<%ELSE%>
				<div class="pcSpacer"></div>

				<%
					col_SKUClass			= "pcCol-2 pcViewGR_SKU"
					col_NameClass			= "pcCol-4 pcViewGR_Name"
					col_PriceClass		= "pcCol-1 pcViewGR_Price"
					col_WantsClass		= "pcCol-1 pcViewGR_Wants"
					col_HasClass			= "pcCol-1 pcViewGR_Has"
					col_QuantityClass	= "pcCol-1 pcViewGR_Quantity"
				%>

				<div id="pcTableViewGR" class="pcTable">
					<div class="pcTableHeader">
						<div class="<%= col_SKUClass %>"><%= dictLanguage.Item(Session("language")&"_viewGR_2")%>&nbsp;<a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=products.sku&gSort=asc"><img src="<%=pcf_getImagePath("images","sortasc.gif")%>"></a><a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=products.sku&gSort=desc"><img src="<%=pcf_getImagePath("images","sortdesc.gif")%>"></a></div>
						<div class="<%= col_NameClass %>"><%= dictLanguage.Item(Session("language")&"_viewGR_3")%>&nbsp;<a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=products.description&gSort=asc"><img src="<%=pcf_getImagePath("images","sortasc.gif")%>"></a><a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=products.description&gSort=desc"><img src="<%=pcf_getImagePath("images","sortdesc.gif")%>"></a></div>
						<div class="<%= col_PriceClass %>"><%= dictLanguage.Item(Session("language")&"_viewGR_4")%>&nbsp;<a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_Price&gSort=asc"><img src="<%=pcf_getImagePath("images","sortasc.gif")%>"></a><a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_Price&gSort=desc"><img src="<%=pcf_getImagePath("images","sortdesc.gif")%>"></a></div>
						<div class="<%= col_WantsClass %>"><%= dictLanguage.Item(Session("language")&"_viewGR_5")%>&nbsp;<a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_Qty&gSort=asc"><img src="<%=pcf_getImagePath("images","sortasc.gif")%>"></a><a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_Qty&gSort=desc"><img src="<%=pcf_getImagePath("images","sortdesc.gif")%>"></a></div>
						<div class="<%= col_HasClass %>"><%= dictLanguage.Item(Session("language")&"_viewGR_6")%>&nbsp;<a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_HQty&gSort=asc"><img src="<%=pcf_getImagePath("images","sortasc.gif")%>"></a><a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_HQty&gSort=desc"><img src="<%=pcf_getImagePath("images","sortdesc.gif")%>"></a></div>
						<div class="<%= col_QuantityClass %>"><%= dictLanguage.Item(Session("language")&"_viewGR_12")%></div>
					</div>

					<div class="pcSpacer"></div>
					<%
					Count=0
					ExList="**"
					LowList="**"
					LowName="**"
					LowQStock="**"
					do while not rstemp.eof

						gparentPrdId=rstemp("pcProd_ParentPrd")
						gsku=rstemp("sku")
						gname=rstemp("description")

						'// Find product image
						gtimage1=rstemp("ImageUrl")
						gtimage2=rstemp("smallImageUrl")
						if gtimage2<>"" then
							gtimage=gtimage2
						else
							gtimage=gtimage1
						end if

						If statusAPP="1" Then
							
							query="SELECT imageUrl, smallImageUrl FROM products WHERE idproduct = " & gparentPrdId
							set rsTempImg=server.CreateObject("ADODB.RecordSet")
							set rsTempImg=connTemp.execute(query)
							if not rsTempImg.EOF then
								pcvParentImg=rsTempImg("imageUrl")
								pcvParentImgSm=rsTempImg("smallImageUrl")
							end if
							Set rsTempImg = Nothing

							if trim(pcvParentImgSm)="" then
								pcvParentImgSm=pcvParentImg
							end if
							if trim(gtimage)="" then
								gtimage=pcvParentImgSm
							end if

						End If
			
						if trim(gtimage)="" then
							gtimage="no_image.gif"
						end if

						gstock=rstemp("stock")
						if gstock<>"" then
						else
							gstock="0"
						end if
						gnostock=rstemp("nostock")
						if gnostock<>"" then
						else
							gnostock="0"
						end if
						gservice=rstemp("ServiceSpec")
						if gservice<>"" then
						else
							gservice="0"
						end if
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
						gPrice=rstemp("pcEP_Price")
						if gPrice<>"" then
						else
							gPrice="0"
						end if
		
						pcv_strSelectedOptions=""
						pcv_strSelectedOptions = rstemp("pcEP_OptionsArray")
						pcv_strSelectedOptions=pcv_strSelectedOptions&""
		
						if gGC<>"1" then
						Count=Count+1%>
						<div class="pcTableRow"> 
							<div class="<%= col_SKUClass %>"><%=gsku%></div>
								<div class="<%= col_NameClass %>">
									<% if gtimage<>"no_image.gif" then %>
										<img src="<%=pcf_getImagePath("catalog",gtimage)%>" alt="<%= gname %>">
									<% else %>
										<span class="spacer"></span>
									<% end if %>

									<a href="ggg_viewEP.asp?grCode=<%=grCode%>&geID=<%=geID%>"><%=gname%></a>

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
											<span class="small"><%= pOptionGroupDesc & ": " & pOptionDescrip%></span>
											<%
											if pPriceToAdd="" or pPriceToAdd="0" then
												response.write "&nbsp;"
											else 
												response.write " (" & scCurSign&money(pPriceToAdd) & ")"
											end if%>
											<br>
										<%end if
										set rsQ=nothing
									Next
		
									END IF%>
								</div>

								<div class="<%= col_PriceClass %>"><%=scCurSign & money(gPrice)%></div>
								<div class="<%= col_WantsClass %>"><%=gQty%></div>
								<div class="<%= col_HasClass %>"><%=clng(gHQty)%></div>
								<div class="<%= col_QuantityClass %>">
									<% if clng(gQty)-clng(gHQty)<=0 then
									ExList=ExList & Count & "**"%>
										<input name="add<%=Count%>" value="0" type=hidden>
										<% '// Fullfilled
										response.write dictLanguage.Item(Session("language")&"_viewGR_7")%>
									<% else
										if (clng(gQty)-clng(gHQty)<=clng(gstock)) or (gnostock<>"0") or ((gservice=-1) and (iBTOOutofStockPurchase=0)) or (scOutofstockpurchase=0) then %>
											<input name="add<%=Count%>" value="0" type=text size="3" style="float: right; text-align:right">
										<% else
											if clng(gstock)>0 then
											LowList=LowList & Count & "**"
											LowName=LowName & gname & "**"
											LowQStock=LowQStock & gstock & "**"%>
											<input name="add<%=Count%>" value="0" type=text size="3" style="float: right; text-align:right"><br>
											<% '// In Stock < Wanted quantity
											response.write dictLanguage.Item(Session("language")&"_viewPrd_19") & gstock%>
											<%else
												ExList=ExList & Count & "**"%>
												<input name="add<%=Count%>" value="0" type=hidden>
												<% '// Out of Stock 
												response.write dictLanguage.Item(Session("language")&"_viewGR_8")%>
											<% end if
										end if
									end if %>
									<input type=hidden name="geID<%=Count%>" value="<%=geID%>">
									<%if clng(gQty)-clng(gHQty)>clng(gstock) AND clng(gstock)>0 then%>
										<input name="remain<%=Count%>" value="<%=gstock%>" type=hidden>
									<%else%>
										<input name="remain<%=Count%>" value="<%if clng(gQty)-clng(gHQty)<0 then%>0<%else%><%=clng(gQty)-clng(gHQty)%><%end if%>" type=hidden>
									<%end if%>
								</div>
							</div>
							<%end if
							rstemp.MoveNext
						loop
						set rstemp=nothing
		
						Count1=Count
		
						query="select products.sku,products.description,products.smallImageUrl,products.stock,products.nostock,products.ServiceSpec,pcEvProducts.pcEP_ID,pcEvProducts.pcEP_Qty,pcEvProducts.pcEP_HQty,pcEvProducts.pcEP_GC,pcEvProducts.pcEP_Price from products,pcEvProducts where products.active=-1 and pcEvProducts.pcEP_IDEvent=" & gIDEvent & " and products.idproduct=pcEvProducts.pcEP_IDProduct and pcEvProducts.pcEP_GC=1 and products.removed=0 " & mOrder
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=connTemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rstemp=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
						IF NOT rstemp.eof then%>
						<div class="pcSpacer"></div>
						<div class="pcTableHeader">
							<div class="pcTableRowFull"><%= dictLanguage.Item(Session("language")&"_GRDetails_8")%></div>
						</div>
						<div class="pcSpacer"></div>
						<%do while not rstemp.eof
							gsku=rstemp("sku")
							gname=rstemp("description")
							gtimage=rstemp("smallImageUrl")
							if gtimage<>"" then
							else
								gtimage="no_image.gif"
							end if
							gstock=rstemp("stock")
							if gstock<>"" then
							else
								gstock="0"
							end if
							gnostock=rstemp("nostock")
							if gnostock<>"" then
							else
								gnostock="0"
							end if
							gservice=rstemp("ServiceSpec")
							if gservice<>"" then
							else
								gservice="0"
							end if
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
							gPrice=rstemp("pcEP_Price")
							if gPrice<>"" then
							else
								gPrice="0"
							end if
							if ((gGC="1") and (geincgc="1")) or (clng(gHQty)>0) then
							Count=Count+1%>
							<div class="pcTableRow"> 
								<div class="<%= col_SKUClass %>"><%=gsku%></div>
								<div class="<%= col_NameClass %>">
									<% if gtimage<>"no_image.gif" then %>
										<img src="<%=pcf_getImagePath("catalog",gtimage)%>">
									<% else %>
										<span class="spacer"></span>
									<% end if %>
									<a href="ggg_viewEP.asp?grCode=<%=grCode%>&geID=<%=geID%>"><%=gname%></a>
								</div>
								<div class="<%= col_PriceClass %>"><%=scCurSign & money(gPrice)%></div>
								<div class="<%= col_WantsClass %>">&nbsp;</div>
								<div class="<%= col_HasClass %>">&nbsp;</div>
								<div class="<%= col_QuantityClass %>">
									<%if (clng(gQty)-clng(gHQty)<=clng(gstock)) or (gnostock<>"0") OR (scOutofstockpurchase=0) then%>
										<input name="add<%=Count%>" type=text value="0" size="3" style="float: right; text-align:right">
									<%else%>
										<input name="add<%=Count%>" type=hidden value="0">
										<%= dictLanguage.Item(Session("language")&"_viewGR_8")%>
									<%end if%>
									<input type=hidden name="geID<%=Count%>" value="<%=geID%>">
									<input name="remain<%=Count%>" value="99999" type=hidden>
								</div>
							</div>
							<%end if
							rstemp.MoveNext
						loop
						set rstemp=nothing
					END IF 'Have GCs
					set rstemp=nothing%>

					<div class="pcSpacer"></div>

					<div class="pcFormButtons">
						<a class="pcButton pcButtonBack" href="javascript:history.go(-1);">
							<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
						</a>

						<button class="pcButton pcButtonAddToCart" id="submit" name="submit" value="<%= dictLanguage.Item(Session("language")&"_viewGR_9")%>">
							<img src="<%=pcf_getImagePath("",rslayout("addtocart"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_addtocart") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_addtocart") %></span>
						</button>
						
						<input type=hidden name="grCode" value="<%=grCode%>">
						<input type=hidden name="Count" value="<%=Count%>">
						<input type=hidden name="IDEvent" value="<%=gIDEvent%>">
					</div>
				<%END IF 'Have products%>
			</div>
		</form>
	</div>
</div>
<script type=text/javascript>

function isDigit(s)
{
var test=""+s;
if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
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
<%For k=1 to Count
if Instr(ExList,"**" & k & "**")=0 then%>
	if (theForm.add<%=k%>.value != "")
  	{
	if (allDigit(theForm.add<%=k%>.value) == false)
	{
		alert("Please enter a valid number for this field.");
		theForm.add<%=k%>.focus();
    return (false);
	}
	
	<%if Instr(LowList,"**" & k & "**")=0 then%>
	if (eval(theForm.add<%=k%>.value) > eval(theForm.remain<%=k%>.value))
	{
		alert("Your entered a quantity greater than remaining quantity.");
		theForm.add<%=k%>.focus();
    return (false);
	}
	<%else
		tmp1=split(LowList,"**")
		tmp2=split(LowName,"**")
		tmp3=split(LowQStock,"**")
		prdName=""
		prdStock=""
		For l=lbound(tmp1) to ubound(tmp1)
			if tmp1(l)<>"" then
				if clng(tmp1(l))=clng(k) then
					prdName=tmp2(l)
					prdStock=tmp3(l)
					exit for
				end if
			end if
		Next
		%>
		if (eval(theForm.add<%=k%>.value) > eval(theForm.remain<%=k%>.value))
		{
			alert("<%= dictLanguage.Item(Session("language")&"_instPrd_2")%><%=prdName%><%= dictLanguage.Item(Session("language")&"_instPrd_3")%><%=prdStock%><%= dictLanguage.Item(Session("language")&"_instPrd_4")%>");
			theForm.add<%=k%>.focus();
    		return (false);
		}
	<%end if%>
}
<%end if%>
<%Next%>

return (true);
}
</script>
<!--#include file="footer_wrapper.asp"-->