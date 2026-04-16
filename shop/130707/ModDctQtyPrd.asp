<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Modify Quantity Discounts (Tiered Pricing)" %>
<% Section="specials" %>
<%PmAdmin=3%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<% 
Dim idproduct, discountdesc, percentage, baseproductonly, discountPerUnit1, discountPerWUnit1, quantityfrom1, 	nquantityfrom, nquantityUntil, ndiscountPerUnit

CanNotRun=0

pIDProduct=Request("idproduct")
if pIDProduct="" OR IsNull(pIDProduct) then
	pIDProduct=0
end if

'~~~~~~~~~~~~~~ Delete last tier only ~~~~~~~~~~~~~~~~~~~~~~~
if request("sAction")="D" then
	intId=request("Id")
	idproduct=request("idProduct")
	
	
	query="DELETE FROM discountsPerQuantity WHERE idDiscountPerQuantity="&intId&";"
	Set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query) 
	
	set rs=nothing
	
	
	call closeDb()
response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct
end if
'~~~~~~~~~~~~~~ Delete last tier only ~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~ DELETE ~~~~~~~~~~~~~~~~~~~~~~~
dMode=Request.QueryString("Delete")
if dMode<>"" then
	
	Session("adminidproduct")=Request.QueryString("idproduct")
	idproduct=Session("adminidproduct")
	Set rs=Server.CreateObject("ADODB.Recordset")
	query="DELETE FROM discountsPerQuantity WHERE idProduct="&idProduct 
	set rs=conntemp.execute(query)
	Set rs=nothing
	
	Session("adminidproduct")=""
	Session("admindiscountdesc")=""
	Session("admindiscountPerUnit1")=""
	Session("adminquantityfrom1")=""
	Session("adminidDiscountPerQuantity1")=""
	
	call closeDb()
response.redirect "modifyProduct.asp?idproduct="&idProduct
end if
'~~~~~~~~~~~~~~ DELETE ~~~~~~~~~~~~~~~~~~~~~~~

'// Check for conflict with Product Promotions

query="SELECT DISTINCT idproduct FROM pcPrdPromotions WHERE idproduct=" & pIDProduct & ";"
set rs=connTemp.execute(query)
if not rs.eof then
	CanNotRun=1%>
	<table class="pcCPcontent">
		   <tr>
				<td colspan="3">
					<div class="pcCPmessage">You cannot add quantity discounts to this product because it has a promotion assigned to it. <a href="ModPromotionPrd.asp?idproduct=<%=pIdProduct%>&iMode=start">Review the promotion</a>.</div>
				</td>
			</tr>
			<tr>
				<td>
					<input type="button" class="btn btn-default"  name="back" value=" Product Quantity Discounts " onclick="location='viewDisca.asp';">
					&nbsp;&nbsp;<input type="button" class="btn btn-default"  name="back2" value=" Back to Main menu " onclick="location='menu.asp';">
				</td>
			</tr>		
	</table>
	<%
end if
set rs=nothing

IF CanNotRun=0 THEN

'~~~~~~~~~~~~~~ UPDATE W/O ADDING ~~~~~~~~~~~~~~~~~~~~~~~
uMode=Request.Form("SubmitUPD")
If uMode<>"" Then
	
	ndiscountdesc=Request("discountdesc")
	npercentage=Request("percentage")
	nbaseproductonly=Request("baseproductonly")
	idproduct=Request("idproduct")
	query="UPDATE discountsPerQuantity SET percentage="&npercentage&", baseproductonly="&nbaseproductonly&" WHERE idProduct="&idProduct
	Set rs=server.CreateObject("ADODB.RecordSet")
	Set rs=conntemp.execute(query)
	Set rs=Nothing
	
	call closeDb()
response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct
End If

'~~~~~~~~~~~~~ ADD ~~~~~~~~~~~~~~~~~~
sMode=Request.Form("Submit")
If sMode<>"" Then
	iNextNum=request("iNextNum")
	iPrevNum=request("iPrevNum")
	idproduct=Request("idproduct")
	ndiscountdesc=Request("discountdesc")
	npercentage=Request("percentage")
	nbaseproductonly=Request("baseproductonly")	
	ndiscountPerUnit=replacecomma(Request("discountPerUnitAdd"&iNextNum))
	ndiscountPerWUnit=replacecomma(Request("discountPerWUnitAdd"&iNextNum))
	if ndiscountPerUnit="" then
		ndiscountPerUnit="0"
	end if
	if ndiscountPerWUnit="" then
		ndiscountPerWUnit="0"
	end if
	idr=request("ID"&iPrevNum)
	iPriority=int(iPrevNum)+1
	nquantityFrom=int(Request("quantityUntil"&idr))+1
	nquantityUntil=Request("quantityuntilAdd"&iNextNum)
	nidDiscountPerQuantity=Request("idDiscountPerQuantityAdd"&iNextNum)
	
	'check to make sure there are no overlaps
	if nquantityfrom = "" OR nquantityUntil = "" OR ndiscountPerUnit="" then
		msg="Both the 'To' and 'Retail Price/Percentage' fields are required."
		call closeDb()
response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct&"&msg="&msg
	end if
	'check to make sure there are no overlaps
	if nquantityfrom <> "" AND nquantityUntil <> "" AND ndiscountPerUnit="" then
		msg="You must specify a discount price for each tier."
		call closeDb()
response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct&"&msg="&msg
	end if
	'make sure the from < until
	if int(nquantityfrom)>int(nquantityUntil) then
		msg="Your quantity 'To' must be greater then the 'To' in the previous Tier."
		call closeDb()
response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct&"&msg="&msg
	end if
	
	If (money(ndiscountPerUnit) > 0 OR money(ndiscountPerWUnit)>0) AND nquantityfrom <> "" AND nquantityUntil <> "" AND nidDiscountPerQuantity="" Then
		
		query="INSERT INTO discountsPerQuantity (idproduct,idcategory,discountDesc,discountPerUnit,discountPerWUnit,quantityuntil,quantityfrom,num,percentage,baseproductonly) VALUES ("&idproduct&",0,'PD',"&ndiscountPerUnit&","&ndiscountPerWUnit&","&nquantityuntil&","&nquantityfrom&","&iPriority&","&npercentage&","&nbaseproductonly&");"
		Set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		Set rs=Nothing
		
	End If
	call closeDb()
response.redirect "ModDctQtyPrd.asp?idproduct="&idproduct
End If
'~~~~~~~~~~~~~~ ADD ~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~ SHOW ADMIN ~~~~~~~~~~~~~~~~~~~~~~~
idproduct=request("idproduct")

query="SELECT discountdesc,percentage,baseproductonly,num,discountPerUnit,discountPerWUnit,quantityfrom,quantityUntil,discountPerUnit,idDiscountPerQuantity  FROM discountsPerQuantity WHERE idproduct="&idproduct&" AND discountdesc='PD' ORDER BY num;"
Set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query) 

	if rs.eof then
		set rs=nothing
		
		call closeDb()
response.redirect "AdminDctQtyPrd.asp?idproduct="&idproduct
	end if

	discountdesc=rs("discountdesc")
	Session("adminpercentage")=rs("percentage")
	Session("adminbaseproductonly")=rs("baseproductonly")
	%>
	<form method="POST" action="ModDctQtyPrd.asp" class="pcForms">
		<table class="pcCPcontent">
			<tr> 
				<td colspan="5">
					<% 'get product info
					query="SELECT description,serviceSpec,sku,configonly,price,BtoBPrice FROM products WHERE idproduct="&idproduct
					set rsPrdObj=server.CreateObject("ADODB.RecordSet")
					set rsPrdObj=conntemp.execute(query)
					strDescription=rsPrdObj("description")
					pServiceSpec=rsPrdObj("serviceSpec")
					StrSKU=rsPrdObj("sku")
					configonly=rsPrdObj("configonly")
					pcv_dblProductPrice=Cdbl(rsPrdObj("price"))
					pcv_dblProductWPrice=Cdbl(rsPrdObj("btoBprice"))
					set rsPrdObj=nothing
					%>
                    
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>

					<h2><a href="FindProductType.asp?id=<%=idProduct%>"><%=strDescription%></a> - Sku: <%=strSKU%></h2>
						<div style="padding-top: 10px;">
							Online Price: <b><%=scCurSign%><%=money(pcv_dblProductPrice)%></b>
							<br>Wholesale Price: <b><%=scCurSign%><%=money(pcv_dblProductWPrice)%></b></br> 
						</div>
					<input type="hidden" name="discountdesc" value="PD">
					<input type="hidden" name="idproduct" value="<%=idproduct%>">
				</td>
			</tr>

			<tr> 
				<td colspan="5">
					<div style="padding: 10px; margin: 10px 0 10px 0; border: 1px solid #e1e1e1;">
						<div style="padding-bottom: 10px;">
							Discount based on:
							<% if Session("adminpercentage")="" then %>
								<input type="radio" name="percentage" value="0" class="clearBorder"><%=scCurSign%> 
								<input type="radio" name="percentage" value="-1" class="clearBorder">% 
							<% else %>
								<% if Session("adminpercentage")="0" then %>
									<input type="radio" name="percentage" value="0" checked class="clearBorder"><%=scCurSign%> 
									<input type="radio" name="percentage" value="-1" class="clearBorder">% 
								<%else %>
									<input type="radio" name="percentage" value="0" class="clearBorder"><%=scCurSign%> 
									<input type="radio" name="percentage" value="-1" checked class="clearBorder">% 
								<% end if %>
							<% end if %>
						</div>
							<% if pServiceSpec=True then %>
								<input type="hidden" name="baseproductonly" value="-1" checked class="clearBorder">
							<% else %>
								<div style="padding-bottom: 5px;">
										<% if Session("adminbaseproductonly")="-1" then %>
											<input type="radio" name="baseproductonly" value="-1" checked class="clearBorder">
										<% else 
												if Session("adminbaseproductonly")="" then %>
												<input type="radio" name="baseproductonly" value="-1" checked class="clearBorder">
											<% else %>
												<input type="radio" name="baseproductonly" value="-1" class="clearBorder">				
											<% end if %>		
										<% end if %>
										<%if configonly <> 0 then%>
											Apply discount to base price
										<%else%>
											Apply discount to base price only (product options not included)
										<%end if%>
								</div>
								<div style="padding-bottom: 10px;">
									<%if configonly<>true then%>					
										<% if Session("adminbaseproductonly")="0" then %>
											<input type="radio" name="baseproductonly" value="0" checked class="clearBorder">
										<% else %>
											<input type="radio" name="baseproductonly" value="0" class="clearBorder">
										<% end if %>
										Apply discount to base price + options prices (if any)
									<%end if%>
								</div>
							<% end if %>
							<input name="SubmitUPD" type="submit" id="SubmitUPD" value="Update" class="btn btn-primary">
						</div>
				</td>
			</tr>
			<tr>
				<td colspan="5"></td>
			</tr>
			<tr> 
				<td colspan="5">
					<table class="pcCPcontent">
						<tr>
							<th width="16%">Disc. Tiers</th>
							<th width="14%">From</th>
							<th width="15%">To</th>
							<th width="23%"><%=scCurSign%> or	% (retail)</th>
							<th width="27%" colspan="2"><%=scCurSign%> or	% (wholesale)</th>
						</tr>
					
						<%
						iDCnt=0
						do until rs.eof
							r=rs("num")
							discountPerUnit=rs("discountPerUnit")
							discountPerWUnit=rs("discountPerWUnit")
							quantityfrom=rs("quantityfrom")
							quantityUntil=rs("quantityUntil")
							discountPerUnit=rs("discountPerUnit")
							idDiscountPerQuantity=rs("idDiscountPerQuantity")
							%>
							<tr>
								<td>
									<%
									if iDCnt=0 then
										response.write "Quantity:"
									else 
										response.write "&nbsp;"
									end if 
									%>
								</td>
								<td>
									<%=quantityFrom%>
									<input type="hidden" name="idDiscountPerQuantity" value="<%=idDiscountPerQuantity%>">
									<input type="hidden" name="ID<%=r%>" value="<%=idDiscountPerQuantity%>">
									<input type="hidden" name="quantityFrom<%=idDiscountPerQuantity%>" value="<%=quantityFrom%>">
								</td>
								<td>
									<%=quantityUntil%>
									<input type="hidden" name="quantityUntil<%=idDiscountPerQuantity%>" value="<%=quantityUntil%>">
								</td>
								<td>
									<%=money(discountPerUnit)%> 
									<input type="hidden" name="discountPerUnit<%=idDiscountPerQuantity%>" value="<%=discountPerUnit%>">
								</td>
								<td>
									<%=money(discountPerWUnit)%>
									<input type="hidden" name="discountPerWUnit<%=idDiscountPerQuantity%>" value="<%=discountPerWUnit%>">
								</td>
								<td align="right"><a href="ModAllDctQtyPrd.asp?idProduct=<%=idProduct%>">Edit</a></td>
							</tr>
							<%
							iDCnt=iDCnt + 1
							rs.movenext
						loop
						Set rs=Nothing
						
						iPrevNum=r
						iNextNum=r+1
						%>
							<tr>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>
								<input name="quantityUntilAdd<%=iNextNum%>" type="text" size="6">
								<input type="hidden" name="iNextNum" value="<%=iNextNum%>">
								<input type="hidden" name="iPrevNum" value="<%=iPrevNum%>">
							</td>
							<td><input name="discountPerUnitAdd<%=iNextNum%>" type="text" size="6"></td>
							<td><input name="discountPerWUnitAdd<%=iNextNum%>" type="text" size="6"></td>
							<td align="right"><% if iDCnt>1 then %><a href="ModDctQtyPrd.asp?Id=<%=idDiscountPerQuantity%>&idProduct=<%=idProduct%>&sAction=D">Delete Last Tier</a><% end if %></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="5" class="pcCPspacer"></td>
			</tr>
			<tr> 
				<td colspan="5" align="center">
				<input type="submit" name="Submit" value="Add New Tier" class="btn btn-primary">&nbsp;
				<input type="button" class="btn btn-default"  name="Delete" value="Delete discount" onClick="javascript:if (confirm('You are about to permanantly delete this discount from the database. Are you sure you want to complete this action?')) location='moddctQtyPrd.asp?Delete=Yes&idproduct=<%=idProduct%>'">
			</td>
			</tr>
			<tr>
				<td colspan="5"><hr></td>
			</tr>
			<tr> 
				<td colspan="5" align="center">
				<%
				If statusAPP="1" OR scAPP=1 Then

					pcv_IDParent=request("idparent")
					if (request("idparent")="") or (request("idparent")="0") then
						call opendb()
						query="SELECT pcProd_ParentPrd FROM Products WHERE idproduct=" & idProduct & ";"
						set rs=connTemp.execute(query)
						if not rs.eof then
							pcv_IDParent=rs("pcProd_ParentPrd")
							if IsNull(pcv_IDParent) or pcv_IDParent="" then
								pcv_IDParent=0
							end if
						end if
						set rs=nothing
						call closedb()
					end if
								
					if (pcv_IDParent<>"0") then
						%>
						<input type="button" name="Apply" value="Apply to Other Sub-Products" onClick="location.href='app-ApplyDctToPrds.asp?idproduct=<%=idProduct%>&idparent=<%=pcv_IDParent%>'">
						
					<%
					end if

				End If
				%>
				<input type="button" class="btn btn-default"  name="Apply" value="Apply to Other Products" onClick="location.href='ApplyDctToPrds.asp?idproduct=<%=idProduct%>'">&nbsp;
				<% If statusAPP="1" OR scAPP=1 Then %>
					<input type="button" class="btn btn-default" value="Locate Another Product" onClick="location.href='<%if (pcv_IDParent<>"0") then%>app-viewDisca.asp?idparent=<%=pcv_IDParent%><%else%>viewDisca.asp<%end if%>'">
				<% Else %>
					<input type="button" class="btn btn-default"  value="Locate Another Product" onClick="location.href='viewDisca.asp'">
				<% End If %>
				</td>
			</tr>
			<tr>
				<td colspan="5" class="pcCPspacer"></td>
			</tr>
		</table>
	</form>
<%END IF 'CanNotRun%>
<!--#include file="AdminFooter.asp"-->