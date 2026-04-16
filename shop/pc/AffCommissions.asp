<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="AffLIv.asp"-->
<!--#include file="../includes/common.asp"-->

<!--#include file="header_wrapper.asp"-->
<div id="pcMain">
	<div class="pcMainContent">
		<h1><%=dictLanguage.Item(Session("language")&"_AffCom_1")%></h1>

		<div class="pcShowContent">
			<%
			' Load affiliate ID
			affVar=session("pc_idaffiliate")
			if not validNum(affVar) then
				response.redirect "AffiliateLogin.asp"
			end if

			Dim tempId
			tempId=0
					
			' Our Connection Object
			Dim con
			Set con=CreateObject("ADODB.Connection")
			con.Open scDSN 
		
			' Choose the records to display	
			query="SELECT * FROM Orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12))"
			query=query&" AND idaffiliate="& affVar &" ORDER BY orders.orderDate desc"
			' Our Recordset Object
	
			Set rs=CreateObject("ADODB.Recordset")
			rs.CursorLocation=adUseClient
			rs.Open query, scDSN , 3, 3
				
			' If the returning recordset is not empty
			If rs.EOF Then %>
				<div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_AffCom_2")%></div>
			<% Else						
				query="SELECT SUM(affiliatePay) AS AfftotalSum FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND idAffiliate=" & affVar
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=connTemp.execute(query)
				AffTotalSum=rstemp("AfftotalSum")
				if AffTotalSum<>"" then
				else
					AffTotalSum=0
				end if
					
				query="SELECT SUM(pcAffpay_Amount) AS AfftotalPaid FROM pcAffiliatesPayments WHERE pcAffpay_idAffiliate=" & affVar
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=connTemp.execute(query)
				AffTotalPaid=rstemp("AfftotalPaid")
				if AffTotalPaid<>"" then
				else
					AffTotalPaid=0
				end if
						
				CurrentBalance=AffTotalSum-AffTotalPaid
				%>
				<div class="pcFormItem">
					<div class="pcFormFull">
						<%=dictLanguage.Item(Session("language")&"_AffCom_3")%><%=scCursign%><%=money(CurrentBalance)%>
					</div>
				</div>

				<div class="pcFormItem">
					<div class="pcFormFull">
						<%=dictLanguage.Item(Session("language")&"_AffCom_4")%>
					</div>
				</div>

				<div id="AutoNumber1" class="pcTable">
					<div class="pcTableHeader">
						<div class="pcAffCommissions_PaymentDate"><%=dictLanguage.Item(Session("language")&"_AffCom_5")%></div>
						<div class="pcAffCommissions_PaymentAmount"><%=dictLanguage.Item(Session("language")&"_AffCom_6")%></div>
						<div class="pcAffCommissions_PaymentStatus"><%=dictLanguage.Item(Session("language")&"_AffCom_7")%></div>
					</div>
					
					<%query="SELECT pcAffpay_idpayment, pcAffpay_Amount, pcAffpay_PayDate, pcAffpay_Status FROM pcAffiliatesPayments WHERE pcAffpay_idAffiliate=" & affVar & " ORDER BY pcAffpay_PayDate DESC;"
					set rstemp=server.CreateObject("ADODB.RecordSet")
					set rstemp=connTemp.execute(query)
					if rstemp.eof then%>
						<div class="pcClear"></div>
						<div class="pcErrorMessage">
							<%=dictLanguage.Item(Session("language")&"_AffCom_8")%>
						</div>
					<%else
						do while not rstemp.eof
							IDpayment=rstemp("pcAffpay_idpayment")
							PaidAmount=rstemp("pcAffpay_Amount")
							PaidDate=rstemp("pcAffpay_PayDate")
							'// Format Date
							PaidDate=ShowDateFrmt(PaidDate)
							PaidStatus=rstemp("pcAffpay_Status")%>
							<div class="pcTableRow">
								<div class="pcAffCommissions_PaymentDate"><%=PaidDate%></div>
								<div class="pcAffCommissions_PaymentAmount"><%=scCurSign%>&nbsp;<%=money(PaidAmount)%></div>
								<div class="pcAffCommissions_PaymentStatus"><%=PaidStatus%></div>
							</div>
							<%rstemp.MoveNext
						loop%>
						<div class="pcSpacer"></div>
						<div class="pcTableRow">
							<div class="pcAffCommissions_PaymentDate"><%=dictLanguage.Item(Session("language")&"_AffCom_9")%></div>
							<div class="pcAffCommissions_PaymentAmount"><%=scCurSign%>&nbsp;<%=money(AffTotalPaid)%></div>
							<div class="pcAffCommissions_PaymentStatus">&nbsp;</div>
						</div>
					<%end if%>
				</div>
				<div class="pcClear"></div>
				<div class="pcFormItem">
					<div class="pcFormFull">
						<%=dictLanguage.Item(Session("language")&"_AffCom_10")%><%=rs.RecordCount%>
					</div>
				</div>

				<div class="pcSpacer"></div>

				<div class="pcTable">
					<div class="pcTableHeader">
						<div class="pcAffCommissions_Date"><%=dictLanguage.Item(Session("language")&"_AffCom_11")%></div>
						<div class="pcAffCommissions_OrderNumber"><%=dictLanguage.Item(Session("language")&"_AffCom_12")%></div>
						<div class="pcAffCommissions_TotalSales"><%=dictLanguage.Item(Session("language")&"_AffCom_15")%></div>
						<div class="pcAffCommissions_OrderTotal"><%=dictLanguage.Item(Session("language")&"_AffCom_16")%></div>
						<div class="pcAffCommissions_Shipping"><%=dictLanguage.Item(Session("language")&"_AffCom_17")%></div>
						<div class="pcAffCommissions_Tax"><%=dictLanguage.Item(Session("language")&"_AffCom_18")%></div>
						<div class="pcAffCommissions_Commission"><%=dictLanguage.Item(Session("language")&"_AffCom_13")%></div>
					</div>
					<% 
					gTotalsales=0
					gTotaltaxes=0
					gTotalOrder=0
					aTotalShip=0
					gTotalTax=0
					gTotalcomm=0
					do until rs.EOF
						gSubOrder=0
						gSubTax=0
						gSubShip=0
						gSubCom=0

						intIdOrder=rs("idOrder")
						intIdCustomer=rs("idcustomer")
						dtOrderDate=rs("orderDate")
						dblAffiliatePay=rs("affiliatePay")
						porderdetails=rs("details")
				
						'Calculate "NET" Order Amount
						ptotal=rs("total")
						gSubOrder=rs("total")
						ptaxAmount=rs("taxAmount")
						ptaxDetails=rs("taxDetails")
						pord_VAT=rs("ord_VAT")
						gSubTax=ptaxAmount+pord_VAT
						pshipmentDetails=rs("shipmentDetails")
						Postage=0
						serviceHandlingFee=0
						shipping=split(pshipmentDetails,",")
						if ubound(shipping)>1 then
							if NOT isNumeric(trim(shipping(2))) then
							else
								Postage=trim(shipping(2))
								if ubound(shipping)=>3 then
									serviceHandlingFee=trim(shipping(3))
									if NOT isNumeric(serviceHandlingFee) then
										serviceHandlingFee=0
									end if
								else
									serviceHandlingFee=0
								end if
							end if
						end if
						gSubShip=Postage
						ppaymentDetails=trim(rs("paymentDetails"))
						payment = split(ppaymentDetails,"||")
						PayCharge=0
						If ubound(payment)>=1 then					
							If payment(1)="" then
								PayCharge=0
							else
								PayCharge=payment(1)
							end If
						End if

						PrdSales=ptotal
						PrdSales=PrdSales-postage
						PrdSales=PrdSales-serviceHandlingFee
						PrdSales=PrdSales-PayCharge
							
						gSubOrder=gSubOrder-postage


						pdiscountDetails=rs("discountDetails")
						pcv_CatDiscounts=rs("pcOrd_CatDiscounts")
						if isNULL(pcv_CatDiscounts) OR pcv_CatDiscounts="" then
							pcv_CatDiscounts="0"
						end if
							
						if (instr(pdiscountDetails,"- ||")>0) or (pcv_CatDiscounts>"0")  then
							if instr(pdiscountDetails,",") then
								DiscountDetailsArry=split(pdiscountDetails,",")
								intArryCnt=ubound(DiscountDetailsArry)
							else
								intArryCnt=0
							end if
									
							dim discounts, discountType 
													
							discount=0
							for k=0 to intArryCnt
								if instr(pTempDiscountDetails,"- ||") then
									discounts = split(pTempDiscountDetails,"- ||")
									tdiscount = discounts(1)
								else
									tdiscount=0
								end if
								discount=discount+tdiscount
							Next
							PrdSales=PrdSales+discount+pcv_CatDiscounts
						end if
							
						if pord_VAT>0 then
							PrdSales=PrdSales-pord_VAT
						else
							if isNull(ptaxDetails) OR trim(ptaxDetails)="" then
								PrdSales=PrdSales-ptaxAmount
							else 
								taxArray=split(ptaxDetails,",")
								for i=0 to (ubound(taxArray)-1)
									taxDesc=split(taxArray(i),"|")
									PrdSales=PrdSales-taxDesc(1)
									gSubTax=gSubTax+taxDesc(1)
								next 
							end if
						end if

						gSubOrder=gSubOrder-gSubTax
							
						gTotalsales=gTotalsales + PrdSales
						gTotalOrder=gTotalOrder+gSubOrder
						gTotalShip=gTotalShip+gSubShip
						gTotalTax=gTotalTax+gSubTax
						%>
						<div class="pcTableRow"> 
							<% '// Format Date
							dtOrderDate=rs("orderDate")
							dtOrderDate=ShowDateFrmt(dtOrderDate) %>
							<div class="pcAffCommissions_Date"><%=dtOrderDate%></div>
							<div class="pcAffCommissions_OrderNumber">#<%=cdbl(scpre)+cdbl(rs("idOrder"))%></div>
							<div class="pcAffCommissions_TotalSales"><%=scCurSign&money(PrdSales)%></div>
							<div class="pcAffCommissions_OrderTotal"><%=scCurSign&money(gSubOrder)%></div>
							<div class="pcAffCommissions_Shipping"><%=scCurSign&money(gSubShip)%></div>
							<div class="pcAffCommissions_Tax"><%=scCurSign&money(gSubTax)%></div>
							<div class="pcAffCommissions_Commission"><%=scCurSign&money(rs("affiliatePay"))%></div>
						</div>
						<% gTotalcomm=gTotalcomm + rs("affiliatePay") %>
						<% rs.MoveNext
						loop
						End If %>
					</div>

					<div class="pcClear"></div>

					<div class="pcFormItem"><hr></div>

					<div class="pcFormItem">
						<div class="pcFormFull" style="text-align: right">
							<strong><%=dictLanguage.Item(Session("language")&"_AffCom_14")%></strong>
						</div>
					</div>

					<div class="pcFormItem">
						<div class="pcFormFull" style="text-align: right">
							<strong><%=scCurSign&money(gTotalcomm)%></strong>
						</div>
					</div>
			
					<div class="pcFormItem"><hr></div>

					<div class="pcFormButtons">
						<a class="pcButton pcButtonBack" href="javascript:history.back(-1);">
							<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
						</a>
					</div>
				</div>
	</div>
</div>
<%	' Done. Now release Objects
con.Close
Set con=Nothing
Set rs=Nothing
%>
<% call closedb() %>
<!--#include file="footer_wrapper.asp"-->
