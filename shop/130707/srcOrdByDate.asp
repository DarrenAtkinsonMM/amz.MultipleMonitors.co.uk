<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=true %>
<% 
pageTitle="Sales Reports" 
pageIcon="pcv4_icon_sales.gif"
%>
<% Section="genRpts" %>
<%PmAdmin="8*10*"%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/javascripts/pcDateFunctions.js"-->
<% 
dim f, counter
dim strDateFormat
strDateFormat="mm/dd/yyyy"
if scDateFrmt="DD/MM/YY" then
	strDateFormat="dd/mm/yyyy"
end if
counter=0
%>
<script type=text/javascript>
	function Validate_Dates(theForm)
	{
	
		if (theForm.FromDate.value == "")
		{
			alert("Please enter From Date and try again.");
			theForm.FromDate.focus();
			return (false);
		}
		
		if (theForm.ToDate.value == "")
		{
			alert("Please enter To Date and try again.");
			theForm.ToDate.focus();
			return (false);
		}
		
		if (isDate(theForm.FromDate.value,"<%=strDateFormat%>","From Date")==false)
		{
			theForm.FromDate.focus()
			return false
		}
		
		if (isDate(theForm.ToDate.value,"<%=strDateFormat%>","To Date")==false)
		{
			theForm.ToDate.focus()
			return false
		}
		
		if (CompareDates(theForm.FromDate,theForm.ToDate,"From <= To")==false)
		{
			alert("From Date should be less than To Date.")
			theForm.ToDate.focus()
			return false
		}
	return (true);
	}
</script>
<%



%>
<!--#include file="pcCharts.asp"-->
<div class="panel-group" id="accordion">

    
        <div class="panel panel-default">        
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse1">
                        Quick Summary: Monthly product sales
                    </a>
                  </h4>
            </div>
            <div id="collapse1" class="panel-collapse collapse in">
                <div class="panel-body">
                
                    <p>
                    
                    <div style="background-image: url(images/pcv4_graphic_piechart.png); background-repeat: no-repeat; background-position: 0px 25px; padding-left: 210px;"></div>
            
                    <table width="100%" cellpadding="3" cellspacing="0">
                        <tr>
                            <td colspan="2">
                                <div id="chartMonthlySales" style="height:250px;"></div>
                                <%Dim pcv_YearTotal
                                pcv_YearTotal=0
                                call pcs_MonthlySalesChart("chartMonthlySales",Year(Date()),0,1)%>
                            </td>
                        </tr>
                    <%if pcv_YearTotal>0 then%>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <td align="left" valign="middle" nowrap>Year Total: <b><%=scCurSign & money(pcv_YearTotal)%></b></td>
                            <td width="100%">&nbsp;</td>
                        </tr>
                    <% end if 'Have sales data%>
                    </table>
                    
                    <%
                    ' count statistic registers
                    totalyear=0
            
                    query="SELECT year(orderdate) AS yearsql FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) GROUP BY year(orderdate) ORDER BY year(orderdate) DESC;" 
                    set rstemp=Server.CreateObject("ADODB.Recordset")
                    set rstemp=conntemp.execute(query)
            
                    stryear=""
                    do until rstemp.eof 
                        yearvalue=rstemp("yearsql")
                        if clng(yearvalue)<>clng(year(now())) then
                            stryear=stryear & yearvalue & "***"
                            totalyear=totalyear+1
                        end if   
                        rstemp.movenext
                    loop
                    set rstemp=nothing
                        if totalyear>0 then
                    %>
                        <div style="margin: 3px;">
                        Previous years: &nbsp;
                        <%
                        Ayear=split(stryear,"***")
                        For dd=1 to totalyear %>
                            <a href="#" onClick="chgWin('salescharts.asp?year=<%=Ayear(dd-1)%>','window2')"><%=Ayear(dd-1)%></a>
                        <% 
                        If dd <> totalyear Then Response.Write " - " End if
                        Next
                        %>
                        </div>
                    <%
                        end if
                    %>
                
                    </p>
                    
		        </div>
            </div>
        </div>



        <div class="panel panel-default">
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse2">
                        Sales by Date
                    </a>
                  </h4>
            </div>
            <div id="collapse2" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>

                    <div style="float: right; margin: 10px; width: 200px;">Specify a date range to view all sales in that period. <b>Note</b>: You must enter both dates in the format <%=strDateFormat%></div>
                    <form action="salesReport.asp" name="date_form" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
                    <% todayDate=Date()
                    Dim varMonth, varDay, varYear
                    varMonth=Right("0" & DatePart("m",Date), 2)
                    varDay=Right("0" & DatePart("d",Date), 2)
                    varYear=DatePart("yyyy", Date) 
                    dim dtInputStrStart, dtInputStr
                    dtInputStrStart=(varMonth&"/01/"&varYear)
                    if scDateFrmt="DD/MM/YY" then
                        dtInputStrStart=("01/"&varMonth&"/"&varYear)
                    end if
                    dtInputStr=(varMonth&"/"&varDay&"/"&varYear)
                    if scDateFrmt="DD/MM/YY" then
                        dtInputStr=(varDay&"/"&varMonth&"/"&varYear)
                    end if
                    %>
                    From: <input name="FromDate" type="text" class="datepicker" autocomplete="off" value="<%=dtInputStrStart%>" size="10">&nbsp;
                    To: <input class="datepicker" autocomplete="off" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">&nbsp;
                    Based on: 
                    <select name="basedon">
                        <option value="1" selected>Ordered Date</option>
                        <option value="2">Processed Date</option>
                        <option value="3">Shipped Date</option>
                    </select>
                    <br /><br />
                    Customer Type:&nbsp;
                    <select name="customerType">
                        <option value="" selected>All</option>
                            <option value='0'>Retail Customer</option>
                            <option value='1'>Wholesale Customer</option>
                            <% 'START CT ADD %>
                            <% 'if there are PBP customer type categories - List them here 
                            query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories;"
                            SET rs=Server.CreateObject("ADODB.RecordSet")
                            SET rs=conntemp.execute(query)
                            if NOT rs.eof then 
                                do until rs.eof 
                                    intIdcustomerCategory=rs("idcustomerCategory")
                                    strpcCC_Name=rs("pcCC_Name")
                                    %>
                                    <option value='CC_<%=intIdcustomerCategory%>'
                                    <%if Session("pcAdmincustomertype")="CC_"&intIdcustomerCategory then 
                                        response.write "selected"
                                    end if%>
                                    ><%=strpcCC_Name%></option>
                                    <% 
                                rs.moveNext
                                loop
                            end if
                            SET rs=nothing
                            'END CT ADD %>
                        </select>
                    <br /><br />
                    Country:
                    <%
                    query="SELECT CountryCode,countryName FROM countries ORDER BY countryName ASC"
                    set rstemp=Server.CreateObject("ADODB.Recordset")
                    set rstemp=conntemp.execute(query)
                    if err.number <> 0 then
                    set rstemp=nothing
                    
                    call closeDb()
response.redirect "techErr.asp?error="&Server.Urlencode("Error in order: "&err.description)
                    end If
                    %>
                    <select name="CountryCode">
                    <option value="" selected>-- All countries --</option>
                    <%
                    do while not rstemp.eof
                        pCountryCode2=rstemp("CountryCode")%>
                        <option value="<%response.write pCountryCode2%>" <%if pCountryCode2=scShipFromPostalCountry then%>selected<%end if %>><%response.write rstemp("countryName")%></option>
                    <%
                    rstemp.movenext
                    loop
                    set rstemp = nothing
                    %>
                    </select>
                    <br /><br />
                    <%
                    query="SELECT pcSales_ID, pcSales_Name FROM pcSales ORDER BY pcSales_Name ASC;"
                    set rs=connTemp.execute(query)
                    IF NOT rs.eof THEN
                        tmpArr=rs.getRows()
                        set rs=nothing
                        intCount=ubound(tmpArr,2)%>
                        Sale Name: <select name="saleID">
                        <option value="" selected>-- All Sales --</option>
                        <%For i=0 to intCount%>
                        <option value="<%=tmpArr(0,i)%>"><%=tmpArr(1,i)%></option>
                        <%Next%>
                        </select>
                        <br /><br />
                    <%END IF
                    set rs=nothing%>
                    <input type="radio" name="onlyShow" value="All" class="clearBorder" checked> Show all orders within the above date range.
                    <br>
                    <input type="radio" name="onlyShow" value="onlyDisc" class="clearBorder"> Only show orders for which a discount code was used
                    <br>
                    <input type="radio" name="onlyShow" value="onlyGC" class="clearBorder"> Only show orders for which a gift certificate code was used
                    <br /><br />
                    <input type="submit" value="Search" name="submit" class="btn btn-primary">
                    </form>
                    
                    </p>
                    
		        </div>
            </div>
        </div>



        <div class="panel panel-default">
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse3">
                        Sales by Product
                    </a>
                  </h4>
            </div>
            <div id="collapse3" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                    <div style="float: right; margin: 10px; width: 200px;">Specify a product and date range to view all sales in that period. Only products for which sales have been recorded are shown in the drop-down. <b>Note</b>: You must enter both dates in the format <strong><%=strDateFormat%>.</strong></div>
                
                    <form action="PrdsalesReport.asp" name="prdsales_form" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
                    <% todayDate=Date() %>
                    From:	<input name="FromDate" type="text" class="datepicker" autocomplete="off" value="<%=dtInputStrStart%>" size="10">&nbsp;
                    To: <input class="datepicker" autocomplete="off" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">&nbsp;
                    Based on:&nbsp;<select name="basedon">
                    <option value="1" selected>Ordered Date</option>
                    <option value="2">Processed Date</option>
                    <option value="3">Shipped Date</option>
                    </select>
                    <br /><br />
                    Select a product (only products that have been sold are shown):
                    <br />
                    <%
                    query="SELECT idproduct,description,sku FROM products WHERE removed=0 AND active=-1 AND sales>0 ORDER BY description ASC"
                    set rstemp=Server.CreateObject("ADODB.Recordset")
                    set rstemp=conntemp.execute(query)
                    if err.number <> 0 then
                        set rstemp=nothing
                        
                        call closeDb()
response.redirect "techErr.asp?error="&Server.Urlencode("Error in order: "&err.description)
                    end If
        
                    intCount=CInt(-1)
                    if not rstemp.eof then
                        prdArray = rstemp.getRows()
                        if ubound(prdArray,2) <> "" then
                            intCount=ubound(prdArray,2)
                        end if
                    end if
                    set rstemp = nothing
        
                    %>
        
                    <select name="IDProduct">
                    <option value="" selected>-- All products --</option>
                    <% for i=0 to intCount%>
                        <option value="<%=prdArray(0,i)%>"><%=prdArray(1,i)%> (<%=prdArray(2,i)%>)</option>
                    <% next %>
                    </select>
                    <br /><br />
                    <%
                    query="SELECT pcSales_ID, pcSales_Name FROM pcSales ORDER BY pcSales_Name ASC;"
                    set rs=connTemp.execute(query)
                    IF NOT rs.eof THEN
                        tmpArr=rs.getRows()
                        set rs=nothing
                        intCount=ubound(tmpArr,2)%>
                        Sale Name: <select name="saleID">
                        <option value="" selected>-- All Sales --</option>
                        <%For i=0 to intCount%>
                        <option value="<%=tmpArr(0,i)%>"><%=tmpArr(1,i)%></option>
                        <%Next%>
                        </select>
                        <br /><br />
                    <%END IF
                    set rs=nothing%>
                    <input type="submit" value="Search" name="submit" class="btn btn-primary">
                    </form>
            
                    </p>
                    
		        </div>
            </div>        
        </div>




        <div class="panel panel-default">
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse4">
                        Sales by Affiliate
                    </a>
                  </h4>
            </div>
            <div id="collapse4" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                    <div style="float: right; margin: 10px; width: 200px;">Specify a date range to view all sales by affiliate in that period. <b>Note</b>: You must enter both dates in the format <strong><%=strDateFormat%>.</strong> You then must select the affiliate, you can insert an ID, or choose from the drop-down list.</div>
        
                    <form action="salesReportAffiliate.asp" name="aff_form" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
                    From:	<input class="datepicker" autocomplete="off" type="text" size="10" name="FromDate" value="<%=dtInputStrStart%>">&nbsp;
                    To:	<input class="datepicker" autocomplete="off" type="text" size="10" name="ToDate" value="<%=dtInputStr%>">&nbsp;
                    Based on:&nbsp;<select name="basedon">
                        <option value="1" selected>Ordered Date</option>
                        <option value="2">Processed Date</option>
                        <option value="3">Shipped Date</option>
                        </select>
                    <br /><br />	
                    ID: 
                    <input type="text" size="5" maxlength="100" name="idaffiliate1">
                    <b>OR </b> 
                    Name: 
                    <%
                    query="SELECT idAffiliate, affiliateName FROM affiliates WHERE idaffiliate>1 ORDER BY affiliateName"
                    set rsAffObj=Server.CreateObject("ADODB.Recordset")
                    set rsAffObj=conntemp.execute(query)
                    if err.number <> 0 then
                        response.end
                        set rsAffObj = nothing
                        
                        call closeDb()
response.redirect "techErr.asp?error="&Server.Urlencode("Error in order: "&err.description)
                    end If
                    %>
                    <select name="idaffiliate2">
                        <option value="0">Select Affiliate</option>
                        <option value="ALL">Show All</option>
                        <% if not rsAffObj.eof then
                                do until rsAffObj.eof %>
                                <option value="<%=rsAffObj("idAffiliate")%>"><%=rsAffObj("affiliateName")%></option>
                        <%
                                    rsAffObj.moveNext
                                loop 
                                End If
                        set rsAffObj = nothing
                        %>
                    </select>
                    <br /><br />
                    <input type="submit" value="Search" name="submit" class="btn btn-primary">
                    </form>
                    
                    </p>
            
		        </div>
            </div>
        </div>




        <div class="panel panel-default">    
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse5">
                        Sales by Discount Code (Electronic Coupon)
                    </a>
                  </h4>
            </div>
            <div id="collapse5" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                    <div style="float: right; margin: 10px; width: 200px;">Specify a discount code and date range to view all sales in that period. <b>Note</b>: You must enter both dates in the format <%=strDateFormat%></div>
                    
                    <form action="DiscsalesReport.asp" name="discsales_form" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
                    <% todayDate=Date() %>
                    From: <input name="FromDate" type="text" class="datepicker" autocomplete="off" value="<%=dtInputStrStart%>" size="10">&nbsp;
                    To: <input class="datepicker" autocomplete="off" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">&nbsp;
                    Based on:&nbsp;<select name="basedon">
                    <option value="1" selected>Ordered Date</option>
                    <option value="2">Processed Date</option>
                    <option value="3">Shipped Date</option>
                    </select>
                    <br /><br />
                    Discount Code:
                    <%
        '			err.clear()
                    query="SELECT iddiscount,discountdesc,discountcode FROM discounts ORDER BY discountdesc asc"
                    set rstemp=Server.CreateObject("ADODB.Recordset")
                    set rstemp=conntemp.execute(query)
                    if err.number <> 0 then
                        set rstemp = nothing
                        
                        call closeDb()
response.redirect "techErr.asp?error="&Server.Urlencode("Error in order: "&err.description)
                    end If
                    %>
                    <select name="IDDiscount">
                    <option value="" selected>-- All --</option>
                    <%do while not rstemp.eof
                    %>
                    <option value="<%=rstemp("IDdiscount")%>"><%=rstemp("discountdesc") & " (" & rstemp("discountcode") & ")"%></option>
                    <%rstemp.movenext
                    loop
                    set rstemp = nothing
                    %>
                    </select>
                    <br />
                    <br />
                    <input type="submit" value="Search" name="submit" class="btn btn-primary">
                    </form>
                    
                    </p>
            
		        </div>
            </div>
        </div>



        <div class="panel panel-default">
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse6">
                        Sales by Pricing Category
                    </a>
                  </h4>
            </div>
            <div id="collapse6" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                    <div style="float: right; margin: 10px; width: 200px;">Specify a discount code and date range to view all sales in that period. <b>Note</b>: You must enter both dates in the format <%=strDateFormat%></div>
                    
                    <form action="PricingCatReport.asp" name="pcsales_form" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
                    <% todayDate=Date() %>
                    From: <input name="FromDate" type="text" class="datepicker" autocomplete="off" value="<%=dtInputStrStart%>" size="10">&nbsp;
                    To: <input class="datepicker" autocomplete="off" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">&nbsp;
                    Based on:&nbsp;<select name="basedon">
                    <option value="1" selected>Ordered Date</option>
                    <option value="2">Processed Date</option>
                    <option value="3">Shipped Date</option>
                    </select>
                    <br /><br />
                    Customer Type:&nbsp;
                    <select name="customerType">
                        <option value="" selected>All</option>
                            <option value='0'>Retail Customer</option>
                            <option value='1'>Wholesale Customer</option>
                            <% 'START CT ADD %>
                            <% 'if there are PBP customer type categories - List them here 
                            query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories;"
                            SET rs=Server.CreateObject("ADODB.RecordSet")
                            SET rs=conntemp.execute(query)
                            if NOT rs.eof then 
                                do until rs.eof 
                                    intIdcustomerCategory=rs("idcustomerCategory")
                                    strpcCC_Name=rs("pcCC_Name")
                                    %>
                                    <option value='CC_<%=intIdcustomerCategory%>'
                                    <%if Session("pcAdmincustomertype")="CC_"&intIdcustomerCategory then 
                                        response.write "selected"
                                    end if%>
                                    ><%=strpcCC_Name%></option>
                                    <% rs.moveNext
                                loop
                            end if
                            SET rs=nothing
                            'END CT ADD %>
                        </select>
                    <br />
                    <br />
                    <input type="submit" value="Search" name="submit" class="btn btn-primary">
                    </form>
                
                    </p>
                    
		        </div>
            </div>
        </div>



        <div class="panel panel-default">
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse7">
                        Sales by Referrer
                    </a>
                  </h4>
            </div>
            <div id="collapse7" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                    <div style="float: right; margin: 10px; width: 200px;">Select a referrer and a date range to view all sales in that period. <b>Note</b>: You must enter both dates in the format <strong><%=strDateFormat%>.</strong></div>
                    
                    <form action="RefsalesReport.asp" name="Refsales_form" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
                    <% todayDate=Date() %>
                    From:	<input name="FromDate" type="text" class="datepicker" autocomplete="off" value="<%=dtInputStrStart%>" size="10">&nbsp;
                    To: <input class="datepicker" autocomplete="off" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">&nbsp;
                    Based on:&nbsp;<select name="basedon">
                    <option value="1" selected>Ordered Date</option>
                    <option value="2">Processed Date</option>
                    <option value="3">Shipped Date</option>
                    </select>
                    <br /><br />
                    Referrer:
                    <% 		
                    queryStrRef="SELECT Name,IdRefer FROM Referrer ORDER BY Name"
                    Set rsCustRef=CreateObject("ADODB.Recordset")
                    rsCustRef.CursorLocation=adUseClient
                    rsCustRef.Open queryStrRef, scDSN , 3, 3
                    if rsCustRef.EOF Then
                        Response.Write(" No referrers have been setup. Referrers are setup in the <a href=""checkoutOptions.asp#referrer"">Checkout Options</a> area.")
                    else %>
                    <select name="IdRefer">
                    <%do while not rsCustRef.eof%>
                    <option value="<%=rsCustRef("IdRefer")%>"><%=rsCustRef("Name")%></option>
                    <%rsCustRef.movenext
                    loop
                    rsCustRef.Close
                    set rsCustRef=nothing %>
                    </select>
                    &nbsp;<span class="pcSmallText">Referrers are setup in the <a href="checkoutOptions.asp#referrer">Checkout Options</a> area.</span>
                    <br /><br />
                    <input type="submit" value="Search" name="submit" class="btn btn-primary">
                    <% end if %>
                    </form>
                
                    </p>
                    
		        </div>
            </div>
        </div>




        <div class="panel panel-default">
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse8">
                        Sales by Payment Type
                    </a>
                  </h4>
            </div>
            <div id="collapse8" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                        <div style="float: right; margin: 10px; width: 200px;">Specify a date range to view all sales by payment type in that period. <b>Note</b>: 
                    You must enter both dates in the format <strong><%=strDateFormat%>.</strong></div>
                        
                        <form action="salesReportPayment.asp" name="payment_form" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
                        From: <input name="FromDate" type="text" class="datepicker" autocomplete="off" value="<%=dtInputStrStart%>" size="10">&nbsp;
                        To: <input class="datepicker" autocomplete="off" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">&nbsp;
                        Based on:&nbsp;<select name="basedon">
                        <option value="1" selected>Ordered Date</option>
                        <option value="2">Processed Date</option>
                        <option value="3">Shipped Date</option>
                        </select>
                        <br /><br />
                        <%
                        query="SELECT DISTINCT (paymentDesc), idPayment FROM payTypes ORDER BY paymentDesc ASC"
                        Set rs=Server.CreateObject("ADODB.Recordset")
                        set rs=conntemp.execute(query)
                        tmp1=0%>
                                
                        Payment Type:&nbsp;
                        <select class="select" name="PayType" size="1">
                            <% Do While Not rs.EOF
                                strPaymentDesc=rs("paymentDesc")
                                intIdPayment=rs("idPayment") %>
                                <option value="<%=strPaymentDesc%>" <%if tmp1=0 then%>selected<%tmp1=1%><%end if%>><%=strPaymentDesc %></option>
                                <% rs.movenext					
                            loop %>
                            <%
                            Set rs=Nothing
                            %>
                        </select>
                        <br /><br />
                        <input type="submit" value="Search" name="submit" class="btn btn-primary">
                        </form>
                
                    </p>
                    
		        </div>
            </div>
        </div>




        <div class="panel panel-default">
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse9">
                        Sales by Brand
                    </a>
                  </h4>
            </div>
            <div id="collapse9" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                        <div style="float: right; margin: 10px; width: 200px;">Specify a brand and date range to view all sales in that period. <b>Note</b>: You must enter both dates in the format <%=strDateFormat%></div>
                        
                        <form action="BrandSalesReport.asp" name="brandsales_form" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
                        <% todayDate=Date() %>
                        <p><b>View Sales by Brand</b>
                        <br /><br />
                        From: <input name="FromDate" type="text" class="datepicker" autocomplete="off" value="<%=dtInputStrStart%>" size="10">&nbsp;
                        To: <input class="datepicker" autocomplete="off" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">&nbsp;
                        Based on:&nbsp;<select name="basedon">
                        <option value="1" selected>Ordered Date</option>
                        <option value="2">Processed Date</option>
                        <option value="3">Shipped Date</option>
                        </select>
                        <br /><br />
                        Brand Name:
                        <%
                        query="SELECT DISTINCT brands.idbrand,brands.BrandName FROM Brands INNER JOIN Products ON Brands.IDBrand=Products.IDBrand WHERE Products.active<>0 AND Products.removed=0 AND Brands.pcBrands_Active=1 ORDER BY Brands.BrandName ASC;"
                        set rstemp=Server.CreateObject("ADODB.Recordset")
                        set rstemp=conntemp.execute(query)
                        if err.number <> 0 then
                            set rstemp = nothing
                            
                            call closeDb()
response.redirect "techErr.asp?error="&Server.Urlencode("Error in order: "&err.description)
                        end If
                        if rstemp.eof then
                        %>
                        No brands containing products found. <a href="BrandsManage.asp">Manage brands</a>.
                        <%
                        else
                        %>
                            <select name="IDBrand">
                            <%do while not rstemp.eof
                            %>
                            <option value="<%=rstemp("IDBrand")%>"><%=rstemp("BrandName")%></option>
                            <%rstemp.movenext
                            loop
                            set rstemp = nothing
                            %>
                            </select>
                            <br />
                            <br />
                            <input type="submit" value="Search" name="submit" class="btn btn-primary">
                         <%
                         end if
                         set rstemp=nothing
                         %>
                            </p>
                        </form>
                
                    </p>
                    
		        </div>
            </div>
        </div>




        <div class="panel panel-default">

            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse10">
                        Sales by Category
                    </a>
                  </h4>
            </div>
            <div id="collapse10" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                        <div style="float: right; margin: 10px; width: 200px;">Specify a category and date range to view all sales in that period. <b>Note</b>: You must enter both dates in the format <%=strDateFormat%></div>
                        
                        <form action="CatSalesReport.asp" name="catsales_form" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
                        <% todayDate=Date() %>
                        From: <input name="FromDate" type="text" class="datepicker" autocomplete="off" value="<%=dtInputStrStart%>" size="10">&nbsp;
                        To: <input class="datepicker" autocomplete="off" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">&nbsp;
                        Based on:&nbsp;<select name="basedon">
                        <option value="1" selected>Ordered Date</option>
                        <option value="2">Processed Date</option>
                        <option value="3">Shipped Date</option>
                        </select>
                        <br /><br />
                        Category Name:
                            <%cat_DropDownName="idcategory"
                            cat_Type="1"
                            cat_DropDownSize="1"
                            cat_MultiSelect="0"
                            cat_ExcBTOHide="0"
                            cat_StoreFront="0"
                            cat_ShowParent="1"
                            cat_DefaultItem="Select a category"
                            cat_SelectedItems="0,"
                            cat_ExcItems=""
                            cat_ExcSubs="0"
                            cat_ExcBTOItems="1"
                            cat_EventAction=""
                            %>
                            
                            <%call pcs_CatList()%>
                        <br />
                        <br />
                        <input type="submit" value="Search" name="submit" class="btn btn-primary">
                        </form>
                
                    </p>
                    
		        </div>
            </div>
        </div>



        <!--
        <%      
        '// Suppliers have been removed in v5
        %>
        <div class="panel panel-default">
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse11">
                        Sales by Supplier
                    </a>
                  </h4>
            </div>
            <div id="collapse11" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                        <div style="float: right; margin: 10px; width: 200px;">Specify a supplier and date range to view all sales in that period. <b>Note</b>: You must enter both dates in the format <%=strDateFormat%></div>
                        
                        <form action="SupplierSalesReport.asp" name="suppsales_form" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
                        <% todayDate=Date() %>
                        <p><b>View Sales by Supplier</b>
                        <br /><br />
                        From: <input name="FromDate" type="text" class="datepicker" autocomplete="off" value="<%=dtInputStrStart%>" size="10">&nbsp;
                        To: <input class="datepicker" autocomplete="off" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">&nbsp;
                        Based on:&nbsp;<select name="basedon">
                        <option value="1" selected>Ordered Date</option>
                        <option value="2">Processed Date</option>
                        <option value="3">Shipped Date</option>
                        </select>
                        <br /><br />
                        Supplier Name:
                        <%
            '			err.clear()
                        query="SELECT DISTINCT pcSuppliers.pcSupplier_ID,pcSuppliers.pcSupplier_FirstName,pcSuppliers.pcSupplier_Lastname,pcSuppliers.pcSupplier_Company FROM pcSuppliers INNER JOIN Products ON pcSuppliers.pcSupplier_ID=Products.pcSupplier_ID WHERE Products.active<>0 AND Products.removed=0 ORDER BY pcSuppliers.pcSupplier_Company ASC;"
                        set rstemp=Server.CreateObject("ADODB.Recordset")
                        set rstemp=conntemp.execute(query)
                        if err.number <> 0 then
                            set rstemp = nothing
                            
                            call closeDb()
response.redirect "techErr.asp?error="&Server.Urlencode("Error in order: "&err.description)
                        end If
                        
                        if rstemp.eof then
                        %>
                        No suppliers found. 
                        <%
                        else
                        %>
                            <select name="IDSupplier">
                            <%do while not rstemp.eof
                            %>
                            <option value="<%=rstemp("pcSupplier_ID")%>"><%=rstemp("pcSupplier_Company") & " (" & rstemp("pcSupplier_FirstName") & " " & rstemp("pcSupplier_Lastname") & ")"%></option>
                            <%rstemp.movenext
                            loop
                            set rstemp = nothing
                            %>
                            </select>
                            <br />
                            <br />
                            <input type="submit" value="Search" name="submit" class="btn btn-primary">
                        <%
                        end if
                        set rstemp=nothing
                        %>
                        </form>
                
                    </p>
                    
		        </div>
            </div>
        </div>
        -->




        <div class="panel panel-default">
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse12">
                        Sales by Drop-Shipper
                    </a>
                  </h4>
            </div>
            <div id="collapse12" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                        <div style="float: right; margin: 10px; width: 200px;">Specify a drop-shipper and date range to view all sales in that period. <b>Note</b>: You must enter both dates in the format <%=strDateFormat%></div>
                        
                        <form action="DShipperSalesReport.asp" name="dshipsales_form" target="_blank" class="pcForms" onSubmit="return Validate_Dates(this)">
                        <% todayDate=Date() %>
                        From: <input name="FromDate" type="text" class="datepicker" autocomplete="off" value="<%=dtInputStrStart%>" size="10">&nbsp;
                        To: <input class="datepicker" autocomplete="off" type="text" size="10" value="<%=dtInputStr%>" name="ToDate">&nbsp;
                        Based on:&nbsp;<select name="basedon">
                        <option value="1" selected>Ordered Date</option>
                        <option value="2">Processed Date</option>
                        <option value="3">Shipped Date</option>
                        </select>
                        <br /><br />
                        Drop-Shipper Name:
                        <%
            '			err.clear()
                        query="SELECT DISTINCT pcDropShippers.pcDropShipper_ID,pcDropShippers.pcDropShipper_FirstName,pcDropShippers.pcDropShipper_Lastname,pcDropShippers.pcDropShipper_Company FROM pcDropShippers INNER JOIN Products ON pcDropShippers.pcDropShipper_ID=Products.pcDropShipper_ID WHERE Products.active<>0 AND Products.removed=0 ORDER BY pcDropShippers.pcDropShipper_Company ASC;"
                        set rstemp=Server.CreateObject("ADODB.Recordset")
                        set rstemp=conntemp.execute(query)
                        if err.number <> 0 then
                            set rstemp = nothing
                            
                            call closeDb()
response.redirect "techErr.asp?error="&Server.Urlencode("Error in order: "&err.description)
                        end If
                        
                        if rstemp.eof then
                        %>
                        No drop-shippers found. <a href="sds_addnew.asp?pagetype=1">Click here</a> to add Drop-shippers.
                        <%
                        else
                        %>
                            <select name="IDDropShipper">
                            <%do while not rstemp.eof
                            %>
                            <option value="<%=rstemp("pcDropShipper_ID")%>"><%=rstemp("pcDropShipper_Company") & " (" & rstemp("pcDropShipper_FirstName") & " " & rstemp("pcDropShipper_Lastname") & ")"%></option>
                            <%rstemp.movenext
                            loop
                            set rstemp = nothing
                            %>
                            </select>
                            <br />
                            <br />
                            <input type="submit" value="Search" name="submit" class="btn btn-primary">
                        <%
                        end if
                        set rstemp=nothing
                        %>
                        </form>
                
                    </p>
                    
                </div>
            </div>
        </div>





        <%
        query="SELECT pcSales_ID, pcSales_Name FROM pcSales ORDER BY pcSales_Name ASC;"
        set rs=connTemp.execute(query)
        IF NOT rs.eof THEN 
        %>			
        <div class="panel panel-default">

            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse13">
                        Sale Summary Report
                    </a>
                  </h4>
            </div>
            <div id="collapse13" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                        <div style="float: right; margin: 10px; width: 200px;">Use this form to see sale summary report.</div>
                    
                        <form action="sm_salereport.asp" name="saleReport" target="_blank" class="pcForms">
                        <%tmpArr=rs.getRows()
                        set rs=nothing
                        intCount=ubound(tmpArr,2)%>
                        Sale Name: <select name="id">
                        <option value="" selected>-- All Sales --</option>
                        <%For i=0 to intCount%>
                        <option value="<%=tmpArr(0,i)%>"><%=tmpArr(1,i)%></option>
                        <%Next%>
                        </select>
                        <br /><br />
                        <input type="submit" name="Submit" value="Search" class="btn btn-primary">
                        </form>
                
                    </p>
                    
                </div>
            </div>
        </div>
        <%
        END IF
        set rs=nothing
        %>





        <div class="panel panel-default">
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse14">
                        Top Products and Customers
                    </a>
                  </h4>
            </div>
            <div id="collapse14" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                        <div style="float: right; margin: 10px; width: 200px;">Use this form to list top selling products or top buying customers. Specify the number of results to be returned using the <i>Return Results</i> field.</div>
                    
                        <form action="resultsTopSells.asp" name="TopRep" target="_blank" class="pcForms" onSubmit="return(isDate(this.FromDate) && isDate(this.ToDate));">
                        <span id="show1">
                        From: <input class="datepicker" autocomplete="off" type="text" size="10" name="FromDate" value="<%=dtInputStrStart%>">&nbsp;
                        To: <input class="datepicker" autocomplete="off" type="text" size="10" name="ToDate" value="<%=dtInputStr%>">&nbsp;
                        Based on:&nbsp;<select name="basedon">
                            <option value="1" selected>Ordered Date</option>
                            <option value="2">Processed Date</option>
                            <option value="3">Shipped Date</option>
                            </select>
                        </span>
                        <br /><br />                    
                        Return Results: 
                        <input type="text" name="resultCnt" size="2" value="10">
                        <br /><br /> 
                        <select class="select" name="src" size="1" onchange="javascript: if ((document.TopRep.src.value=='2') || (document.TopRep.src.value=='4')) {document.getElementById('show1').style.display='none'} else {document.getElementById('show1').style.display=''};">
                        <option value="1">Top Selling Products</option>
                        <!-- The following option was disabled in v4.1 - See http://www.productcart.com/faqs/afmviewfaq.asp?faqid=588 -->
                        <!--<option value="2">Top Viewed Products</option>-->
                        <option value="4">Top 'Wish List' Products</option>
                        <option value="3">Top Customers</option>
                        </select>
                        <br /><br />
                        <input type="submit" name="Submit" value="Search" class="btn btn-primary">
                        </form>
                
                    </p>
                    
                </div>
            </div>
        </div>





        <div class="panel panel-default">
        
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse15">
                        Export Sales Data
                    </a>
                  </h4>
            </div>
            <div id="collapse15" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                        <div style="float: right; margin: 10px; width: 200px;">Export sales information by defining a date range and an export format. <b>Note</b>: You must enter both dates in the format <strong><%=strDateFormat%>.</strong></div>
                        
                        <% 
                            dim xlsTest, xlsObj
                            on error resume next
                            xlsObj=0
                            xlsTest=CreateObject("Excel.Application")
                            if err.number<>0 then
                                xlsObj=1
                            end if
                            err.number=0
                            err.clear
                            %>
                            <FORM ACTION="runquerySR.asp" METHOD="POST" class="pcForms" target="_blank" onSubmit="return(isDate(this.FromDate) && isDate(this.ToDate));">
                            From: <input class="datepicker" autocomplete="off" type="text" size="10" name="FromDate" value="<%=dtInputStrStart%>">&nbsp;
                            To: <input class="datepicker" autocomplete="off" type="text" size="10" name="ToDate" value="<%=dtInputStr%>">&nbsp;
                            Based on:&nbsp;<select name="basedon">
                                        <option value="1" selected>Ordered Date</option>
                                        <option value="2">Processed Date</option>
                                        <option value="3">Shipped Date</option>
                                        </select>
                            <br /><br />
                            Export&nbsp;format: 
                                <SELECT NAME="ReturnAS">
                                    <option value="HTML">HTML Table</option>
                                    <option value="CSV">CSV</option>
                                    <% if xlsObj=0 then %>
                                        <option value="XLS">Excel</option>
                                    <% end if %>
                                </SELECT>
                                <br /><br />
                                <INPUT TYPE="Submit" NAME="Submit" VALUE="Export" class="btn btn-primary">
                                &nbsp;
                                <input type="button" class="btn btn-default"  name="otherLinks" value="Other Export Tools" onClick="document.location.href='exportData.asp'">
                            </FORM>

							<% If statusAPP="1" OR scAPP=1 Then %>
			
								<form action="app-reportSales.asp" name="AppSale" class="pcForms" target="_blank" onSubmit="return Validate_Dates(this)">
									<p><b>View Total Sales by Apparel Products</b>
									<br /><br />
									From:  <input class="datepicker" autocomplete="off" type="text" size="10" name="FromDate" value="<%=dtInputStrStart%>">
									To:  <input  class="datepicker" autocomplete="off" type="text" size="10" name="ToDate" value="<%=dtInputStr%>">
									<br /><br />
									Products: <select class="select" name="src" size="1">
									<option value="0" selected> -- All Apparel Products -- </option>
									<%
									query="SELECT idproduct,description FROM products WHERE removed=0 and active=-1 and pcProd_Apparel=1 ORDER BY description ASC"
									set rstemp=conntemp.execute(query)
									if err.number <> 0 then
										response.redirect "techErr.asp?error="&Server.Urlencode("Error in order: "&err.description)
									end If
									if not rstemp.eof then
										pcArray=rstemp.getRows()
										intCount=ubound(pcArray,2)
										set rstemp=nothing
										For i=0 to intCount%>
											<option value="<%=pcArray(0,i)%>"><%=pcArray(1,i)%></option>
										<%Next
									end if
									set rstemp=nothing
									%>
									</select>                 
									<input type="submit" name="Submit" value="Search" class="submit2">
									</p>
								</form>
								<div>
									Use this form to list total sales for the Apparel products and for each sub-product. 
								</div>
			
							<% End If %>
                
                    </p>
                    
		        </div>
            </div>
        </div>




        <div class="panel panel-default">
            <div class="panel-heading">
                  <h4 class="panel-title">
                    <a data-toggle="collapse" data-parent="#accordion" href="#collapse16">
                        Back Order Report
                    </a>
                  </h4>
            </div>
            <div id="collapse16" data-parent="#accordion" class="panel-collapse collapse">
                <div class="panel-body">
                
                    <p>
                    
                    <div style="float: right; margin: 10px; width: 200px;">List of products that have been ordered in a date range, but are out of stock (they were ordered because back-ordering is allowed on those products). <b>Note</b>: You must enter both dates in the format <strong><%=strDateFormat%>.</strong></div>
                    
                    <FORM ACTION="BackOrderReport.asp" METHOD="POST" class="pcForms" target="_blank" onSubmit="return(isDate(this.FromDate) && isDate(this.ToDate));">
                    From: <input class="datepicker" autocomplete="off" type="text" size="10" name="FromDate" value="<%=dtInputStrStart%>">&nbsp;
                    To: <input class="datepicker" autocomplete="off" type="text" size="10" name="ToDate" value="<%=dtInputStr%>">
                    <br /><br>
                    <INPUT TYPE="Submit" NAME="Submit" VALUE="Search" class="btn btn-primary">
                    </FORM>
                
                    </p>
                    
                </div>
            </div> 
		</div>
        
</div>
<!--#include file="Adminfooter.asp"-->