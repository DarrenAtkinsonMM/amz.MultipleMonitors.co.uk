<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle = "Dashboard - Sales Charts and Graphs" 
pageIcon = ""
pcStrPageName = "dashboard.asp"
%>
<%PmAdmin=10%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/UpdateVersionCheck.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim pcvShowCharts

pcvShowCharts=1
%>
<!--#include file="pcCharts.asp"-->
	<table class="pcCPcontent" id="waitbox">    
		<tr>
			<td colspan="2">
				<div class="pcCPmessageInfo">
					Generating charts. Please wait...
				</div>
			</td>
		</tr>
	</table>
	<table class="pcCPcontent">    
		<%
		tmpYear=Year(Date())
		query="SELECT TOP 1 idorder FROM Orders WHERE OrderStatus>=2 AND Year(OrderDate)=" & tmpYear & ";"
		set rs=connTemp.execute(query)
		if (not rs.eof) AND (pcvShowCharts=1) then
		set rs=nothing%>
        <tr>
        	<td colspan="2">
            <div class="panel panel-default" style="background-image:url(images/pcv4_icon_chart.gif); background-position: 10px -10px; background-repeat:no-repeat;">
            <h2 style="padding-left: 60px;">Sales &amp; Other Data - Last 30 Days</h2>
			<table class="pcCPcontent">
                <tr>
                    <td colspan="2" valign="top">
                    <!--<div id="chartOrder30days" style="height:250px; width:49%; position:relative; float:left;"></div>-->
                    <div id="chartSales30days" style="height:330px; width:100%; position:relative; float:left; margin-bottom: -8px;"></div>
                    <!--<div id="chartTop10Prds30days" style="height:330px; width:49%; position:relative; float:left; margin-top: 15px;"></div>-->
                    <div id="chartTop10PrdsAmount30days" style="height:330px; width:100%; position:relative; float:left; margin-top: 15px; margin-bottom: -8px;"></div>
                    <div id="chartTop10Custs30days" style="height:330px; width:100%; position:relative; float:left; margin-top: 18px;"></div>
                    <div id="chartNewCusts30days" style="height:330px; width:100%; position:relative; float:right; margin-top: 18px;"></div>
                    <div style="clear: both;"></div>
                    <%
                    Dim ChartCount
                    ChartCount=0
                    'call pcs_Gen30daysALLOrdersCharts("chartOrder30days",0)
                    call pcs_Gen30daysCharts("chartOrder30days","chartSales30days",0,2)
                    'call pcs_Top10Prds30Days("chartTop10Prds30days")
                    call pcs_Top10PrdsAmount30Days("chartTop10PrdsAmount30days")
                    call pcs_Top10Custs30Days("chartTop10Custs30days")
                    call pcs_NewCusts30Days("chartNewCusts30days")
                    %>
                    </td>
                </tr>
            </table>
            </div>
            </td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
		<tr>
			<td width="100%" valign="top" colspan="2">     
            <div class="panel panel-default" style="background-image:url(images/pcv4_icon_sales.gif); background-position: 10px -10px; background-repeat:no-repeat; min-height: 142px; overflow: auto;">
				<h2 style="padding-left: 60px;"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_232")%> <span class="pcSmallText">&nbsp;|&nbsp;<a href="srcOrdByDate.asp">Other Sales Reports</a></span></h2>
				<table class="pcCPcontent">
					<tr>
						<td colspan="2">
							<div id="chartMonthlySales" style="height:250px;margin-bottom:18px;"></div>
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
						<td colspan="2"> 
							<%=dictLanguageCP.Item(Session("language")&"_cpCommon_231")%>: <b><%=scCurSign & money(pcv_YearTotal)%></b>
						</td>
					</tr>
					<%end if%>
										
						<% 
						
						
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
							<tr>
								<td colspan="2">
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_230")%>: &nbsp;
								<%
								Ayear=split(stryear,"***")
								For dd=1 to totalyear
								%>
								<a href="#" onClick="chgWin('salescharts.asp?year=<%=Ayear(dd-1)%>','window2')"><%=Ayear(dd-1)%></a>
								<%
								If dd <> totalyear Then Response.Write " - " End if
								Next
								%>
								</td>
							</tr>
						<%
						end if
						%>	
				</table>
                </div>
			</td>
		</tr>
		<%else%>
		<tr>
			<td>
				<div class="pcCPmessageInfo">Does not have any sales data of this year.</div>
			</td>
		</tr>
		<%end if
		set rs=nothing
		
		%>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
</table>
<script type=text/javascript>
	$pc(document).ready(function()
	{	
		document.getElementById("waitbox").style.display="none";
	});	
</script>
<!--#include file="AdminFooter.asp"-->
