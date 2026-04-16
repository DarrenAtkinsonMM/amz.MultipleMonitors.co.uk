<link rel="stylesheet" type="text/css" href="charts/jquery.jqplot.min.css" />
<!--[if lt IE 9]><script type="text/javascript" src="charts/excanvas.min.js"></script><![endif]-->
<script type="text/javascript" src="charts/jquery.jqplot.min.js"></script>
<script type="text/javascript" src="charts/plugins/jqplot.logAxisRenderer.min.js"></script>
<script type="text/javascript" src="charts/plugins/jqplot.pointLabels.min.js"></script>
<link rel="stylesheet" type="text/css" href="jchartfx/styles/jchartfx.css" />
<script type="text/javascript" src="jchartfx/js/jchartfx.system.js"></script>
<script type="text/javascript" src="jchartfx/js/jchartfx.coreVector.js"></script>
<script type="text/javascript" src="jchartfx/js/jchartfx.advanced.js"></script>
<script type="text/javascript" src="jchartfx/js/jchartfx.animation.js"></script>
<style>
.chartTab 
{
  /*
  border-top:1px solid #ccc;
  border-left:1px solid #ccc;
  border-right:1px solid #ccc;
  */
  font-weight:bold;
  overflow:hidden;
  position:relative;
  background:#b2b2b2;
  -webkit-border-top-left-radius:8px;
  -webkit-border-top-right-radius:8px;
  -moz-border-radius-topleft:8px;
  -moz-border-radius-topright:8px;
  border-top-left-radius:8px;
  border-top-right-radius:8px;
  margin:0 5px -1px 0;
  padding: 3px 10px 0px 10px;
  text-decoration:none !important;
  color:White !important;
}
.chartTab:hover {
  background-color:#0d293f;
  text-decoration:none !important;
}
.chartSelected 
{
  background-color:#226faa !important;
  text-decoration:none !important;
}
.tableTabContainer
{
  float:left;
  margin-bottom: 0px;
}
.tableTab 
{
  font-weight:bold;
  cursor:pointer;
  /*
  border-top:1px solid #ccc;
  border-left:1px solid #ccc;
  border-right:1px solid #ccc;
  */
  overflow:hidden;
  position:relative;
  background:#b2b2b2;
  -webkit-border-bottom-left-radius:8px;
  -webkit-border-bottom-right-radius:8px;
  -moz-border-radius-bottomleft:8px;
  -moz-border-radius-bottomright:8px;
  border-bottom-left-radius:8px;
  border-bottom-right-radius:8px;
  margin:0 10px 0 10px;
  padding: 0px 10px 5px 10px;
  text-decoration:none !important;
  color:White;
}
.tableTab a:link {
  color:White !important;
  text-decoration:none !important;
}
.tableTab a:visited  {
  color:White !important;
  text-decoration:none !important;
}
.tableTab a:active  {
  color:White !important;
  text-decoration:none !important;
}
.tableTab:hover {
  color:White !important;
  background-color: #226FAA;
  text-decoration:none !important;
}
.tableTabSelected 
{
  background-color:#226FAA;
  text-decoration:none !important;
}
.pcDailyTableContainer
{
  display:none;
}
.pcDailyTable tr td:first-child + td
{
  text-align:center;
}
.pcDailyTable tr td:first-child + td + td
{
  text-align:right;
}
.pcDailyTable tbody tr td
{
  border-bottom: 1px solid #CCC;
}
.pcDailyTable tbody tr:hover
{
  background-color:#E6E6E6;
}
.PointLabel {
    display: none;   
}
.Title0 {
    color: #226FAA;
    font-weight: 600;
}
</style>
<script>
var pcv_strSymbol = '<%=scCurSign%>';
</script>
<script type=text/javascript>

var chart;

function CurrencyFormatted(amount) {
  var i = parseFloat(amount);
  if (isNaN(i)) { i = 0.00; }
  var minus = '';
  if (i < 0) { minus = '-'; }
  i = Math.abs(i);
  i = parseInt((i + .005) * 100);
  i = i / 100;
  s = new String(i);
  if (s.indexOf('.') < 0) { s += '.00'; }
  if (s.indexOf('.') == (s.length - 2)) { s += '0'; }
  s = minus + s;
  s = "$" + s;
  return s;
}
</script>
<%
Dim pcvHave30Days, gridOptions
pcvHave30Days=0
gridOptions=", grid: {borderWidth:0.5, borderColor:'#CCC', shadow:false}"


Function pcf_GetOrderStatusTXT(porderstatus)

    select case porderstatus
        case "0",""
            pcf_GetOrderStatusTXT="N/A"
        case "1"
          pcf_GetOrderStatusTXT="Incomplete"
        case "2"
          pcf_GetOrderStatusTXT="Pending" 
        case "3"
          pcf_GetOrderStatusTXT="Processed" 
        case "4"
          pcf_GetOrderStatusTXT="Shipped" 
        case "5"
          pcf_GetOrderStatusTXT="Canceled" 
        case "6"
          pcf_GetOrderStatusTXT="Return" 
        case "7"
          pcf_GetOrderStatusTXT="Partially Shipped"
        case "8"
          pcf_GetOrderStatusTXT="Shipping"
        case "9"
          pcf_GetOrderStatusTXT="Partially Return"
        case "10"
          pcf_GetOrderStatusTXT="Delivered" 
        case "11"
          pcf_GetOrderStatusTXT="Will Not Deliver" 
        case "12"
          pcf_GetOrderStatusTXT="Archived"
	end select
    
End Function



Function pcf_CustTypeTXT(pcusttype)

    select case pcusttype
        case "0",""
            pcf_CustTypeTXT="Registered"
        case "1"
            pcf_CustTypeTXT="Guests"
        case "2"
            pcf_CustTypeTXT="Duplicated" 
	end select
    
End Function



Private Sub pcs_Gen30daysALLOrdersCharts(DivName,ShowLegend)

    Dim past30,Datenow,rs,query,tmpArr,i,j,intCount,tmpline1,tmpline2,xname,xname1
    Dim line1(30),line2(30),line3(30),line4(30),line5(30),tmpDate

	Datenow=Date()
	past30=Date()-29
	
	For i=29 to 0 step -1
		tmpDate=Date()-i
		line1(29-i)=Day(tmpDate)
		line2(29-i)=Month(tmpDate)
		line3(29-i)=0
		line4(29-i)=0
		if scDateFrmt="DD/MM/YY" then
			line5(29-i)=Day(tmpDate) & "/" & Month(tmpDate) & "/" & Year(date())
		else
			line5(29-i)=Month(tmpDate) & "/" & Day(tmpDate) & "/" & Year(date())
		end if
	Next
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	query="SELECT Day(OrderDate) As TheDay,Month(OrderDate) As TheMonth,Count(*) As TotalOrders FROM orders WHERE ((orders.orderStatus>=2 AND orders.orderStatus<5) OR (orders.orderStatus>=6)) AND orderdate>='" & past30 & "' AND orderdate<='" & Datenow & "' GROUP BY month(orderdate),day(orderdate) ORDER BY month(orderdate) ASC,day(orderdate) ASC;"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		pcvHave30Days=1
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		tmpline1=""
		xname=""
		xname1=""
		For i=0 to intCount
			For j=0 to 29
			if (Cint(tmpArr(0,i))=Cint(line1(j))) AND (Cint(tmpArr(1,i))=Cint(line2(j))) then
				line3(j)=Round(tmpArr(2,i),2)
				if scDateFrmt="DD/MM/YY" then
					line5(j)=tmpArr(0,i) & "/" & tmpArr(1,i) & "/" & Year(date())
				else
					line5(j)=tmpArr(1,i) & "/" & tmpArr(0,i) & "/" & Year(date())
				end if
				exit for
			end if
			Next
		Next
		
		For i=0 to 29
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				xname=xname & ","
				xname1=xname1 & ","
			end if
			tmpline1=tmpline1 & line3(i)
			xname1=xname1 & "{'Month':'" & line5(i) & "','Orders':" & line3(i) & "}"
			xname=xname & "'" & replace(line5(i),"/","\%2F") & "'"
		Next

		%>
		<script type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>
		<script type=text/javascript>
            
            var chart2;
            $pc(document).ready(function () {

                xname = [<%=xname%>];
            
                var xname1 = [<%=xname1%>];
                
                chart2 = new cfx.Chart();
                chart2.setDataSource(xname1);
                chart2.setOptions({
                    gallery: cfx.Gallery.Bar,
                    titles: [{ text: "Number of Orders"}],
                    animations: {
                        load: { enabled: true }
                    }
                });
                chart2.getAllSeries().getPointLabels().setVisible(true);               
                chart2.create("<%=DivName%>")
            
                $pc("#<%=DivName%>").click(function(evt) {
                    if ((evt.hitType == cfx.HitType.Point) || (evt.hitType == cfx.HitType.Between)) {
                        var s = "Series " + evt.series + " Point " + evt.point;
                        if (evt.hitType == cfx.HitType.Between)
                        s += " Between";
				        var tmpURL="resultsAdvancedAll.asp?fromdate=" + xname[evt.point] + "&todate=" + xname[evt.point] + "&otype=0&PayType=&B1=Search+Orders"
				        window.open(tmpURL,"_blank");
                    }
                });
      
            });
    </script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else
	pcvHave30Days=0%>
	<div>A quick summary for last 30 days cannot be created as no orders have yet been processed. Please note that pending, returned, and cancelled orders that might have been placed in the current year are not included in sales reports.</div>
	<script type=text/javascript>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
End Sub



Private Sub pcs_Gen30daysCharts(DivName,DivName1,ShowLegend,NumCharts)
    
    Dim past30,Datenow,rs,query,tmpArr,i,j,intCount,tmpline1,tmpline2,xname,chartTitle,xname1
    Dim line1(30),line2(30),line3(30),line4(30),line5(30),tmpDate

	Datenow=Date()
	past30=Date()-29
	
	For i=29 to 0 step -1
		tmpDate=Date()-i
		line1(29-i)=Day(tmpDate)
		line2(29-i)=Month(tmpDate)
		line3(29-i)=0
		line4(29-i)=0
		if scDateFrmt="DD/MM/YY" then
			line5(29-i)=Day(tmpDate) & "/" & Month(tmpDate) & "/" & Year(tmpDate)
		else
			line5(29-i)=Month(tmpDate) & "/" & Day(tmpDate) & "/" & Year(tmpDate)
		end if
	Next
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	query="SELECT Day(OrderDate) As TheDay,Month(OrderDate) As TheMonth,Count(*) As TotalOrders, Sum(Total-rmaCredit) AS TotalAmounts, Sum(Total) AS TotalLessRMA, Year(OrderDate) As TheYear FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orderdate>='" & past30 & "' AND orderdate<='" & Datenow & "' GROUP BY year(orderdate), month(orderdate),day(orderdate) ORDER BY year(orderdate), month(orderdate) ASC,day(orderdate);"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		tmpline1=""
		tmpline2=""
		xname=""
		xname1=""
        xname2=""
        xname3=""
		For i=0 to intCount
			For j=0 to 29
			if (Cint(tmpArr(0,i))=Cint(line1(j))) AND (Cint(tmpArr(1,i))=Cint(line2(j))) then
				line3(j)=Round(tmpArr(2,i),2)
				if isNULL(tmpArr(3,i)) then
					tmpArr(3,i) = tmpArr(4,i)
				end if
				line4(j)=Round(tmpArr(3,i),2)
				if scDateFrmt="DD/MM/YY" then
					line5(j)=tmpArr(0,i) & "/" & tmpArr(1,i) & "/" & tmpArr(5,i)
				else
					line5(j)=tmpArr(1,i) & "/" & tmpArr(0,i) & "/" & tmpArr(5,i)
				end if
				exit for
			end if
			Next
		Next
		
		For i=0 to 29
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				tmpline2=tmpline2 & ","
				xname=xname & ","
				xname1=xname1 & ","
                xname2=xname2 & ","
                xname3=xname3 & ","
			end if
			tmpline1=tmpline1 & line3(i)
			tmpline2=tmpline2 & line4(i)
            xname1=xname1 & "{'Month':'" & line5(i) & "','Sales':" & line4(i)  &  "}"
            xname2=xname2 & "{'Month':'" & line5(i) & "','Orders':" & line3(i) &  "}"
            xname3=xname3 & "{'Month':'" & line5(i) & "','Orders':" & line3(i) & ",'Sales':" & line4(i)  &"}"
			xname=xname & "'" & replace(line5(i),"/","\%2F") & "'"
		Next
		%>
		<script type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>
		
		<%if NumCharts="1" OR NumCharts="0" then%>
		<script type=text/javascript>$pc(document).ready(function(){
		    
            line1 = [<%=tmpline1%>];
		    
            plot2 = $pc.jqplot('<%=DivName%>', [line1], {
			
            <% if ShowLegend=1 then %>
			    legend:{show:true, location:'ne', xoffset:55},
			<% end if %>
            
                title:'Number of Orders',
                series:[
                    {
                        renderer:$pc.jqplot.BarRenderer, 
                        rendererOptions: {
                            barWidth:8   
                        },
                        label:'Number of Orders',
                        pointLabels:{show:true, stackedValue: true, hideZeros:true}
                    }
                ],
                axes:{
                    xaxis:{
                        renderer:$pc.jqplot.CategoryAxisRenderer,
                        ticks: [<%=xname%>],
                        rendererOptions:{tickRenderer:$pc.jqplot.CanvasAxisTickRenderer},
                        tickOptions:{
                        fontSize:'10px', 
                        fontFamily:'Arial', 
                        angle:-30
                        }
                    },
                    yaxis:{min:0,autoscale:true, tickOptions:{formatString:'%.0f'}}
                }
                <%=gridOptions%>
		    });	
        
            });
        
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
		<%end if%>
		<%if NumCharts="2" OR NumCharts="0" then
		chartTitle="Daily Sales Amount"%>
		<script type=text/javascript>
    var isDailySales = true;
    var isByDate = true;
    
    var chart3;
    function loadDailyOrder() {
      
        var dailylist = [<%=xname2%>];
        if(!isByDate){
            dailylist.sort(function(a,b){return b['Orders'] - a['Orders']});
        }
        isDailySales = false;

        chart3.setDataSource(dailylist);
        chart3.setOptions({
            gallery: cfx.Gallery.Lines,
            titles: [{ text: "Daily Orders"}],
            axisY : {
                dataFormat: {
                    format: cfx.AxisFormat.Number,
                    decimals: 0
                },
                labelsFormat: {
                    format: cfx.AxisFormat.Number
                }
                },
                animations: {
                    load: {
                        enabled: true
                }
            }
        });
        chart3.getAllSeries().getPointLabels().setVisible(true);        

    }
    
    function getTool(args){
      //$pc('.jcharttip').append(2);
    }


    function loadDailyAmount() {
      
        var dailylist = [<%=xname1%>];
        if(!isByDate){
            dailylist.sort(function(a,b){return b['Sales'] - a['Sales']});
        }
        isDailySales = true;
            
        chart3.setDataSource(dailylist);
        chart3.setOptions({
            gallery: cfx.Gallery.Bar,
            titles: [{ text: "Daily Sales"}],
            axisY : {
                dataFormat: {
                    format: cfx.AxisFormat.Currency,
                    decimals: 2
                },
                labelsFormat: {
                    format: cfx.AxisFormat.Currency
                }
                },
                animations: {
                    load: {
                        enabled: true
                }
            }
        });
        chart3.getAllSeries().getPointLabels().setVisible(true);        
 
    }
    $pc(document).ready(function () {
        chart3 = new cfx.Chart(); 
        chart3.create("<%=DivName1%>")
    });

    function OrderBy(isDate)
    {
      isByDate= isDate;
      if(isDailySales){

        loadDailyAmount();

      } else {
          
        loadDailyOrder();
        
      }
    }
    function loadDailySalesTable()
    {
      if($pc('.pcDailyTab').text().indexOf("Show") >=0){
        $pc('.pcDailyTab').text('Hide Daily Sales Table');
        $pc('.pcDailyTab').attr('class','tableTab tableTabSelected pcDailyTab');
      }
      else {
        $pc('.pcDailyTab').text('Show Daily Sales Table');
        $pc('.pcDailyTab').attr('class','tableTab pcDailyTab');
      }
      $pc('.pcDailySalesTableContainer').slideToggle(500);
    }
		$pc(document).ready(function () {
		  datename=[<%=xname%>];
      dailyTableArray=[<%=xname3%>];
      loadDailyAmount();

      $pc("#<%=DivName1%>").prepend("<div  style='margin: 0 10px -10px 10px; cursor:pointer;'><a class='chartTab chartSelected' onclick='loadDailyAmount();'>Sales</a> <a class='chartTab' onclick='loadDailyOrder();'>Orders</a></div>");
      $pc("#<%=DivName1%>").prepend("<div  style='margin:0 10px -10px 10px;float:right; cursor:pointer;'><a class='chartTab chartSelected' onclick='OrderBy(true);'>Sort by Date</a> <a class='chartTab'  onclick='OrderBy(false);'>Sort by Sales/Unit</a></div>");
      $pc("#<%=DivName1%>").bind( "mouseover", getTool );
      $pc("#<%=DivName1%>").click(function(evt) {
        if ((evt.hitType == cfx.HitType.Point) || (evt.hitType == cfx.HitType.Between)) {
          var s = "Series " + evt.series + " Point " + evt.point;
          if (evt.hitType == cfx.HitType.Between)
            s += " Between";
				  var tmpURL="viewDateOrders.asp?FromDate=" + datename[evt.point] + "&ToDate=" + datename[evt.point] + "&basedon=1&customerType=&CountryCode=&submit=Search"
				  window.open(tmpURL,"_blank");
        }
      });
      $pc("<div class='tableTabContainer'><div class='pcDailyTableContainer pcDailySalesTableContainer'><table class='pcDailyTable pcDailySalesTable'><tr><th>Date</th><th>Number of Orders</th><th>Sales</th></tr></table></div><a class='tableTab pcDailyTab' onclick='loadDailySalesTable();'>Show Daily Sales Table</a></div>").insertAfter($pc("#<%=DivName1%>"));
      $pc.each(dailyTableArray, function(i,field){
        $pc(".pcDailySalesTable").append('<tr><td>' + field.Month + '</td><td>' + field.Orders + '</td><td>' + CurrencyFormatted(field.Sales) + '</td></tr>');
      });
      $pc('.chartTab').click(function () {
        $pc(this).parent().find('.chartTab').attr("class", "chartTab");
        $pc(this).attr("class", "chartTab chartSelected");
      });
    });
    </script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script type=text/javascript>
			document.getElementById("<%=DivName1%>").style.clear='both';
			document.getElementById("<%=DivName1%>").style.float='left';
		</script>
		<%else%>
		<script type=text/javascript>
			document.getElementById("<%=DivName1%>").style.float='right';
		</script>
		<%end if%>
		<%end if%>
	<%else%>
	<%if NumCharts="1" OR NumCharts="0" then%>
	<div>A quick summary for last 30 days cannot be created as no orders have yet been processed. Please note that pending, returned, and cancelled orders that might have been placed in the current year are not included in sales reports.</div>
	<%end if%>
	<script type=text/javascript>
		<%if NumCharts="1" OR NumCharts="0" then%>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
		<%end if%>
		<%if NumCharts="2" OR NumCharts="0" then%>
		document.getElementById("<%=DivName1%>").style.height='0px';
		document.getElementById("<%=DivName1%>").style.display='none';
		<%end if%>
	</script>
	<%end if
	set rs=nothing
End Sub



Private Sub pcs_MonthlySalesChart(DivName,TheYear,FullYear,ShowLegend)

	Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname
	Dim line1(12),line2(12),tmpDate

	For i=0 to 11
		line1(i)=MonthName(i+1, True)
		line2(i)=0
	Next
	
	yearnow = TheYear
		
	query = "SELECT Month(OrderDate) As TheMonth,Sum(Total-rmaCredit) AS TotalAmounts, Sum(Total) AS TotalLessRMA, Sum(rmaCredit) AS TotalRMA FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND year(orderdate)=" & yearnow & " GROUP BY month(orderdate) ORDER BY month(orderdate) ASC;"
	Set rs = Server.CreateObject("ADODB.Recordset")
    Set rs = connTemp.execute(query)	
	If Not rs.Eof Then
    
		tmpArr = rs.getRows()
		Set rs = Nothing
		
        intCount=ubound(tmpArr,2)
		tmpline1=""
		xname=""
        
		For i=0 to intCount
			'Override TotalAmounts due to possible NULLS
			if NOT isNumeric(tmpArr(3,i)) then
				tmpArr(3,i)=0
			end if
			tmpArr(1,i) = tmpArr(2,i)-tmpArr(3,i)

			pcv_YearTotal=pcv_YearTotal+Clng(tmpArr(1,i))

			For j=0 to 11
			if (Cint(tmpArr(0,i))=Cint(j+1)) then
				line2(j)=Clng(tmpArr(1,i))
				exit for
			end if
			Next
		Next

		For i=0 to 11 step +1
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				xname=xname & ","
			end if
			tmpline1=tmpline1 & "{ 'Sales' : " & line2(i) & ", 'Month': '" & line1(i) & "'}"
			if FullYear=0 then
				if (Cint(i+1)>Cint(Month(Date()))) then
					xname=xname & "' '"
				else
					xname=xname & "'" & line1(i) & "'"
				end if
			else
				xname=xname & "'" & line1(i) & "'"
			end if
		Next

		%>
		<script type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		
        <script type=text/javascript>

            function monthlyOrderBy(isByDate) {
                
                line1 = [<%=tmpline1%>];
                
                var monthlylist = [<%=tmpline1%>];
                if(!isByDate){
                    monthlylist.sort(function(a,b){return b['Sales'] - a['Sales']});
                }
    
                chart.setDataSource(monthlylist);
                chart.setOptions({
                    gallery: cfx.Gallery.Bar,
                    titles: [{ text: "Monthly Sales - <%=TheYear%>"}],
                    axisY : {
                        dataFormat: {
                            format: cfx.AxisFormat.Currency
                        },
                        labelsFormat: {
                            format: cfx.AxisFormat.Currency
                        }
                        },
                        animations: {
                            load: {
                                enabled: true
                        }
                    }
                });
                chart.getAllSeries().getPointLabels().setVisible(true);
    
            }
            $pc(document).ready(function () {
                chart = new cfx.Chart();
                chart.create("<%=DivName%>")
                monthlyOrderBy(true);
            });
        
            $pc("#<%=DivName%>").prepend("<div style='margin: 0 10px -10px 10px; float:right; cursor:pointer;'><a class='chartTab chartSelected' onclick='monthlyOrderBy(true);'>Sort by Date</a> <a class='chartTab' onclick='monthlyOrderBy(false);'>Sort by Sales</a></div>");
            $pc('.chartTab').click(function () {
            $pc(this).parent().find('.chartTab').attr("class", "chartTab");
            $pc(this).attr("class", "chartTab chartSelected");
            });
		</script>
        
	<% Else %>
    
        <div class="pcCPmessageInfo">A sales report for the current year cannot be created as no orders have yet been processed. Please note that pending, returned, and cancelled orders that might have been placed in the current year are not included in sales reports.</div>
        <script type=text/javascript>
            document.getElementById("<%=DivName%>").style.height='0px';
            document.getElementById("<%=DivName%>").style.display='none';
        </script>
    
	<%
    End If
	Set rs = Nothing

End Sub



Private Sub pcs_Top10Prds30Days(DivName)

    Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,pcArr,icount,rsQ
    Dim Datenow,past30

	Datenow=Date()
	past30=Date()-29
	
	if scDateFrmt="DD/MM/YY" then
		tmpDateNow=Day(Datenow) & "/" & Month(Datenow) & "/" & Year(Datenow)
		tmppast30=Day(past30) & "/" & Month(past30) & "/" & Year(past30)
	else
		tmpDateNow=Month(Datenow) & "/" & Day(Datenow) & "/" & Year(Datenow)
		tmppast30=Month(past30) & "/" & Day(past30) & "/" & Year(past30)
	end if

	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	query="SELECT TOP 10 ProductsOrdered.IDProduct, SUM(ProductsOrdered.Quantity) As PrdSales FROM ProductsOrdered, Orders where Orders.IDOrder=ProductsOrdered.IDOrder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >='" & past30 & "' AND orders.orderDate <='" & Datenow & "' GROUP BY ProductsOrdered.IDProduct ORDER BY SUM(ProductsOrdered.Quantity) DESC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if not rs.eof then
		icount=0
		tmpline1=""
		tmpline2=""
		tmpline3=""
        tmpline4=""
        
		do while (not rs.eof) AND (icount<10)
			icount=icount+1
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				tmpline2=tmpline2 & ","
				tmpline3=tmpline3 & ","
                tmpline4=tmpline4 & ","
			end if
			query="SELECT description FROM Products WHERE idproduct=" & rs("IDProduct") & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				pcStrProductCompact = rsQ("description")
			else
				pcStrProductCompact = "N/A"
			end if
			set rsQ=nothing
			if len(pcStrProductCompact)>25 then
			 pcStrProductCompact = left(pcStrProductCompact,22) & "..."
			end if
			pcStrProductCompact=replace(pcStrProductCompact,"'","\'")
			tmpline1=tmpline1 & Clng(rs("PrdSales"))
			tmpline2=tmpline2 & "'" & pcStrProductCompact & "'"
			tmpline3=tmpline3 & rs("IDProduct")
            tmpline4=tmpline4 & "{'Product' : '" & pcStrProductCompact &"','Units': "& Clng(rs("PrdSales")) & "}"
			rs.MoveNext
		loop
		set rs=nothing
        
		if icount<10 then
			For k=icount+1 to 10
			tmpline1=tmpline1 & ",0"
			if k<10 then
				tmpline2=tmpline2 & ",' '"
			else
				tmpline2=tmpline2 & ",'.'"
			end if
			Next
		end if	
		%>
		<script type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>
		
		<script type=text/javascript>
    
            $pc(document).ready(function () {
            
                line1 = [<%=tmpline1%>];
		        prdArr1=[<%=tmpline3%>];
                var tmpline4 = [<%=tmpline4%>];
                
                $pc("#<%=DivName%>").innerHTML('');
                chart = new cfx.Chart();            
                chart.setDataSource(tmpline4);
                chart.setOptions({
                    gallery: cfx.Gallery.Bar,
                    titles: [{ text: "Top 10 Selling Products (Units)"}],
                    animations: {
                        load: {
                            enabled: true
                        }
                    }
                });
                chart.getAllSeries().getPointLabels().setVisible(true);        
                chart.create("<%=DivName%>")
                
                
      
                $pc("#<%=DivName%>").click(function(evt) {
                    if ((evt.hitType == cfx.HitType.Point) || (evt.hitType == cfx.HitType.Between)) {
                        var s = "Series " + evt.series + " Point " + evt.point;
                        if (evt.hitType == cfx.HitType.Between)
                        s += " Between";
				        var tmpURL="viewPrdDateOrders.asp?FromDate=<%=replace(tmppast30,"/","\%2F")%>&ToDate=<%=replace(tmpDateNow,"/","\%2F")%>&basedon=1&IDProduct=" + prdArr1[evt.point] + "&submit=Search"
				        window.open(tmpURL,"_blank");
                    }
                });
            });
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
	<script type=text/javascript>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
End Sub



Private Sub pcs_Top10PrdsAmount30Days(DivName)

    Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,pcArr,icount,rsQ
    Dim Datenow,past30

	Datenow=Date()
	past30=Date()-29
	
	if scDateFrmt="DD/MM/YY" then
		tmpDateNow=Day(Datenow) & "/" & Month(Datenow) & "/" & Year(Datenow)
		tmppast30=Day(past30) & "/" & Month(past30) & "/" & Year(past30)
	else
		tmpDateNow=Month(Datenow) & "/" & Day(Datenow) & "/" & Year(Datenow)
		tmppast30=Month(past30) & "/" & Day(past30) & "/" & Year(past30)
	end if
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	query="SELECT TOP 10 ProductsOrdered.IDProduct, SUM(ProductsOrdered.Quantity) As PrdAmounts , SUM(ProductsOrdered.UnitPrice*ProductsOrdered.Quantity-ProductsOrdered.QDiscounts-ProductsOrdered.ItemsDiscounts) As PrdSales FROM ProductsOrdered, Orders where Orders.IDOrder=ProductsOrdered.IDOrder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >='" & past30 & "' AND orders.orderDate <='" & Datenow & "' GROUP BY ProductsOrdered.IDProduct ORDER BY SUM(ProductsOrdered.UnitPrice*ProductsOrdered.Quantity) DESC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if not rs.eof then
		icount=0
		tmpline1=""
		tmpline2=""
		tmpline3=""
        tmpline4=""
        tmpline5=""
        tmpline6=""
		do while (not rs.eof) AND (icount<10)
			icount=icount+1
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				tmpline2=tmpline2 & ","
				tmpline3=tmpline3 & ","
                tmpline4=tmpline4 & ","
                tmpline5=tmpline5 & ","
                tmpline6=tmpline6 & ","
			end if
			query="SELECT description FROM Products WHERE idproduct=" & rs("IDProduct") & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				pcStrProductCompact = rsQ("description")
                pcStrProduct =rsQ("description")
			else
				pcStrProductCompact = "N/A"
                pcStrProduct = "N/A"
			end if
			set rsQ=nothing
			if len(pcStrProductCompact)>25 then
			 pcStrProductCompact = left(pcStrProductCompact,22) & "..."
			end if
			pcStrProductCompact=replace(pcStrProductCompact,"'","\'")
            pcStrProduct=replace(pcStrProduct,"'","\'")
            
			tmpline1=tmpline1 & Round(rs("PrdSales"),2)
			tmpline2=tmpline2 & "'" & pcStrProductCompact & "'"
			tmpline3=tmpline3 & rs("IDProduct")
            tmpline4=tmpline4 & "{'Product' : '" & pcStrProductCompact &"','Sales': "& rs("PrdSales") & "}"
            tmpline5=tmpline5 & "{'Product' : '" & pcStrProductCompact & "','Units': "& rs("PrdAmounts") & "}"
            tmpline6=tmpline6 & "{'Product' : '" &  pcStrProduct &"','Sales': "& rs("PrdSales") & "}"
			rs.MoveNext
		loop
		set rs=nothing 
		if icount<10 then
			For k=icount+1 to 10
			tmpline1=tmpline1 & ",0"
			if k<10 then
				tmpline2=tmpline2 & ",' '"
			else
				tmpline2=tmpline2 & ",'.'"
			end if
			Next
		end if	
		%>
		<script type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>

		<script type=text/javascript>
        
        var chart4;
        function loadUnits() {
          
            var amountarray = [<%=tmpline5%>];
            amountarray.sort(function(a,b){return b['Units'] - a['Units']});
        
            chart4.setDataSource(amountarray);
            chart4.setOptions({
                gallery: cfx.Gallery.Bar,
                titles: [{ text: "Top 10 Selling Products (Units)"}],
                axisY : {
                    dataFormat: {
                        format: cfx.AxisFormat.Number,
                        decimals: 0
                    },
                    labelsFormat: {
                        format: cfx.AxisFormat.Number
                    }
                    },
                    animations: {
                        load: {
                            enabled: true
                    }
                }
            });
            chart4.getAllSeries().getPointLabels().setVisible(true);        

        }

        function loadAmount(){

            var amountarray = [<%=tmpline4%>];
            chart4.setDataSource(amountarray);
            chart4.setOptions({
                gallery: cfx.Gallery.Bar,
                titles: [{ text: "Top 10 Selling Products (Sales)"}],
                axisY : {
                    dataFormat: {
                        format: cfx.AxisFormat.Currency,
                        decimals: 2
                    },
                    labelsFormat: {
                        format: cfx.AxisFormat.Currency
                    }
                    },
                    animations: {
                        load: {
                            enabled: true
                    }
                }
            });
            chart4.getAllSeries().getPointLabels().setVisible(true);        

        }
        $pc(document).ready(function () {
            chart4 = new cfx.Chart(); 
            chart4.create("<%=DivName%>")
        });
    
    
    function loadTopSellingTable()
    {
      if($pc('.pcTopSellingTab').text().indexOf("Show") >=0){
        $pc('.pcTopSellingTab').text('Hide Top 10 Selling Products Table');
        $pc('.pcTopSellingTab').attr('class','pcTopSellingTab tableTab tableTabSelected');
      }
      else {
        $pc('.pcTopSellingTab').text('Show Top 10 Selling Products Table');
        $pc('.pcTopSellingTab').attr('class','pcTopSellingTab tableTab');
      }
      $pc('.pcTopSellingTabContainer').slideToggle(500);
    }
    $pc(document).ready(function () {
    line1 = [<%=tmpline1%>];
		prdArr2=[<%=tmpline3%>];
    topSellingArray=[<%=tmpline6%>];
    loadAmount();

		  $pc("#<%=DivName%>").prepend("<div style='margin:0 10px -10px 10px; cursor:pointer;'><a class='chartTab chartSelected' onclick='loadAmount();'>Sales</a> <a class='chartTab' onclick='loadUnits();'>Units</a></div>");
      $pc("#<%=DivName%>").click(function(evt) {
        if ((evt.hitType == cfx.HitType.Point) || (evt.hitType == cfx.HitType.Between)) {
          var s = "Series " + evt.series + " Point " + evt.point;
          if (evt.hitType == cfx.HitType.Between)
            s += " Between";
				  var tmpURL="viewPrdDateOrders.asp?FromDate=<%=replace(tmppast30,"/","\%2F")%>&ToDate=<%=replace(tmpDateNow,"/","\%2F")%>&basedon=1&IDProduct=" + prdArr2[evt.point] + "&submit=Search"
				  window.open(tmpURL,"_blank");
        }
      });
      $pc("<div class='tableTabContainer'><div class='pcDailyTableContainer pcTopSellingTabContainer'><table class='pcDailyTable pcTopSelling'><tr><th>Product</th><th>Number of Orders</th></tr></table></div><a class='tableTab pcTopSellingTab' onclick='loadTopSellingTable();'>Show Top 10 Selling Products Table</a></div>").insertAfter($pc("#<%=DivName%>"));
      $pc.each(topSellingArray, function(i,field){
        $pc(".pcTopSelling").append('<tr><td>' + field.Product + '</td><td>' + CurrencyFormatted(field.Sales) + '</td></tr>');
      });
      $pc('.chartTab').click(function () {
        $pc(this).parent().find('.chartTab').attr("class", "chartTab");
        $pc(this).attr("class", "chartTab chartSelected");
      });
    });
	 	</script>
	<%end if
	set rs=nothing
End Sub

Private Sub pcs_Top10Custs30Days(DivName)

    Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,pcArr,icount,rsQ
    Dim Datenow,past30

	Datenow=Date()
	past30=Date()-29
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	query="SELECT TOP 10 idcustomer, sum(total) As AmountTotal, count(*) As NumOrders FROM Orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >='" & past30 & "' AND orders.orderDate <='" & Datenow & "' GROUP BY idcustomer ORDER BY sum(total) DESC,count(*) DESC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if not rs.eof then
        pcvHave30Days=1
		icount=0
		tmpline1=""
		tmpline2=""
		tmpline3=""
        tmpline4=""
		do while (not rs.eof) AND (icount<10)
			icount=icount+1
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				tmpline2=tmpline2 & ","
				tmpline3=tmpline3 & ","
                tmpline4=tmpline4 & ","
			end if
			query="SELECT name, lastname FROM Customers WHERE idcustomer=" & rs("idcustomer")
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				pcStrNameCompact = rsQ("name") & " " & rsQ("lastname")
			else
				pcStrNameCompact = ""
			end if
			set rsQ=nothing
			if len(pcStrNameCompact)>25 then
			 pcStrNameCompact = left(pcStrNameCompact,22) & "..."
			end if
			pcStrNameCompact=replace(pcStrNameCompact,"'","\'")
			tmpline1=tmpline1 & Clng(rs("AmountTotal"))
			tmpline2=tmpline2 & "'" & pcStrNameCompact & "'"
			tmpline3=tmpline3 & rs("idcustomer")
            tmpline4=tmpline4 & "{'Customer' : '" & pcStrNameCompact &"','Sales': "&  rs("AmountTotal") & "}"
			rs.MoveNext
		loop
		set rs=nothing
		if icount<10 then
			For k=icount+1 to 10
			tmpline1=tmpline1 & ",0"
			if k<10 then
				tmpline2=tmpline2 & ",' '"
			else
				tmpline2=tmpline2 & ",'.'"
			end if
			Next
		end if	
		%>
		<script type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>
		
		<script type=text/javascript>
    
        
        $pc(document).ready(function () {
        
            line1 = [<%=tmpline1%>];
            custArr1=[<%=tmpline3%>];

            var chart5;
            chart5 = new cfx.Chart();
            chart5.create("<%=DivName%>")
            
            var topCustomersList = [<%=tmpline4%>];
            
            <% If len(tmpline4)>0 Then %> 
                 
                chart5.setDataSource(topCustomersList);
                chart5.setOptions({
                    gallery: cfx.Gallery.Bar,
                    titles: [{ text: "Top 10 Customers (Sales)"}],
                    axisY : {
                        dataFormat: {
                            format: cfx.AxisFormat.Currency,
                            decimals: 2
                        },
                        labelsFormat: {
                            format: cfx.AxisFormat.Currency
                        }
                        },
                        animations: {
                            load: {
                                enabled: true
                        }
                    }
                });
                chart5.getAllSeries().getPointLabels().setVisible(true);      

            <% End If %>
          
          
              $pc("#<%=DivName%>").click(function(evt) {
                if ((evt.hitType == cfx.HitType.Point) || (evt.hitType == cfx.HitType.Between)) {
                  var s = "Series " + evt.series + " Point " + evt.point;
                  if (evt.hitType == cfx.HitType.Between)
                    s += " Between";
                          var tmpURL="viewCustOrders.asp?idcustomer=" + custArr1[evt.point]
                          window.open(tmpURL,"_blank");
                }
              });

        });
		</script>
	<%end if
	set rs=nothing
End Sub

Private Sub pcs_OrdStatus30Days(DivName)
Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,pcArr,icount
Dim Datenow,past30

	
	
	Datenow=Date()
	past30=Date()-29
	
	if scDateFrmt="DD/MM/YY" then
		tmpDateNow=Day(Datenow) & "/" & Month(Datenow) & "/" & Year(Datenow)
		tmppast30=Day(past30) & "/" & Month(past30) & "/" & Year(past30)
	else
		tmpDateNow=Month(Datenow) & "/" & Day(Datenow) & "/" & Year(Datenow)
		tmppast30=Month(past30) & "/" & Day(past30) & "/" & Year(past30)
	end if
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	query="SELECT OrderStatus,Count(*) As TotalOrders FROM Orders WHERE (Orders.OrderStatus>=2) AND Orderdate>='" & past30 & "' AND Orderdate<='" & Datenow & "' GROUP BY OrderStatus ORDER BY Count(*) DESC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		tmpline1=""
		tmpline2=""
		For i=0 to intCount
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				tmpline2=tmpline2 & ","
			end if
			tmpline1=tmpline1 & "['" & pcf_GetOrderStatusTXT(tmpArr(0,i)) & ": " & Clng(tmpArr(1,i)) & "'," & Clng(tmpArr(1,i)) & "]"
			tmpline2=tmpline2 & tmpArr(0,i)
		Next
		%>
		<script type="text/javascript" src="charts/plugins/jqplot.pieRenderer.min.js"></script>
		
		<script type=text/javascript>$pc(document).ready(function(){
		line1 = [<%=tmpline1%>];
		OrdStatusArr = [<%=tmpline2%>];
		plot2 = $pc.jqplot('<%=DivName%>', [line1], {
    	title: 'Order Status',
    	seriesDefaults:{renderer:$pc.jqplot.PieRenderer, rendererOptions:{showDataLabels: true,dataLabels: 'percent', dataLabelFormatString: '%.1f%%', sliceMargin:0}},
    	legend:{show:true}
		<%=gridOptions%>
		});
		
		$pc('#<%=DivName%>').bind('jqplotDataClick', 
            function (ev, seriesIndex, pointIndex, data) {
				var tmpURL="resultsAdvancedAll.asp?FromDate=<%=replace(tmppast30,"/","\%2F")%>&ToDate=<%=replace(tmpDateNow,"/","\%2F")%>&otype=" + OrdStatusArr[pointIndex] + "&PayType=&B1=Search+Orders"
				window.open(tmpURL,"_blank");
            }
        );
		
		$pc('#<%=DivName%>').bind('jqplotDataHighlight', 
            function (ev, seriesIndex, pointIndex, data) {
				document.getElementById("<%=DivName%>").style.cursor='pointer';
            }
        );
		
		$pc('#<%=DivName%>').bind('jqplotDataUnhighlight', 
            function (ev) {
                document.getElementById("<%=DivName%>").style.cursor='default';
            }
        );
		
		});
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
	<script type=text/javascript>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
End Sub


Private Sub pcs_NewCusts30Days(DivName)

    Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,pcArr,icount
    Dim Datenow,past30

    IF (pcvHave30Days=1) AND ((scGuestCheckoutOpt=0) OR (scGuestCheckoutOpt=1)) THEN

        Datenow=Date()
        past30=Date()-29
        
        if SQL_Format="1" then
            Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
        else
            Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
        end if
        
        if SQL_Format="1" then
            past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
        else
            past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
        end if
        
        query="SELECT pcCust_Guest, Count(*) FROM Customers WHERE pcCust_DateCreated>='" & past30 & "' AND pcCust_DateCreated<='" & Datenow & "' GROUP BY pcCust_Guest ORDER BY Count(*) DESC;"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=connTemp.execute(query)
        
        TotalCustomer=0
        if not rs.eof then
            tmpArr=rs.getRows()
            set rs=nothing
            intCount=ubound(tmpArr,2)
            tmpline1=""
            For i=0 to intCount
                if tmpline1<>"" then
                    tmpline1=tmpline1 & ","
                end if
                TotalCustomer=TotalCustomer+Clng(tmpArr(1,i))
                tmpline1=tmpline1 & "{'Name' : '" & pcf_CustTypeTXT(tmpArr(0,i)) & "', 'Amount' : " & Clng(tmpArr(1,i)) & "}"
            Next
            %>
            <script type="text/javascript" src="charts/plugins/jqplot.pieRenderer.min.js"></script>
            <script type=text/javascript>
                $pc(document).ready(function () {

                    line1 = [<%=tmpline1%>];
                    
                    $pc("#<%=DivName%>").html('');
                    chart = new cfx.Chart();            
                    chart.setDataSource(line1);
                    chart.setOptions({
                        gallery: cfx.Gallery.Pie,
                        titles: [{ text: "New Customer Registrations: <%=TotalCustomer%>"}],
                        animations: {
                            load: {
                                enabled: true
                            }
                        }
                    });
                    chart.getAllSeries().getPointLabels().setVisible(true);        
                    chart.create("<%=DivName%>")
                      
                      
                });
            </script>
            <%ChartCount=ChartCount+1
            if (ChartCount mod 2)=1 then%>
            <script type=text/javascript>
                document.getElementById("<%=DivName%>").style.clear='both';
                document.getElementById("<%=DivName%>").style.float='left';
            </script>
            <%else%>
            <script type=text/javascript>
                document.getElementById("<%=DivName%>").style.float='right';
            </script>
            <%end if%>
        <%else%>
        <script type=text/javascript>
            document.getElementById("<%=DivName%>").style.height='0px';
            document.getElementById("<%=DivName%>").style.display='none';
        </script>
        <%end if
        set rs=nothing
    ELSE
        if pcvHave30Days=1 then
            call pcs_NewCustsOnly30Days(DivName)
        else%>
        <script type=text/javascript>
            document.getElementById("<%=DivName%>").style.height='0px';
            document.getElementById("<%=DivName%>").style.display='none';
        </script>
        <%end if%>
    <%END IF
End Sub


Private Sub pcs_NewCustsOnly30Days(DivName)

    Dim past30,Datenow,rs,query,tmpArr,i,j,intCount,tmpline1,tmpline2,xname
    Dim line1(30),line2(30),line3(30),line4(30),line5(30),tmpDate

	Datenow=Date()
	past30=Date()-29
	
	For i=29 to 0 step -1
		tmpDate=Date()-i
		line1(29-i)=Day(tmpDate)
		line2(29-i)=Month(tmpDate)
		line3(29-i)=0
		if scDateFrmt="DD/MM/YY" then
			line5(29-i)=Day(tmpDate) & "/" & Month(tmpDate) & "/" & Year(date())
		else
			line5(29-i)=Month(tmpDate) & "/" & Day(tmpDate) & "/" & Year(date())
		end if
	Next
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	query="SELECT Day(pcCust_DateCreated) As TheDay,Month(pcCust_DateCreated) As TheMonth,Count(*) FROM Customers WHERE pcCust_DateCreated>='" & past30 & "' AND pcCust_DateCreated<='" & Datenow & "' GROUP BY month(pcCust_DateCreated),day(pcCust_DateCreated) ORDER BY month(pcCust_DateCreated) ASC,day(pcCust_DateCreated) ASC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	TotalCustomer=0
	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		tmpline1=""
		xname=""
		tmpline1=""
		xname=""
		For i=0 to intCount
			For j=0 to 29
			if (Cint(tmpArr(0,i))=Cint(line1(j))) AND (Cint(tmpArr(1,i))=Cint(line2(j))) then
				line3(j)=Round(tmpArr(2,i),2)
				TotalCustomer=TotalCustomer+Round(tmpArr(2,i),2)
				if scDateFrmt="DD/MM/YY" then
					line5(j)=tmpArr(0,i) & "/" & tmpArr(1,i) & "/" & Year(date())
				else
					line5(j)=tmpArr(1,i) & "/" & tmpArr(0,i) & "/" & Year(date())
				end if
				exit for
			end if
			Next
		Next
		
		For i=0 to 29
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				xname=xname & ","
			end if
			tmpline1=tmpline1 & line3(i)
			if ((i+1)=1) OR ((i+1)=30) OR ((i+1) mod 5 = 0) then
			xname=xname & "'" & line5(i) & "'"
			else
			xname=xname & "' '"
			end if
		Next
		%>
		<script type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>
		
		<script type=text/javascript>
            $pc(document).ready(function(){
                
                line1 = [<%=tmpline1%>];
                
                plot2 = $pc.jqplot('<%=DivName%>', [line1], {
                <%if ShowLegend=1 then%>
                    legend:{show:true, location:'ne', xoffset:55},
                <%end if%>			
                title:'New Customers: <%=TotalCustomer%>',
                series:[
                    {
                        renderer:$pc.jqplot.BarRenderer, 
                        rendererOptions: {
                            barWidth:8   
                        },
                        label:'New Customers',
                        pointLabels:{show:true, stackedValue: true, hideZeros:true}
                    }
                ],
                axes:{
                    xaxis:{
                        renderer:$pc.jqplot.CategoryAxisRenderer,
                        ticks: [<%=xname%>],
                        rendererOptions:{tickRenderer:$pc.jqplot.CanvasAxisTickRenderer},
                        tickOptions:{
                        fontSize:'10px', 
                        fontFamily:'Arial', 
                        angle:-30
                        }
                    },
                    yaxis:{min:0,autoscale:true, tickOptions:{formatString:'%.0f'}}
                }
                <%=gridOptions%>
                });	
            });        
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.height='0px';
			document.getElementById("<%=DivName%>").style.display='none';
		</script>
	<%end if
	set rs=nothing

End Sub

Function pcf_PricingCatName(IDCat)
Dim query,rs

	query="SELECT pcCC_Name FROM pcCustomerCategories WHERE idCustomerCategory=" & IDCat & ";"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		pcf_PricingCatName=replace(rs("pcCC_Name"),"'","\'")
	else
		pcf_PricingCatName=""
	end if
	
	set rs=nothing

End Function


Private Sub pcs_PricingCatsChart(DivName)

    Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,xvalue,pcArr,icount
    Dim Datenow,past30,line1(100),line2(100)

	query="SELECT idCustomerCategory,Count(*) FROM Customers WHERE idCustomerCategory>0 GROUP BY idCustomerCategory ORDER BY Count(*) DESC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	TotalCustomer=0
	tmpline1=""
	iCount=0
	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		
		For i=0 to intCount
			line1(icount)=Clng(tmpArr(1,i))
			line2(icount)=pcf_PricingCatName(tmpArr(0,i))
			icount=icount+1
			TotalCustomer=TotalCustomer+Clng(tmpArr(1,i))
		Next
	end if
	set rs=nothing
	
	query="SELECT customerType,Count(*) FROM Customers WHERE idCustomerCategory=0 GROUP BY customerType ORDER BY Count(*) DESC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		For i=0 to intCount
			TotalCustomer=TotalCustomer+Clng(tmpArr(1,i))
			line1(icount)=Clng(tmpArr(1,i))
			if tmpArr(0,i)="0" then
				xname="Retail"
			else
				xname="Wholesale"
			end if
			line2(icount)=xname
			icount=icount+1
		Next
	end if
	set rs=nothing
	
	if TotalCustomer>0 then
		For i=0 to icount-1
			For j=i+1 to icount-1
				if Clng(line1(i))<Clng(line1(j)) then
					xname=line1(i)
					xvalue=line2(i)
					line1(i)=line1(j)
					line2(i)=line2(j)
					line1(j)=xname
					line2(j)=xvalue
				end if
			Next
		Next
		
		tmpline1=""
		
		For i=0 to icount-1
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
			end if
			tmpline1=tmpline1 & "['" & line2(i) & ": " & Clng(line1(i)) & "'," & Clng(line1(i)) & "]"
		Next
	
		%>
		<script type="text/javascript" src="charts/plugins/jqplot.pieRenderer.min.js"></script>
		
		<script type=text/javascript>$pc(document).ready(function(){
		line1 = [<%=tmpline1%>];
		plot2 = $pc.jqplot('<%=DivName%>', [line1], {
    	title: 'Total Customers: <%=TotalCustomer%>',
    	seriesDefaults:{renderer:$pc.jqplot.PieRenderer, rendererOptions:{showDataLabels: true,dataLabels: 'percent', dataLabelFormatString: '%.1f%%',sliceMargin:0}},
    	legend:{show:true}
		<%=gridOptions%>
		});});
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
	<script type=text/javascript>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if

End Sub


Private Sub pcs_GenPrd30daysCharts(DivName,DivName1,tmpIDProduct,ShowLegend)

    Dim past30,Datenow,rs,query,tmpline1,tmpline2,tmpline3
    Dim TotalQty,TotalAmount,CurrentQty

	Datenow=Date()
	past30=Date()-29
	
	if scDateFrmt="DD/MM/YY" then
		tmpDateNow=Day(Datenow) & "/" & Month(Datenow) & "/" & Year(Datenow)
		tmppast30=Day(past30) & "/" & Month(past30) & "/" & Year(past30)
	else
		tmpDateNow=Month(Datenow) & "/" & Day(Datenow) & "/" & Year(Datenow)
		tmppast30=Month(past30) & "/" & Day(past30) & "/" & Year(past30)
	end if
	
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	query="SELECT Sum(ProductsOrdered.quantity) AS TotalQty,Sum(ProductsOrdered.quantity*ProductsOrdered.unitPrice) AS TotalAmount FROM Orders,ProductsOrdered WHERE ProductsOrdered.IDProduct=" & tmpIDProduct & " and orders.idorder=ProductsOrdered.idorder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >='" & past30 & "' AND orders.orderDate <='" & Datenow & "' GROUP BY ProductsOrdered.IDProduct;"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		TotalQty=rs("TotalQty")
		TotalAmount=rs("TotalAmount")
		set rs=nothing
		query="SELECT stock FROM Products WHERE idProduct=" & tmpIDProduct & ";"
		set rs=connTemp.execute(query)
		CurrentQty=0
		if not rs.eof then
			CurrentQty=rs("stock")
			if Clng(CurrentQty)<0 then
				CurrentQty=0
			end if
		end if
		set rs=nothing
		tmpline1="[" &  TotalAmount & ",1]"
		tmpline2="[" & TotalQty & ",2]"
		tmpline3="[" & CurrentQty & ",1]"
		%>
		
		<script type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		
		<script type=text/javascript>
        $pc(document).ready(function(){
		    line1 = [<%=tmpline1%>];
		    plot2 = $pc.jqplot('<%=DivName%>', [line1], {
			<%if ShowLegend=1 then%>
			legend:{show:true, location:'ne', xoffset:15, yoffset:220},
			<%end if%>
			title:'Quick Summary: sales in last 30 days',
			series:[
				{
					renderer:$pc.jqplot.BarRenderer, 
					rendererOptions:{barDirection:'horizontal', barWidth:5},
					pointLabels:{show:true, stackedValue: true, hideZeros:true}
				}
			],
			axes:{
				yaxis:{
					renderer:$pc.jqplot.CategoryAxisRenderer,
					ticks: ['Amount Ordered'],
					rendererOptions:{tickRenderer:$pc.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				xaxis:{min:0,autoscale:true, tickOptions:{formatString:'<%=scCurSign%>%.0f'}}
			}
			<%=gridOptions%>
		});
		
		$pc('#<%=DivName%>').bind('jqplotDataClick', 
            function (ev, seriesIndex, pointIndex, data) {
				var tmpURL="PrdsalesReport.asp?FromDate=<%=replace(tmppast30,"/","\%2F")%>&ToDate=<%=replace(tmpDateNow,"/","\%2F")%>&basedon=1&IDProduct=<%=tmpIDProduct%>&submit=Search"
				window.open(tmpURL,"_blank");
            }
        );
		
		$pc('#<%=DivName%>').bind('jqplotDataHighlight', 
            function (ev, seriesIndex, pointIndex, data) {
				document.getElementById("<%=DivName%>").style.cursor='pointer';
            }
        );
		
		$pc('#<%=DivName%>').bind('jqplotDataUnhighlight', 
            function (ev) {
                document.getElementById("<%=DivName%>").style.cursor='default';
            }
        );
		
		});	
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script type=text/javascript>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
		<script type=text/javascript>$pc(document).ready(function(){
		line2 = [<%=tmpline2%>];
		line3 = [<%=tmpline3%>];
		plot2 = $pc.jqplot('<%=DivName1%>', [line2,line3], {
			<%if ShowLegend=1 then%>
			legend:{show:true, location:'ne', xoffset:15, yoffset:220},
			<%end if%>
			title:'',
			seriesDefaults:{
				
					renderer:$pc.jqplot.BarRenderer, 
					rendererOptions:{barDirection:'horizontal', barWidth:5},
					pointLabels:{show:true, stackedValue: true, hideZeros:true}
				
			},
			axes:{
				yaxis:{
					renderer:$pc.jqplot.CategoryAxisRenderer,
					ticks: ['Current Inventory','Qty. Ordered'],
					rendererOptions:{tickRenderer:$pc.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				xaxis:{min:0,autoscale:true, tickOptions:{formatString:'%.0f'}}
			}
			<%=gridOptions%>
		});
		
		$pc('#<%=DivName1%>').bind('jqplotDataClick', 
            function (ev, seriesIndex, pointIndex, data) {
				var tmpURL="PrdsalesReport.asp?FromDate=<%=replace(tmppast30,"/","\%2F")%>&ToDate=<%=replace(tmpDateNow,"/","\%2F")%>&basedon=1&IDProduct=<%=tmpIDProduct%>&submit=Search"
				if (seriesIndex==0) window.open(tmpURL,"_blank");
            }
        );
		
		$pc('#<%=DivName1%>').bind('jqplotDataHighlight', 
            function (ev, seriesIndex, pointIndex, data) {
				if (seriesIndex==0)	document.getElementById("<%=DivName1%>").style.cursor='pointer';
            }
        );
		
		$pc('#<%=DivName1%>').bind('jqplotDataUnhighlight', 
            function (ev) {
                document.getElementById("<%=DivName1%>").style.cursor='default';
            }
        );
		
		});	
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script type=text/javascript>
			document.getElementById("<%=DivName1%>").style.clear='both';
			document.getElementById("<%=DivName1%>").style.float='left';
		</script>
		<%else%>
		<script type=text/javascript>
			document.getElementById("<%=DivName1%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
	<script type=text/javascript>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
		document.getElementById("<%=DivName1%>").style.height='0px';
		document.getElementById("<%=DivName1%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
End Sub
%>
