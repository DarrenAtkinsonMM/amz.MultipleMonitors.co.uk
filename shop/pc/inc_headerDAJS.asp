<%
if pcv_strViewPrdStyle = "computer" OR pcv_strViewPrdStyle = "blo" then
	%>
	<script language="JavaScript">
	<%'WORK OUT MOTHERBOARDS AND POWER SUPPLY COMBINATIONS
	if not InStr(pSku, "PRO1") = 0 Then
		jsPC = "PRO"
		jsMB = "B560M Chipset Motherboard"
		jsPSU = "Low Noise 500w"
		jsCaseCool = "Standard Low Noise"
		jsCPUCool = "Low Noise Air Cooler"
		jsRes = "<li>Up to 2 FHD (1920 x 1080) 60Hz Screens</li>"
		jsPorts = "<li>1 x HDMI & 1 x DVI-D Outputs</li>"
		jsSpecificPorts = "<a data-toggle=""lightbox"" data-title=""Graphics  Setup | 1 x GT 710"" href=""/pop-pages/custpc-ports-210.htm"">View Monitor Ports &amp; Resolutions For This Setup</a>"
		jsStarsSpeed = "<img src=""/images/generic/stars2-5.jpg"" />"
		jsStarsMulti = "<img src=""/images/generic/stars3.jpg"" />"
		jsStarsMulThr = "<img src=""/images/generic/stars2.jpg"" />"
		jsStarsQuiet = "<img src=""/images/generic/stars9.jpg"" />"
		jsnumQuiet = "18"
	end if  
	
	if not InStr(pSku, "ULT1") = 0 Then
		jsPC = "ULT"
		jsMB = "Fast B760 Chipset Motherboard"
		jsPSU = "BeQuiet Premium Low Noise 500w"
		jsCaseCool = "Standard Low Noise"
		jsCPUCool = "BeQuiet Ultra Quiet Air Cooler"
		jsScreens = "2 FHD (1920 x 1080)"
		jsScreensTitle = "2 Monitors Supported"
		jsRes = "<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 4 4K (3840 x 2160) 60Hz Screens</li>"
		jsPorts = "<li>1 x HDMI Output</li><li>3 x DisplayPort Outputs</li>"
		jsSpecificPorts = "<a data-toggle=""lightbox"" data-title=""Graphics  Setup | 1 x GT 1030"" href=""/pop-pages/custpc-ports-630.htm"">View Monitor Ports &amp; Resolutions For This Setup</a>"
		jsStarsSpeed = "<img src=""/images/generic/stars5-5.jpg"" />"
		jsStarsMulti = "<img src=""/images/generic/stars4.jpg"" />"
		jsStarsMulThr = "<img src=""/images/generic/stars2.jpg"" />"
		jsStarsGPU = "<img src=""/images/generic/stars4.jpg"" />"
		jsStarsQuiet = "<img src=""/images/generic/stars10.jpg"" />"
		jsnumQuiet = "20"
	end if  
	
	if not InStr(pSku, "EXT1") = 0 Then
		jsPC = "EXT"
		jsMB = "Fast Z790 Chipset Motherboard"
		jsPSU = "BeQuiet Premium Low Noise 700w"
		jsCaseCool = "Enhanced Low Noise"
		jsCPUCool = "BeQuiet Ultra Quiet Air Cooler"
		jsRes = "<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 4 4K (3840 x 2160) 60Hz Screens</li><li>Up to 4 5K (5120 x 1440) 60Hz Screens</li>"
		jsPorts = "<li>1 x HDMI Output</li><li>3 x DisplayPort Outputs</li>"
		jsStarsSpeed = "<img src=""/images/generic/stars7-5.jpg"" />"
		jsStarsMulti = "<img src=""/images/generic/stars6.jpg"" />"
		jsStarsMulThr = "<img src=""/images/generic/stars4-5.jpg"" />"
		jsStarsGPU = "<img src=""/images/generic/stars4-5.jpg"" />"
		jsStarsQuiet = "<img src=""/images/generic/stars10.jpg"" />"
		jsnumQuiet = "20"
	end if  
	%>
	var strPC = '<%=jsPC%>'
	var strMB = '<%=jsMB%>';
	var strPSU = '<%=jsPSU%>';
	var strCaseCool = '<%=jsCaseCool%>';
	var strCPUCool = '<%=jsCPUCool%>';
	var strScreens = '<%=jsScreens%>';
	var strRes = '<%=jsRes%>';
	var strPorts = '<%=jsPorts%>';
	var strSpecificPorts = '<%=jsSpecificPorts%>';
	var strScreensTitle = '<%=jsScreensTitle%>';
	var strStarsSpeed = '<%=jsStarsSpeed%>';
	var strStarsMulti = '<%=jsStarsMulti%>';
	var strStarsMulThr = '<%=jsStarsMulThr%>';
	var strStarsGPU = '<%=jsStarsGPU%>';
	var strStarsQuiet = '<%=jsStarsQuiet%>';
	var numSpeed = 0;
	var numMulti = 0;
	var numGPU = 0;
	var numQuietMaster = <%=jsnumQuiet%>;
	var numQuiet = <%=jsnumQuiet%>;
	</script>
	<%
	'DA Edit - Check for presence of querystring mid, this indicates bundle, if found then this alters some js settings
	if request.querystring("sid") = "" then
	%>
		<script language="JavaScript">
		var numGPUCost = 0;
		numGPUCost = numGPUCost.toFixed(2);
		var numStandPrice = 0;
		var numMonitorTotal = 0;
		var numMonitorAmount = 0;
		var numBunDiscount = 0;
		var baseCost = <%=pPrice/1.2%>;
		baseCost = baseCost.toFixed(2);
	
		var numSubTotal = 0;
		numSubTotal = numSubTotal.toFixed(2);
			
		var numTotal = eval(baseCost) + eval(numSubTotal);
		numTotal = numTotal.toFixed(2);
		
		var finalTotal = eval(baseCost) * 1.2;
		finalTotal = finalTotal.toFixed(2);
		
		var vatTotal = eval(finalTotal) - eval(numTotal);
		vatTotal = vatTotal.toFixed(2);
		</script>
		<script type="text/javascript" src="/js/custpc.js"></script>
		<%
	else
		if request.querystring("arr") = 1 then
			funArrayCalcs
			pMonNumber = funDAArrayCalcs
		else
			'Call function to get stand and monitor details
			funBundlesCalcs
			strFunResults = funDABundlesCalcs
			arrFunResults = Split(strFunResults,",")
			
			numMonitorTotal = arrFunResults(1) * arrFunResults(2)
			'round up 3 or 5 monitor bundles to 4 and 6 so that graphics card auto select still works
			 'Extra edit as PCs now have default to 3 monitors apart from Pro (2) and Trader (4)
			Select Case arrFunResults(2)
				Case 2
					numMonitorAmount = 2
					numBunDiscount = 25
				Case 3
	   				if jsPC = "PRO" then
	   					numMonitorAmount = 4
	   				else
	   					numMonitorAmount = 3
	   				end if
	   				numBunDiscount = 25
				Case 4
					numMonitorAmount = 4
					numBunDiscount = 50
				Case 5
					numMonitorAmount = 6
					numBunDiscount = 50
				Case 6
					numMonitorAmount = 6
					numBunDiscount = 100
				Case 8
					numMonitorAmount = 8
					numBunDiscount = 100
			End Select
			'Set txt to display in pricing table
			txtStandName = Replace(arrFunResults(4), "Monitor ", "")
			txtStandName = Replace(txtStandName, "Synergy ", "")
			txtMonitorsName = Left(arrFunResults(3),InStr(arrFunResults(3)," ")) & " Monitors x" & arrFunResults(2)
			'Build extra form elements to enable bulk add to cart
			formBundleOptions = "<input type=""hidden"" name=""idproduct2"" value=""" & request.querystring("sid") & """><input type=""hidden"" name=""QtyM" & request.querystring("sid") & """ value=""1"">" & _
								"<input type=""hidden"" name=""idproduct3"" value=""" & request.querystring("mid") & """><input type=""hidden"" name=""QtyM" & request.querystring("mid") & """ value=""" & arrFunResults(2) & """>" & _
								"<input type=""hidden"" name=""pCnt"" value=""3"">"
			
			%>
			<script language="JavaScript">
			var numGPUCost = 0;
			numGPUCost = numGPUCost.toFixed(2);
			var numStandPrice = <%=arrFunResults(0)/1.2%>;
			numStandPrice = numStandPrice.toFixed(2);
			var numMonitorTotal = <%=numMonitorTotal/1.2%>;
			numMonitorTotal = numMonitorTotal.toFixed(2);
			var numMonitorAmount = <%=numMonitorAmount%>;
			var numBunDiscount = <%=numBunDiscount%>;
			numBunDiscount = numBunDiscount.toFixed(2);
			var baseCost = <%=pPrice/1.2%>;
			baseCost = baseCost.toFixed(2);
			
			var numSubTotal = 0;
			numSubTotal = numSubTotal.toFixed(2);
				
			var numTotal = eval(baseCost) + eval(numSubTotal) + eval(numStandPrice) + eval(numMonitorTotal);
			numTotal = numTotal.toFixed(2);
			
			var finalTotal = eval(baseCost) * 1.2;
			finalTotal = finalTotal.toFixed(2);
			
			var vatTotal = eval(finalTotal) - eval(numTotal);
			vatTotal = vatTotal.toFixed(2);
			</script>
			<script type="text/javascript" src="/js/custpc.js"></script>
			<%
		end if 'if arr=1 
	end if 'if sid=""
	%>
	</head>
	<body onLoad="pageLD()" id="page-top" data-spy="scroll" data-target=".navbar-custom" itemscope itemtype="http://schema.org/WebSite">
		
		<!-- Google Tag Manager (noscript) -->
<noscript><iframe src="https://www.googletagmanager.com/ns.html?id=GTM-KRN8HZZ"
height="0" width="0" style="display:none;visibility:hidden"></iframe></noscript>
<!-- End Google Tag Manager (noscript) -->
		
	<%
elseif pcv_strViewPrdStyle = "traderpc" then
	if not request.querystring("sid") = "" then
		if request.querystring("arr") = 1 then
			funArrayCalcs
			pMonNumber = funDAArrayCalcs
		else
			'Call function to get stand and monitor details
			funBundlesCalcs
			strFunResults = funDABundlesCalcs
			arrFunResults = Split(strFunResults,",")
			
			numMonitorTotal = arrFunResults(1) * arrFunResults(2)
			'round up 3 or 5 monitor bundles to 4 and 6 so that graphics card auto select still works
			Select Case arrFunResults(2)
				Case 2
					numMonitorAmount = 3
					numBunDiscount = 25
				Case 3
					numMonitorAmount = 3
					numBunDiscount = 25
				Case 4
					numMonitorAmount = 4
					numBunDiscount = 50
				Case 5
					numMonitorAmount = 6
					numBunDiscount = 50
				Case 6
					numMonitorAmount = 6
					numBunDiscount = 100
				Case 8
					numMonitorAmount = 8
					numBunDiscount = 100
			End Select
			'Set txt to display in pricing table
			txtStandName = Replace(arrFunResults(4), "Monitor ", "")
			txtStandName = Replace(txtStandName, "Synergy ", "")
			txtMonitorsName = Left(arrFunResults(3),InStr(arrFunResults(3)," ")) & " Monitors x" & arrFunResults(2)
			'Build extra form elements to enable bulk add to cart
			formBundleOptions = "<input type=""hidden"" name=""idproduct2"" value=""" & request.querystring("sid") & """><input type=""hidden"" name=""QtyM" & request.querystring("sid") & """ value=""1"">" & _
								"<input type=""hidden"" name=""idproduct3"" value=""" & request.querystring("mid") & """><input type=""hidden"" name=""QtyM" & request.querystring("mid") & """ value=""" & arrFunResults(2) & """>" & _
								"<input type=""hidden"" name=""pCnt"" value=""3"">"
			
			%>
			<script language="JavaScript">
			var numGPUCost = 0;
			numGPUCost = numGPUCost.toFixed(2);
			var numStandPrice = <%=arrFunResults(0)/1.2%>;
			numStandPrice = numStandPrice.toFixed(2);
			var numMonitorTotal = <%=numMonitorTotal/1.2%>;
			numMonitorTotal = numMonitorTotal.toFixed(2);
			var numMonitorAmount = <%=numMonitorAmount%>;
			var numBunDiscount = <%=numBunDiscount%>;
			numBunDiscount = numBunDiscount.toFixed(2);
			var baseCost = <%=pPrice/1.2%>;
			baseCost = baseCost.toFixed(2);
			
			var numSubTotal = 0;
			numSubTotal = numSubTotal.toFixed(2);
				
			var numTotal = eval(baseCost) + eval(numSubTotal) + eval(numStandPrice) + eval(numMonitorTotal);
			numTotal = numTotal.toFixed(2);
			
			var finalTotal = eval(baseCost) * 1.2;
			finalTotal = finalTotal.toFixed(2);
			
			var vatTotal = eval(finalTotal) - eval(numTotal);
			vatTotal = vatTotal.toFixed(2);
			</script>
			<script type="text/javascript" src="/js/custpc-traderpc.js"></script>
			<%
		end if 'if not arr=1
	else
		%>
		<script language="JavaScript">
		var numGPUCost = 0;
		numGPUCost = numGPUCost.toFixed(2);
		var numStandPrice = 0;
		numStandPrice = numStandPrice.toFixed(2);
		var numMonitorTotal = 0;
		numMonitorTotal = numMonitorTotal.toFixed(2);
		var numMonitorAmount = 0;
		var numBunDiscount = 0;
		numBunDiscount = numBunDiscount.toFixed(2);
		
		var baseCost = <%=pPrice/1.2%>;
		baseCost = baseCost.toFixed(2);
		
		var numSubTotal = 0;
		numSubTotal = numSubTotal.toFixed(2);
		
		var numTotal = eval(baseCost) + eval(numSubTotal);
		numTotal = numTotal.toFixed(2);
		
		var vatTotal = 0;
		</script>
		<script type="text/javascript" src="/js/custpc-traderpc.js"></script>
		<%
	end if 'if not sid=""
	%>
	</head>
	<body onLoad="pageLD()" id="page-top" data-spy="scroll" data-target=".navbar-custom" itemscope itemtype="http://schema.org/WebSite">
		
		<!-- Google Tag Manager (noscript) -->
<noscript><iframe src="https://www.googletagmanager.com/ns.html?id=GTM-KRN8HZZ"
height="0" width="0" style="display:none;visibility:hidden"></iframe></noscript>
<!-- End Google Tag Manager (noscript) -->
		
	<%
elseif pcv_strViewPrdStyle = "traderpropc" then
	if not request.querystring("sid") = "" then
		if request.querystring("arr") = 1 then
			funArrayCalcs
			pMonNumber = funDAArrayCalcs
		else
			'Call function to get stand and monitor details
			funBundlesCalcs
			strFunResults = funDABundlesCalcs
			arrFunResults = Split(strFunResults,",")
			
			numMonitorTotal = arrFunResults(1) * arrFunResults(2)
			'round up 3 or 5 monitor bundles to 4 and 6 so that graphics card auto select still works
			Select Case arrFunResults(2)
				Case 2
					numMonitorAmount = 4
					numBunDiscount = 25
				Case 3
					numMonitorAmount = 4
					numBunDiscount = 25
				Case 4
					numMonitorAmount = 4
					numBunDiscount = 50
				Case 5
					numMonitorAmount = 6
					numBunDiscount = 50
				Case 6
					numMonitorAmount = 6
					numBunDiscount = 100
				Case 8
					numMonitorAmount = 8
					numBunDiscount = 100
			End Select
			'Set txt to display in pricing table
			txtStandName = Replace(arrFunResults(4), "Monitor ", "")
			txtStandName = Replace(txtStandName, "Synergy ", "")
			txtMonitorsName = Left(arrFunResults(3),InStr(arrFunResults(3)," ")) & " Monitors x" & arrFunResults(2)
			'Build extra form elements to enable bulk add to cart
			formBundleOptions = "<input type=""hidden"" name=""idproduct2"" value=""" & request.querystring("sid") & """><input type=""hidden"" name=""QtyM" & request.querystring("sid") & """ value=""1"">" & _
								"<input type=""hidden"" name=""idproduct3"" value=""" & request.querystring("mid") & """><input type=""hidden"" name=""QtyM" & request.querystring("mid") & """ value=""" & arrFunResults(2) & """>" & _
								"<input type=""hidden"" name=""pCnt"" value=""3"">"
			
			%>
			<script language="JavaScript">
			var numGPUCost = 0;
			numGPUCost = numGPUCost.toFixed(2);
			var numStandPrice = <%=arrFunResults(0)/1.2%>;
			numStandPrice = numStandPrice.toFixed(2);
			var numMonitorTotal = <%=numMonitorTotal/1.2%>;
			numMonitorTotal = numMonitorTotal.toFixed(2);
			var numMonitorAmount = <%=numMonitorAmount%>;
			var numBunDiscount = <%=numBunDiscount%>;
			numBunDiscount = numBunDiscount.toFixed(2);
			var baseCost = <%=pPrice/1.2%>;
			baseCost = baseCost.toFixed(2);
			
			var numSubTotal = 0;
			numSubTotal = numSubTotal.toFixed(2);
				
			var numTotal = eval(baseCost) + eval(numSubTotal) + eval(numStandPrice) + eval(numMonitorTotal);
			numTotal = numTotal.toFixed(2);
			
			var finalTotal = eval(baseCost) * 1.2;
			finalTotal = finalTotal.toFixed(2);
			
			var vatTotal = eval(finalTotal) - eval(numTotal);
			vatTotal = vatTotal.toFixed(2);
			</script>
			<script type="text/javascript" src="/js/custpc-traderpropc.js"></script>
			<%
		end if 'if not arr=1
	else
		%>
		<script language="JavaScript">
		var numGPUCost = 0;
		numGPUCost = numGPUCost.toFixed(2);
		var numStandPrice = 0;
		numStandPrice = numStandPrice.toFixed(2);
		var numMonitorTotal = 0;
		numMonitorTotal = numMonitorTotal.toFixed(2);
		var numMonitorAmount = 0;
		var numBunDiscount = 0;
		numBunDiscount = numBunDiscount.toFixed(2);
		
		var baseCost = <%=pPrice/1.2%>;
		baseCost = baseCost.toFixed(2);
		
		var numSubTotal = 0;
		numSubTotal = numSubTotal.toFixed(2);
		
		var numTotal = eval(baseCost) + eval(numSubTotal);
		numTotal = numTotal.toFixed(2);
		
		var vatTotal = 0;
		</script>
		<script type="text/javascript" src="/js/custpc-traderpropc.js"></script>
		<%
	end if 'if not sid=""
	%>
	</head>
	<body onLoad="pageLD()" id="page-top" data-spy="scroll" data-target=".navbar-custom" itemscope itemtype="http://schema.org/WebSite">
		
		<!-- Google Tag Manager (noscript) -->
<noscript><iframe src="https://www.googletagmanager.com/ns.html?id=GTM-KRN8HZZ"
height="0" width="0" style="display:none;visibility:hidden"></iframe></noscript>
<!-- End Google Tag Manager (noscript) -->
		
	<%
	elseif pcv_strViewPrdStyle = "charterpc" then
	if not request.querystring("sid") = "" then
		if request.querystring("arr") = 1 then
			funArrayCalcs
			pMonNumber = funDAArrayCalcs
		else
			'Call function to get stand and monitor details
			funBundlesCalcs
			strFunResults = funDABundlesCalcs
			arrFunResults = Split(strFunResults,",")
			
			numMonitorTotal = arrFunResults(1) * arrFunResults(2)
			'round up 3 or 5 monitor bundles to 4 and 6 so that graphics card auto select still works
			Select Case arrFunResults(2)
				Case 2
					numMonitorAmount = 3
					numBunDiscount = 25
				Case 3
					numMonitorAmount = 3
					numBunDiscount = 25
				Case 4
					numMonitorAmount = 4
					numBunDiscount = 50
				Case 5
					numMonitorAmount = 6
					numBunDiscount = 50
				Case 6
					numMonitorAmount = 6
					numBunDiscount = 100
				Case 8
					numMonitorAmount = 8
					numBunDiscount = 100
			End Select
			'Set txt to display in pricing table
			txtStandName = Replace(arrFunResults(4), "Monitor ", "")
			txtStandName = Replace(txtStandName, "Synergy ", "")
			txtMonitorsName = Left(arrFunResults(3),InStr(arrFunResults(3)," ")) & " Monitors x" & arrFunResults(2)
			'Build extra form elements to enable bulk add to cart
			formBundleOptions = "<input type=""hidden"" name=""idproduct2"" value=""" & request.querystring("sid") & """><input type=""hidden"" name=""QtyM" & request.querystring("sid") & """ value=""1"">" & _
								"<input type=""hidden"" name=""idproduct3"" value=""" & request.querystring("mid") & """><input type=""hidden"" name=""QtyM" & request.querystring("mid") & """ value=""" & arrFunResults(2) & """>" & _
								"<input type=""hidden"" name=""pCnt"" value=""3"">"
			
			%>
			<script language="JavaScript">
			var numGPUCost = 0;
			numGPUCost = numGPUCost.toFixed(2);
			var numStandPrice = <%=arrFunResults(0)/1.2%>;
			numStandPrice = numStandPrice.toFixed(2);
			var numMonitorTotal = <%=numMonitorTotal/1.2%>;
			numMonitorTotal = numMonitorTotal.toFixed(2);
			var numMonitorAmount = <%=numMonitorAmount%>;
			var numBunDiscount = <%=numBunDiscount%>;
			numBunDiscount = numBunDiscount.toFixed(2);
			var baseCost = <%=pPrice/1.2%>;
			baseCost = baseCost.toFixed(2);
			
			var numSubTotal = 0;
			numSubTotal = numSubTotal.toFixed(2);
				
			var numTotal = eval(baseCost) + eval(numSubTotal) + eval(numStandPrice) + eval(numMonitorTotal);
			numTotal = numTotal.toFixed(2);
			
			var finalTotal = eval(baseCost) * 1.2;
			finalTotal = finalTotal.toFixed(2);
			
			var vatTotal = eval(finalTotal) - eval(numTotal);
			vatTotal = vatTotal.toFixed(2);
			</script>
			<script type="text/javascript" src="/js/custpc-charterpc.js"></script>
			<%
		end if 'if arr=1
	else
		%>
		<script language="JavaScript">
		var numGPUCost = 0;
		numGPUCost = numGPUCost.toFixed(2);
		var numStandPrice = 0;
		numStandPrice = numStandPrice.toFixed(2);
		var numMonitorTotal = 0;
		numMonitorTotal = numMonitorTotal.toFixed(2);
		var numMonitorAmount = 0;
		var numBunDiscount = 0;
		numBunDiscount = numBunDiscount.toFixed(2);
		
		var baseCost = <%=pPrice/1.2%>;
		baseCost = baseCost.toFixed(2);
		
		var numSubTotal = 0;
		numSubTotal = numSubTotal.toFixed(2);
		
		var numTotal = eval(baseCost) + eval(numSubTotal);
		numTotal = numTotal.toFixed(2);
		
		var vatTotal = 0;
		</script>
		<script type="text/javascript" src="/js/custpc-charterpc.js"></script>
		<%
	end if ' if not sid=""
	%>
	</head>
	<body onLoad="pageLD()" id="page-top" data-spy="scroll" data-target=".navbar-custom" itemscope itemtype="http://schema.org/WebSite">
		
		<!-- Google Tag Manager (noscript) -->
<noscript><iframe src="https://www.googletagmanager.com/ns.html?id=GTM-KRN8HZZ"
height="0" width="0" style="display:none;visibility:hidden"></iframe></noscript>
<!-- End Google Tag Manager (noscript) -->
		
	<%
else
	%>
    <meta name="ahrefs-site-verification" content="8470152a7ff00d441a9b23b8b213125877ad0fc23f582b6b8a6c3c5416e91ec6">
	</head>
	<body itemscope itemtype="http://schema.org/WebSite">
		<!-- Google Tag Manager (noscript) -->
<noscript><iframe src="https://www.googletagmanager.com/ns.html?id=GTM-KRN8HZZ"
height="0" width="0" style="display:none;visibility:hidden"></iframe></noscript>
<!-- End Google Tag Manager (noscript) -->
	<%
end if
%>