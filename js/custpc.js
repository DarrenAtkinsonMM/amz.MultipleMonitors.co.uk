function pageLD() {
 	var ddlop1 = document.getElementById('idOption1');
	if (ddlop1 != null)
	{
	for (var i = 0; i < ddlop1.options.length; i++) {
		if (ddlop1.options[i].selected == true) {
			numSubTotal = eval(ddlop1.options[i].title);
		}
	}
	}
	
	var ddlop2 = document.getElementById('idOption2');
	if (ddlop2 != null)
	{
	for (var i = 0; i < ddlop2.options.length; i++) {
		if (ddlop2.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop2.options[i].title);
		}
	}
	}

	var ddlop3 = document.getElementById('idOption3');
	if (ddlop3 != null)
	{
	for (var i = 0; i < ddlop3.options.length; i++) {
		if (ddlop3.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop3.options[i].title);
		}
	}
	}

	var ddlop4 = document.getElementById('idOption4');
	if (ddlop4 != null)
	{
	for (var i = 0; i < ddlop4.options.length; i++) {
		if (ddlop4.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop4.options[i].title);
		}
	}
	}

	var ddlop5 = document.getElementById('idOption5');
	if (ddlop5 != null)
	{
	for (var i = 0; i < ddlop5.options.length; i++) {
		if (ddlop5.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop5.options[i].title);
		}
	}
	}

	var ddlop6 = document.getElementById('idOption6');
	if (ddlop6 != null)
	{
	for (var i = 0; i < ddlop6.options.length; i++) {
		if (ddlop6.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop6.options[i].title);
		}
	}
	}

	var ddlop7 = document.getElementById('idOption7');
	if (ddlop7 != null)
	{
	for (var i = 0; i < ddlop7.options.length; i++) {
		if (ddlop7.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop7.options[i].title);
		}
	}
	}

	var ddlop8 = document.getElementById('idOption8');
	if (ddlop8 != null)
	{
	for (var i = 0; i < ddlop8.options.length; i++) {
		if (ddlop8.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop8.options[i].title);
		}
	}
	}

	var ddlop9 = document.getElementById('idOption9');
	if (ddlop9 != null)
	{
	for (var i = 0; i < ddlop9.options.length; i++) {
		if (ddlop9.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop9.options[i].title);
		}
	}
	}

	var ddlop10 = document.getElementById('idOption10');
	if (ddlop10 != null)
	{
	for (var i = 0; i < ddlop10.options.length; i++) {
		if (ddlop10.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop10.options[i].title);
		}
	}
	}

	var ddlop11 = document.getElementById('idOption11');
	if (ddlop11 != null)
	{
	for (var i = 0; i < ddlop11.options.length; i++) {
		if (ddlop11.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop11.options[i].title);
		}
	}
	}

	var ddlop12 = document.getElementById('idOption12');
	if (ddlop12 != null)
	{
	for (var i = 0; i < ddlop12.options.length; i++) {
		if (ddlop12.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop12.options[i].title);
		}
	}
	}

	var ddlop13 = document.getElementById('idOption13');
	if (ddlop13 != null)
	{
	for (var i = 0; i < ddlop13.options.length; i++) {
		if (ddlop13.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop13.options[i].title);
		}
	}
	}

	var ddlop14 = document.getElementById('idOption14');
	if (ddlop14 != null)
	{
	for (var i = 0; i < ddlop14.options.length; i++) {
		if (ddlop14.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop14.options[i].title);
		}
	}
	}

	numSubTotal = numSubTotal.toFixed(2);

	numTotal = eval(numSubTotal) + eval(baseCost) + eval(numStandPrice) + eval(numMonitorTotal) - eval(numBunDiscount);
	numTotal = numTotal.toFixed(2);
	
	finalTotal = eval(numTotal) * 1.2;
	finalTotal = finalTotal.toFixed(2);
	
	vatTotal = eval(finalTotal) - eval(numTotal);
	vatTotal = vatTotal.toFixed(2);

	if(numMonitorTotal != 0){
		document.getElementById('standPrice').innerHTML=numStandPrice;
		document.getElementById('monitorsPrice').innerHTML=numMonitorTotal;
		document.getElementById('bunDiscount').innerHTML=numBunDiscount;
		document.getElementById('gpuPrice').innerHTML=numGPUCost;
	}
    document.getElementById('basePrice').innerHTML=baseCost;
	document.getElementById('extrasPrice').innerHTML=numSubTotal;
	document.getElementById('subtotalPrice').innerHTML=numTotal;
	document.getElementById('vatPrice').innerHTML=vatTotal;
	document.getElementById('finalPrice').innerHTML=finalTotal;
	document.getElementById('optMotherboard').innerHTML=strMB;
	document.getElementById('optPSU').innerHTML=strPSU;
	document.getElementById('optCaseCool').innerHTML=strCaseCool;
	document.getElementById('optCPUCool').innerHTML=strCPUCool;
	//document.getElementById('optScreens').innerHTML=strScreens;
	document.getElementById('optRes').innerHTML=strRes;
	document.getElementById('optPorts').innerHTML=strPorts;
	document.getElementById('optScreensTitle').innerHTML='Monitors &amp; Resolutions';
	document.getElementById('stars-speed').innerHTML=strStarsSpeed;
	document.getElementById('stars-multi').innerHTML=strStarsMulti;
	document.getElementById('stars-mulThr').innerHTML=strStarsMulThr;
	document.getElementById('stars-gpu').innerHTML=strStarsGPU;
	document.getElementById('stars-quiet').innerHTML=strStarsQuiet;
	//document.getElementById('optSpecificPorts').innerHTML=strSpecificPorts;
	
	
	//Run fancybox again due to newly assigned css selector
	$(document).ready(function() {
		//If we have more than 2 bundled monitor then change graphics option
		if (numMonitorAmount > 2) {
		
		//Make 7 select the 8 monitor option
		if (numMonitorAmount > 4) { numMonitorAmount = 8;}
			//Fix to make 5 or 6 mons pick 8 option due to removal of 6 monitor PC option on Ultra's only
			
		var ddlAutoSelectTxt = 'Up to ' + numMonitorAmount;
		
		$("#idOption5").find("option:contains('" + ddlAutoSelectTxt +"')").each(function () {
   			$(this).attr("selected", "selected");return false;
		});
		reCalc();
		reCalcColour();
		}
		//Hide graphics adapters row for all configs with less than 7 screens
		//if (numMonitorAmount < 7) {
			//document.getElementById('trGA').style.display = "none";
		//}
		});

	}

	function reCalc() {
	
	numSpeed = 0;
	numMulti = 0;
	numQuiet = numQuietMaster;
	$("#optMotherboard").css('color','#000000');
	$("#optPSU").css('color','#000000');
	$("#optCaseCool").css('color','#000000');
	
	var ddlop1 = document.getElementById('idOption1');
	if (ddlop1 != null)
	{
	for (var i = 0; i < ddlop1.options.length; i++) {
		if (ddlop1.options[i].selected == true) {
			numSubTotal = eval(ddlop1.options[i].title);
			switch(ddlop1.options[i].id) {
				//START - 2021 Major CPU Update	
				case '337':
					//'14100F
					numSpeed = 11;
					numMulti = 8;
					numMulThr = 4;
					numQuiet = numQuiet;
					document.getElementById('optMotherboard').innerHTML='Fast B760 Chipset Motherboard';
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;
				case '338':
					//'14400F
					numSpeed = 11;
					numMulti = 10;
					numMulThr = 7;
					numQuiet = numQuiet;
					document.getElementById('optMotherboard').innerHTML='Fast B760 Chipset Motherboard';
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;
				case '342':
					//'14600KF
					numSpeed = 14;
					numMulti = 13;
					numMulThr = 11;
					numQuiet = numQuiet;
					if (strPC == "ULT") {
						document.getElementById('optMotherboard').innerHTML='Fast B760 Chipset Motherboard';
					} else {
						document.getElementById('optMotherboard').innerHTML='Fast Z790 Chipset Motherboard';
					}
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;	
				case '345':
					//'14600KF on ultra
					numSpeed = 14;
					numMulti = 13;
					numMulThr = 11;
					numQuiet = numQuiet;
					if (strPC == "ULT") {
						document.getElementById('optMotherboard').innerHTML='Fast B760 Chipset Motherboard';
					} else {
						document.getElementById('optMotherboard').innerHTML='Fast Z790 Chipset Motherboard';
					}
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;	
				case '343':
					//'14700KF
					numSpeed = 15;
					numMulti = 18;
					numMulThr = 15;
					numQuiet = numQuiet;
					if (strPC == "ULT") {
						document.getElementById('optMotherboard').innerHTML='Fast B760 Chipset Motherboard';
					} else {
						document.getElementById('optMotherboard').innerHTML='Fast Z790 Chipset Motherboard';
					}
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;	
				case '344':
					//'14900KF
					numSpeed = 17;
					numMulti = 20;
					numMulThr = 17;
					numQuiet = numQuiet - 2;
					if (strPC == "ULT") {
						document.getElementById('optMotherboard').innerHTML='Fast B760 Chipset Motherboard';
					} else {
						document.getElementById('optMotherboard').innerHTML='Fast Z790 Chipset Motherboard';
					}
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;	
				case '356':
					//'225F
					numSpeed = 15;
					numMulti = 12;
					numMulThr = 9;
					numQuiet = numQuiet;
					if (strPC == "ULT") {
						document.getElementById('optMotherboard').innerHTML='Fast B760 Chipset Motherboard';
					} else {
						document.getElementById('optMotherboard').innerHTML='Fast Z890 Chipset Motherboard';
					}
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;	
				case '351':
					//'245KF
					numSpeed = 17;
					numMulti = 14;
					numMulThr = 12;
					numQuiet = numQuiet;
					if (strPC == "ULT") {
						document.getElementById('optMotherboard').innerHTML='Fast B760 Chipset Motherboard';
					} else {
						document.getElementById('optMotherboard').innerHTML='Fast Z890 Chipset Motherboard';
					}
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;	
				case '350':
					//'265KF
					numSpeed = 18;
					numMulti = 20;
					numMulThr = 17;
					numQuiet = numQuiet;
					if (strPC == "ULT") {
						document.getElementById('optMotherboard').innerHTML='Fast B760 Chipset Motherboard';
					} else {
						document.getElementById('optMotherboard').innerHTML='Fast Z890 Chipset Motherboard';
					}
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;	
				case '349':
					//'285KF
					numSpeed = 20;
					numMulti = 20;
					numMulThr = 20;
					numQuiet = numQuiet - 2;
					if (strPC == "ULT") {
						document.getElementById('optMotherboard').innerHTML='Fast B760 Chipset Motherboard';
					} else {
						document.getElementById('optMotherboard').innerHTML='Fast Z890 Chipset Motherboard';
					}
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;	
				case '347':
					//'9600x
					numSpeed = 16;
					numMulti = 11;
					numMulThr = 8;
					numQuiet = numQuiet ;
					document.getElementById('optMotherboard').innerHTML='Fast X870 Chipset Motherboard';
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;	
				case '348':
					//'9700x
					numSpeed = 16;
					numMulti = 13;
					numMulThr = 10;
					numQuiet = numQuiet - 1;
					document.getElementById('optMotherboard').innerHTML='Fast X870 Chipset Motherboard';
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;	
				case '317':
					//'9900x
					numSpeed = 17;
					numMulti = 20;
					numMulThr = 16;
					numQuiet = numQuiet - 3;
					document.getElementById('optMotherboard').innerHTML='Fast X870 Chipset Motherboard';
					document.getElementById('optCPUCool').innerHTML='Corsair Liquid AIO Cooler';
					break;	
				case '318':
					//'9950x
					numSpeed = 17;
					numMulti = 20;
					numMulThr = 19;
					numQuiet = numQuiet - 3;
					document.getElementById('optMotherboard').innerHTML='Fast X870 Chipset Motherboard';
					document.getElementById('optCPUCool').innerHTML='Corsair Liquid AIO Cooler';
					break;	
				case '355':
					//'9950X3D
					numSpeed = 17;
					numMulti = 20;
					numMulThr = 20;
					numQuiet = numQuiet ;
					document.getElementById('optMotherboard').innerHTML='Fast X870 Chipset Motherboard';
					document.getElementById('optCPUCool').innerHTML='BeQuiet Ultra Quiet Air Cooler';
					break;	
				//END - 2021 Major CPU Update
				default:
					numSpeed = 0;
					numMulti = 0;
					break;
			}

		}
	}
	}
	
	var ddlop2 = document.getElementById('idOption2');
	if (ddlop2 != null)
	{
	for (var i = 0; i < ddlop2.options.length; i++) {
		if (ddlop2.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop2.options[i].title);
			switch(ddlop2.options[i].id) {
				case '235':
					//'8GB(Ultra)
					numMulti = numMulti - 2;
					break;
				case '101':
					//'16GB(Ultra)
					numMulti = numMulti;
					break;
				case '102':
					//'32GB(Ultra)
					numMulti = numMulti;
					break;
				case '129':
					//'64GB(Ultra)
					numMulti = numMulti;
					break;
				case '79':
					//'16GB Ultra PC
					//numMulti = numMulti + 1;
					//'DA Calc to force multi-tasking 1 point lower if it's over 10
					//if (numMulti > 10) {
					//	numMulti = 10;
					//}
					break;
				case '128':
					//'32GB Ultra PC
					//numMulti = numMulti + 1;
					//'DA Calc to force multi-tasking 1 point lower if it's over 10
					//if (numMulti > 10) {
					//	numMulti = 10;
					//}
					break;
				case '221':
					//'64GB Ultra PC
					//numMulti = numMulti + 1;
					//'DA Calc to force multi-tasking 1 point lower if it's over 10
					//if (numMulti > 10) {
					//	numMulti = 10;
					//}
					break;
				case '101':
					//'17GB Extreme PC
					break;
				case '102':
					//'32GB Extreme PC
					break;
				case '129':
					//'64GB Extreme PC
					break;
				default:
					break;
			}

		}
	}
	}

	var ddlop3 = document.getElementById('idOption3');
	if (ddlop3 != null)
	{
	for (var i = 0; i < ddlop3.options.length; i++) {
		if (ddlop3.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop3.options[i].title);
			switch(ddlop3.options[i].id) {
				default:
					break;
			}

		}
	}
	}

	var ddlop4 = document.getElementById('idOption4');
	if (ddlop4 != null)
	{
	for (var i = 0; i < ddlop4.options.length; i++) {
		if (ddlop4.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop4.options[i].title);
		}
	}
	}

	var ddlop5 = document.getElementById('idOption5');
	if (ddlop5 != null)
	{
	for (var i = 0; i < ddlop5.options.length; i++) {
		if (ddlop5.options[i].selected == true) {
			if(numMonitorTotal != 0){
				numGPUCost = eval(ddlop5.options[i].title);
			}
			else {
				numSubTotal = eval(numSubTotal) + eval(ddlop5.options[i].title);
			}
			switch(ddlop5.options[i].id) {
					//START: 2021 Major GPU Update
				
				case '331':
					//'RTX A2000 x1
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 500w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 750w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 4 4K (3840 x 2160) 120Hz Screens</li><li>Up to 4 5K (5120 x 2880) 60Hz Screens</li><li>Up to 2 8K (7680 x 4320) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>4 x Mini-DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='107';
					numGPU = 14;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '332':
					//'RTX A2000 x2
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 600w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 750w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 8 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 8 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 8 4K (3840 x 2160) 60Hz Screens</li><li>Up to 4 5K (5120 x 2880) 60Hz Screens</li><li>Up to 2 8K (7680 x 4320) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>8 x Mini-DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='107';
					numGPU = 14;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '330':
					//'A400 x3
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 600w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 750w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 12 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 12 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 7 4K (3840 x 2160) 60Hz Screens</li><li>Up to 4 5K (5120 x 2880) 60Hz Screens</li><li>Up to 2 8K (7680 x 4320) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>12 x Mini-DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='43';
					numGPU = 8;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '333':
					//'A380 x1
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 500w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 750w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 4 4K (3840 x 2160) 60Hz Screens</li><li>Up to 4 5K (5120 x 1440) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>1 x HDMI Output</li><li>3 x DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='66';
					numGPU = 9;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '335':
					//'A380 & Intel UHD
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 500w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 750w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 6 FHD (1920 x 1080) 60Hz Screens</li><li>Up to 6 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 4 4K (3840 x 2160) 60Hz Screens</li><li>Up to 4 5K (5120 x 1440) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>2 x HDMI Outputs</li><li>4 x DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='66';
					numGPU = 9;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '334':
					//'A380 x2 
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 500w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 750w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 8 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 8 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 6 4K (3840 x 2160) 60Hz Screens</li><li>Up to 6 5K (5120 x 1440) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>2 x HDMI Outputs</li><li>6 x DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='66';
					numGPU = 9;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '336':
					//'A380 x2 & Intel UHD
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 600w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 750w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 10 FHD (1920 x 1080) 60Hz Screens</li><li>Up to 10 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 6 4K (3840 x 2160) 60Hz Screens</li><li>Up to 6 5K (5120 x 1440) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>3 x HDMI Outputs</li><li>7 x DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='66';
					numGPU = 9;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '352':
					//'5050
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 600w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 750w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 4 4K (3840 x 2160) 120Hz Screens</li><li>Up to 4 5K (5120 x 1440) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>1 x HDMI Outputs</li><li>3 x DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='421';
					numGPU = 17;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '353':
					//'5050 x2
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 600w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 750w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 8 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 8 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 8 4K (3840 x 2160) 60Hz Screens</li><li>Up to 4 5K (5120 x 1440) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>2 x HDMI Outputs</li><li>6 x DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='421';
					numGPU = 17;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '354':
					//'4080 Super
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 600w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 850w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 4 4K (3840 x 2160) 120Hz Screens</li><li>Up to 4 5K (5120 x 1440) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>1 x HDMI Outputs</li><li>3 x DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='836';
					numGPU = 20;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '357':
					//'5060 
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 600w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 850w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 4 4K (3840 x 2160) 120Hz Screens</li><li>Up to 4 5K (5120 x 1440) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>1 x HDMI Outputs</li><li>3 x DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='614';
					numGPU = 20;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '358':
					//'5060 x2
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 600w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 850w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 8 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 8 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 8 4K (3840 x 2160) 60Hz Screens</li><li>Up to 4 5K (5120 x 1440) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>2 x HDMI Outputs</li><li>6 x DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='614';
					numGPU = 20;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '359':
					//'5070 
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 600w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 850w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 4 4K (3840 x 2160) 120Hz Screens</li><li>Up to 4 5K (5120 x 1440) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>1 x HDMI Outputs</li><li>3 x DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='988';
					numGPU = 20;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '360':
					//'5070 
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 600w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 1000w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 4 4K (3840 x 2160) 120Hz Screens</li><li>Up to 4 5K (5120 x 1440) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>1 x HDMI Outputs</li><li>3 x DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='1801';
					numGPU = 20;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '361':
					//'RTX A400 x1
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 500w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 750w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 4 FHD (1920 x 1080) 60Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 4 4K (3840 x 2160) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>4 x Mini-DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='43';
					numGPU = 8;
					//document.getElementById('trGA').style.display = "none";
					break;
				case '362':
					//'RTX A400 x2
					if (strPC == "ULT") {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 500w';
					} else {
						document.getElementById('optPSU').innerHTML='BeQuiet Premium Low Noise 750w';
					}
					document.getElementById('optCaseCool').innerHTML='Enhanced Low Noise';
					document.getElementById('optRes').innerHTML='<li>Up to 8 FHD (1920 x 1080) 60Hz Screens</li><li>Up to 8 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 4 4K (3840 x 2160) 60Hz Screens</li>';
					document.getElementById('optPorts').innerHTML='<li>8 x Mini-DisplayPort Outputs</li>';
					document.getElementById('stars-gputops').innerHTML='43';
					numGPU = 8;
					//document.getElementById('trGA').style.display = "none";
					break;
				default:
					document.getElementById('optMotherboard').innerHTML=strMB;
					document.getElementById('optCaseCool').innerHTML=strCaseCool;
					document.getElementById('optCPUCool').innerHTML=strCPUCool;
					//document.getElementById('optScreens').innerHTML=strScreens;
					document.getElementById('optPorts').innerHTML=strPorts;
					document.getElementById('optScreensTitle').innerHTML=strScreensTitle;
					numGPU = 10;
					break;
			}
		}
	}
	}

	var ddlop6 = document.getElementById('idOption6');
	if (ddlop6 != null)
	{
	for (var i = 0; i < ddlop6.options.length; i++) {
		if (ddlop6.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop6.options[i].title);
		}
	}
	}

	var ddlop7 = document.getElementById('idOption7');
	if (ddlop7 != null)
	{
	for (var i = 0; i < ddlop7.options.length; i++) {
		if (ddlop7.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop7.options[i].title);
		}
	}
	}

	var ddlop8 = document.getElementById('idOption8');
	if (ddlop8 != null)
	{
	for (var i = 0; i < ddlop8.options.length; i++) {
		if (ddlop8.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop8.options[i].title);
		}
	}
	}

	var ddlop9 = document.getElementById('idOption9');
	if (ddlop9 != null)
	{
	for (var i = 0; i < ddlop9.options.length; i++) {
		if (ddlop9.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop9.options[i].title);
		}
	}
	}

	var ddlop10 = document.getElementById('idOption10');
	if (ddlop10 != null)
	{
	for (var i = 0; i < ddlop10.options.length; i++) {
		if (ddlop10.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop10.options[i].title);
		}
	}
	}

	var ddlop11 = document.getElementById('idOption11');
	if (ddlop11 != null)
	{
	for (var i = 0; i < ddlop11.options.length; i++) {
		if (ddlop11.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop11.options[i].title);
		}
	}
	}

	var ddlop12 = document.getElementById('idOption12');
	if (ddlop12 != null)
	{
	for (var i = 0; i < ddlop12.options.length; i++) {
		if (ddlop12.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop12.options[i].title);
		}
	}
	}

	var ddlop13 = document.getElementById('idOption13');
	if (ddlop13 != null)
	{
	for (var i = 0; i < ddlop13.options.length; i++) {
		if (ddlop13.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop13.options[i].title);
		}
	}
	}

	var ddlop14 = document.getElementById('idOption14');
	if (ddlop14 != null)
	{
	for (var i = 0; i < ddlop14.options.length; i++) {
		if (ddlop14.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop14.options[i].title);
		}
	}
	}

	numSubTotal = numSubTotal.toFixed(2);
	
	if(numMonitorTotal !=0){
		numGPUCost = numGPUCost.toFixed(2);
		document.getElementById('gpuPrice').innerHTML=numGPUCost;
	}

	numTotal = eval(numSubTotal) + eval(baseCost) + eval(numStandPrice) + eval(numMonitorTotal) + eval(numGPUCost) - eval(numBunDiscount);
	numTotal = numTotal.toFixed(2);
	
	finalTotal = eval(numTotal) * 1.2;
	finalTotal = finalTotal.toFixed(2);
	
	vatTotal = eval(finalTotal) - eval(numTotal);
	vatTotal = vatTotal.toFixed(2);

	document.getElementById('extrasPrice').innerHTML=numSubTotal;
	document.getElementById('subtotalPrice').innerHTML=numTotal;
	document.getElementById('vatPrice').innerHTML=vatTotal;
	document.getElementById('finalPrice').innerHTML=finalTotal;
		
	if(numMulti > 20){
		numMulti=20;
	}
		
	switch(numSpeed) {
				case 2:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars1.jpg" />';
					break;
				case 3:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars1-5.jpg" />';
					break;
				case 4:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars2.jpg" />';
					break;
				case 5:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars2-5.jpg" />';
					break;
				case 6:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars3.jpg" />';
					break;
				case 7:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars3-5.jpg" />';
					break;
				case 8:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars4.jpg" />';
					break;
				case 9:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars4-5.jpg" />';
					break;
				case 10:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars5.jpg" />';
					break;
				case 11:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars5-5.jpg" />';
					break;
				case 12:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars6.jpg" />';
					break;
				case 13:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars6-5.jpg" />';
					break;
				case 14:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars7.jpg" />';
					break;
				case 15:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars7-5.jpg" />';
					break;
				case 16:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars8.jpg" />';
					break;
				case 17:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars8-5.jpg" />';
					break;
				case 18:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars9.jpg" />';
					break;
				case 19:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars9-5.jpg" />';
					break;
				case 20:
					document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars10.jpg" />';
					break;
				default:
					break;
			}
		
		switch(numQuiet) {
				case 3:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars1-5.jpg" />';
					break;
				case 4:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars2.jpg" />';
					break;
				case 5:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars2-5.jpg" />';
					break;
				case 6:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars3.jpg" />';
					break;
				case 7:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars3-5.jpg" />';
					break;
				case 8:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars4.jpg" />';
					break;
				case 9:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars4-5.jpg" />';
					break;
				case 10:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars5.jpg" />';
					break;
				case 11:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars5-5.jpg" />';
					break;
				case 12:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars6.jpg" />';
					break;
				case 13:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars6-5.jpg" />';
					break;
				case 14:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars7.jpg" />';
					break;
				case 15:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars7-5.jpg" />';
					break;
				case 16:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars8.jpg" />';
					break;
				case 17:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars8-5.jpg" />';
					break;
				case 18:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars9.jpg" />';
					break;
				case 19:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars9-5.jpg" />';
					break;
				case 20:
					document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars10.jpg" />';
					break;
				default:
					break;
			}
		
		switch(numMulti) {
				case 2:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars1.jpg" />';
					break;
				case 3:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars1-5.jpg" />';
					break;
				case 4:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars2.jpg" />';
					break;
				case 5:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars2-5.jpg" />';
					break;
				case 6:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars3.jpg" />';
					break;
				case 7:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars3-5.jpg" />';
					break;
				case 8:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars4.jpg" />';
					break;
				case 9:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars4-5.jpg" />';
					break;
				case 10:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars5.jpg" />';
					break;
				case 11:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars5-5.jpg" />';
					break;
				case 12:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars6.jpg" />';
					break;
				case 13:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars6-5.jpg" />';
					break;
				case 14:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars7.jpg" />';
					break;
				case 15:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars7-5.jpg" />';
					break;
				case 16:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars8.jpg" />';
					break;
				case 17:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars8-5.jpg" />';
					break;
				case 18:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars9.jpg" />';
					break;
				case 19:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars9-5.jpg" />';
					break;
				case 20:
					document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars10.jpg" />';
					break;
				default:
					break;
			}
		switch(numMulThr) {
				case 2:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars1.jpg" />';
					break;
				case 3:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars1-5.jpg" />';
					break;
				case 4:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars2.jpg" />';
					break;
				case 5:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars2-5.jpg" />';
					break;
				case 6:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars3.jpg" />';
					break;
				case 7:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars3-5.jpg" />';
					break;
				case 8:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars4.jpg" />';
					break;
				case 9:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars4-5.jpg" />';
					break;
				case 10:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars5.jpg" />';
					break;
				case 11:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars5-5.jpg" />';
					break;
				case 12:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars6.jpg" />';
					break;
				case 13:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars6-5.jpg" />';
					break;
				case 14:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars7.jpg" />';
					break;
				case 15:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars7-5.jpg" />';
					break;
				case 16:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars8.jpg" />';
					break;
				case 17:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars8-5.jpg" />';
					break;
				case 18:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars9.jpg" />';
					break;
				case 19:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars9-5.jpg" />';
					break;
				case 20:
					document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars10.jpg" />';
					break;
			default:
					break;
			}
		switch(numGPU) {
				case 2:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars1.jpg" />';
					break;
				case 3:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars1-5.jpg" />';
					break;
				case 4:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars2.jpg" />';
					break;
				case 5:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars2-5.jpg" />';
					break;
				case 6:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars3.jpg" />';
					break;
				case 7:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars3-5.jpg" />';
					break;
				case 8:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars4.jpg" />';
					break;
				case 9:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars4-5.jpg" />';
					break;
				case 10:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars5.jpg" />';
					break;
				case 11:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars5-5.jpg" />';
					break;
				case 12:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars6.jpg" />';
					break;
				case 13:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars6-5.jpg" />';
					break;
				case 14:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars7.jpg" />';
					break;
				case 15:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars7-5.jpg" />';
					break;
				case 16:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars8.jpg" />';
					break;
				case 17:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars8-5.jpg" />';
					break;
				case 18:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars9.jpg" />';
					break;
				case 19:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars9-5.jpg" />';
					break;
				case 20:
					document.getElementById('stars-gpu').innerHTML='<img src="/images/generic/stars10.jpg" />';
					break;
				default:
					break;
			}
			
	}
	
function reCalcColourEXT() {
	
	var txtFlash = $("#optMotherboard").html();
	var items = ["* * *  - - -  Motherboard Upgraded  - - -  * * *",txtFlash, "* * *  - - -  Motherboard Upgraded  - - -  * * *",txtFlash],
        $text = $("#optMotherboard"),
        delay = 0.3; //seconds	
	
	var txtFlash2 = $("#optPSU").html();
	var items2 = ["* * *  - - -  Power Supply Upgraded  - - -  * * *",txtFlash2, "* * *  - - -  Power Supply Upgraded  - - -  * * *",txtFlash2],
        $text2 = $("#optPSU"),
        delay2 = 0.3; //seconds	
		
	var txtFlash3 = $("#optCPUCool").html();
	var items3 = ["* * *  - - -  CPU Cooling Upgraded  - - -  * * *",txtFlash3, "* * *  - - -  CPU Cooling Upgraded  - - -  * * *",txtFlash3],
        $text3 = $("#optCPUCool"),
        delay3 = 0.3; //seconds	
		
		$.each( items, function ( i, elm ){
            			$text.delay( delay*1E3).fadeOut();
           				$text.queue(function(){
                			$text.html( items[i] );
			                $text.dequeue();
            			});
            		$text.css('color','#50ae36').fadeIn();
            		$text.queue(function(){
                		if ( i == items.length -1 ) {
                    		//loop(delay);   
                		}
                	$text.dequeue();
            		});
        			});

		$.each( items2, function ( i, elm ){
            			$text2.delay( delay*1E3).fadeOut();
           				$text2.queue(function(){
                			$text2.html( items2[i] );
			                $text2.dequeue();
            			});
            		$text2.css('color','#50ae36').fadeIn();
            		$text2.queue(function(){
                		if ( i == items2.length -1 ) {
                    		//loop(delay);   
                		}
                	$text2.dequeue();
            		});
        			});

		$.each( items3, function ( i, elm ){
            			$text3.delay( delay3*1E3).fadeOut();
           				$text3.queue(function(){
                			$text3.html( items3[i] );
			                $text3.dequeue();
            			});
            		$text3.css('color','#50ae36').fadeIn();
            		$text3.queue(function(){
                		if ( i == items3.length -1 ) {
                    		//loop(delay3);   
                		}
                	$text3.dequeue();
            		});
        			});	
		
}

function reCalcColour() {
	
	var txtFlash = $("#optMotherboard").html();
	var items = ["* * *  - - -  Motherboard Upgraded  - - -  * * *",txtFlash, "* * *  - - -  Motherboard Upgraded  - - -  * * *",txtFlash],
        $text = $("#optMotherboard"),
        delay = 0.3; //seconds	
	
	var txtFlash2 = $("#optPSU").html();
	var items2 = ["* * *  - - -  Power Supply Upgraded  - - -  * * *",txtFlash2, "* * *  - - -  Power Supply Upgraded  - - -  * * *",txtFlash2],
        $text2 = $("#optPSU"),
        delay2 = 0.3; //seconds	
		
	var txtFlash3 = $("#optCaseCool").html();
	var items3 = ["* * *  - - -  Case Cooling Upgraded  - - -  * * *",txtFlash3, "* * *  - - -  Case Cooling Upgraded  - - -  * * *",txtFlash3],
        $text3 = $("#optCaseCool"),
        delay3 = 0.3; //seconds	

	
	var colddlop5 = document.getElementById('idOption5');
	if (colddlop5 != null)
	{
	for (var i = 0; i < colddlop5.options.length; i++) {
		if (colddlop5.options[i].selected == true) {
			//numSubTotal = eval(numSubTotal) + eval(ddlop5.options[i].title);
			switch(colddlop5.options[i].id) {
				// Flash Screens Supported Only
				case '70':	//'G210 x 1 - Pro
				case '156':	//'GT 630 x 1 - Ultra	
				case '160':	//'GT 630 x 1 - Extreme	
					$("#optScreensTitle").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
					break;
				// Flash Motherboard & Screens Supported	
				case '153':	//'G210 x 2 - Pro
				case '157': //'GT 630 x 1 GT 610 x 1 - Ultra
        			$.each( items, function ( i, elm ){
            			$text.delay( delay*1E3).fadeOut();
           				$text.queue(function(){
                			$text.html( items[i] );
			                $text.dequeue();
            			});
            		$text.css('color','#50ae36').fadeIn();
            		$text.queue(function(){
                		if ( i == items.length -1 ) {
                    		//loop(delay);   
                		}
                	$text.dequeue();
            		});
        			});
					$("#optScreensTitle").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
					break;
				// Flash Motherboard, Case Cooling & Screens Supported	
				case '154': //'G210 x 3 - Pro
				case '155': //'G210 x 4 - Pro
				case '264': //'710 x1 - Pro
				case '265': //'710 x 2 - Pro
				case '266': //'710 x 3 - Pro
				case '267': //'710 x 4 - Pro
        			$.each( items, function ( i, elm ){
            			$text.delay( delay*1E3).fadeOut();
           				$text.queue(function(){
                			$text.html( items[i] );
			                $text.dequeue();
            			});
            		$text.css('color','#50ae36').fadeIn();
            		$text.queue(function(){
                		if ( i == items.length -1 ) {
                    		//loop(delay);   
                		}
                	$text.dequeue();
            		});
        			});
					
					$.each( items3, function ( i, elm ){
            			$text3.delay( delay3*1E3).fadeOut();
           				$text3.queue(function(){
                			$text3.html( items3[i] );
			                $text3.dequeue();
            			});
            		$text3.css('color','#50ae36').fadeIn();
            		$text3.queue(function(){
                		if ( i == items3.length -1 ) {
                    		//loop(delay3);   
                		}
                	$text3.dequeue();
            		});
        			});			
					
					$("#optScreensTitle").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
					break;
				// Flash Motherboard, PSU & Screens Supported	
 				case '161': //'GT 630 x 1 GT 610 x 1 - Extreme
	       			$.each( items, function ( i, elm ){
            			$text.delay( delay*1E3).fadeOut();
           				$text.queue(function(){
                			$text.html( items[i] );
			                $text.dequeue();
            			});
            		$text.css('color','#50ae36').fadeIn();
            		$text.queue(function(){
                		if ( i == items.length -1 ) {
                    		//loop(delay);   
                		}
                	$text.dequeue();
            		});
        			});

	     			$.each( items2, function ( i, elm ){
            			$text2.delay( delay2*1E3).fadeOut();
           				$text2.queue(function(){
                			$text2.html( items2[i] );
			                $text2.dequeue();
            			});
            		$text2.css('color','#50ae36').fadeIn();
            		$text2.queue(function(){
                		if ( i == items2.length -1 ) {
                    		//loop(delay2);   
                		}
                	$text2.dequeue();
            		});
        			});
					
					$("#optScreensTitle").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
					break;
				// Flash Motherboard, Case Cooling, PSU & Screens Supported	
				case '158':	//'GT 630 x 1 GT 610 x 2 - Ultra
				case '159': //'GT 630 x 1 GT 610 x 3 - Ultra
				case '162': //'GT 630 x 1 GT 610 x 2 - Extreme
				case '163':	//'GT 630 x 1 GT 610 x 3 - Extreme
				case '164':	//'w600 x 2 - Ultra
				case '182':	//'w4100 x 2 - Ultra
				case '183':	//'w4100 x 3 - Ultra
				case '186':	//'w4100 x 3 - Extreme
				case '187':	//'w4100 x 4 - Extreme
				case '166':	//'w600 x 3 - Ultra
				case '168':	//'w600 x 4 - Ultra
				case '165':	//'w600 x 2 - Extreme
				case '167':	//'w600 x 3 - Extreme
				case '169':	//'w600 x 4 - Extreme
        			$.each( items, function ( i, elm ){
            			$text.delay( delay*1E3).fadeOut();
           				$text.queue(function(){
                			$text.html( items[i] );
			                $text.dequeue();
            			});
            		$text.css('color','#50ae36').fadeIn();
            		$text.queue(function(){
                		if ( i == items.length -1 ) {
                    		//loop(delay);   
                		}
                	$text.dequeue();
            		});
        			});
					
					$.each( items2, function ( i, elm ){
            			$text2.delay( delay2*1E3).fadeOut();
           				$text2.queue(function(){
                			$text2.html( items2[i] );
			                $text2.dequeue();
            			});
            		$text2.css('color','#50ae36').fadeIn();
            		$text2.queue(function(){
                		if ( i == items2.length -1 ) {
                    		//loop(delay2);   
                		}
                	$text2.dequeue();
            		});
        			});

					
					$.each( items3, function ( i, elm ){
            			$text3.delay( delay3*1E3).fadeOut();
           				$text3.queue(function(){
                			$text3.html( items3[i] );
			                $text3.dequeue();
            			});
            		$text3.css('color','#50ae36').fadeIn();
            		$text3.queue(function(){
                		if ( i == items3.length -1 ) {
                    		//loop(delay3);   
                		}
                	$text3.dequeue();
            		});
        			});			
					
					$("#optScreensTitle").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
					break;
				default:
					break;
			}
		}
	}
	}

	
}