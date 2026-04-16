function pageLD() {
 	var ddlop1 = document.getElementById('idOption4');
	if (ddlop1 != null)
	{
	for (var i = 0; i < ddlop1.options.length; i++) {
		if (ddlop1.options[i].selected == true) {
			numSubTotal = eval(ddlop1.options[i].title);
		}
	}
	}
	
	//document.getElementById('idOptionA').style.display = "none";
	//document.getElementById('idOptionAlbl').style.display = "none";
	
	numSubTotal = numSubTotal.toFixed(2);

	numPCTotal = eval(numSubTotal) + eval(baseCost);
	numPCTotal = numPCTotal.toFixed(2);
	
	numBunTotal = eval(numPCTotal) + eval(numStandPrice) + eval(numMonitorTotal) - eval(numBunDiscount);
	numBunTotal = numBunTotal.toFixed(2);
		
	vatTotal = eval(numPCTotal) * 1.2;
	vatTotal = vatTotal.toFixed(2);
	
	vatBunTotal = eval(numBunTotal) * 1.2;
	vatBunTotal = vatBunTotal.toFixed(2);

	document.getElementById('vatPrice').innerHTML=vatTotal;
	document.getElementById('pcPrice').innerHTML=numPCTotal;
	
	if (numMonitorTotal != 0) {
		document.getElementById('txtBunPrice').innerHTML='<p class="uppricefont"><strong class="">Synergy Stand:</strong> <strong class="color">&pound;' + numStandPrice + '</strong> <strong class="pri1">+ VAT</strong></p><p class="uppricefont"><strong class="">Monitors:</strong> <strong class="color">&pound;' + numMonitorTotal + '</strong> <strong class="pri1">+ VAT</strong></p><p class="uppricefont"><strong class="">Bundle Discount:</strong> <strong class="color">&pound;' + numBunDiscount + '</strong> <strong class="pri1">+ VAT</strong></p><p class="uppricefont"><strong class="">Total Bundle Price:</strong> <strong class="color">&pound;' + numBunTotal + '</strong> <strong class="pri1">+ VAT</strong></p>';
		document.getElementById('vatPrice').innerHTML=vatBunTotal;
		}
	
	document.getElementById('txtCPU').innerHTML='Intel i5 13600KF // 3.6 - 5.1GHz // 14C - 20T';
	document.getElementById('txtMB').innerHTML='Fast Z790 Chipset Motherboard';
	document.getElementById('txtRAM').innerHTML='32GB DDR5 5,200MHz RAM';
	document.getElementById('txtSSD').innerHTML='250GB Seagate NVMe M.2 SSD';
	document.getElementById('txtGPU').innerHTML='nVidia T600 Graphics Card';
	document.getElementById('txtCPUCool').innerHTML='BeQuiet Ultra Low Noise Air CPU Cooler';
	document.getElementById('optRes').innerHTML='<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 4 4K (3840 x 2160) 120Hz Screens</li><li>Up to 4 5K (5120 x 2880) 60Hz Screens</li><li>Up to 2 8K (7680 x 4320) 60Hz Screens</li>';
	document.getElementById('optPorts').innerHTML='<li>4x Mini-DisplayPort Outputs';
	document.getElementById('txtWAR').innerHTML='5 Year Warranty - 1 Yr Onsite / Replace / Collect';
	document.getElementById('txtWIN').innerHTML='Windows 11 Home Edition - 64-Bit';
	document.getElementById('txtKYB').innerHTML='Wireless Mouse / Keyboard Set';
	document.getElementById('stars-speed').innerHTML='<img src="/images/generic/stars7-5.jpg" />';
	document.getElementById('stars-multi').innerHTML='<img src="/images/generic/stars10.jpg" />';
	document.getElementById('stars-mulThr').innerHTML='<img src="/images/generic/stars6-5.jpg" />';
	document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars6-5.jpg" />';
	document.getElementById('stars-quiet').innerHTML='<img src="/images/generic/stars10.jpg" />';
	
	//Select right graphics setup
	$(document).ready(function() {
		//If we have more than 4 bundled monitor then change graphics option
		if (numMonitorAmount > 4) {
		//var ddlAutoSelectTxt = numMonitorAmount + ' Monitor Connections';
		var ddlAutoSelectTxt = 'Up to ' +numMonitorAmount;	 
		$("#idOption4").find("option:contains('" + ddlAutoSelectTxt +"')").each(function () {
   			$(this).attr("selected", "selected");return false;
		});
		reCalc();
		flashGPU();
		}
		});
		//If we have a numMonitorAmount then show the free bundle gifts in the first description
		if (numMonitorAmount > 1) {
			document.getElementById('txtWIFI').innerHTML='<p id="pWIFI">Wireless AC 867Mbps Network Card</p>';	
			document.getElementById('txtSPK').innerHTML='<p id="pSPK">Logitech Desktop Speakers</p>';
		}
	}

function reCalc() {
	
	var numSubTotal = 0;
	numSubTotal = numSubTotal.toFixed(2);
	
	var ddlop1 = document.getElementById('idOption4');
	if (ddlop1 != null)
	{
	for (var i = 0; i < ddlop1.options.length; i++) {
		if (ddlop1.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop1.options[i].title);
			switch(ddlop1.options[i].id){
				case '0':
				document.getElementById('txtGPU').innerHTML='nVidia T600 Graphics Card';
				document.getElementById('optRes').innerHTML='<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 4 4K (3840 x 2160) 120Hz Screens</li><li>Up to 4 5K (5120 x 2880) 60Hz Screens</li><li>Up to 2 8K (7680 x 4320) 60Hz Screens</li>';
				document.getElementById('optPorts').innerHTML='<li>4 x Mini-DisplayPort Outputs</li>';	
				numQuiet = 20;
				numGPU = 13;
				break;
				case '1':
				document.getElementById('txtGPU').innerHTML='nVidia RTX A2000 Graphics Card';
				document.getElementById('optRes').innerHTML='<li>Up to 4 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 4 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 4 4K (3840 x 2160) 120Hz Screens</li><li>Up to 4 5K (5120 x 2880) 60Hz Screens</li><li>Up to 2 8K (7680 x 4320) 60Hz Screens</li>';
				document.getElementById('optPorts').innerHTML='<li>4 x DisplayPort Outputs</li>';	
				numQuiet = 20;
				numGPU = 20;
				break;
				case '2':
				document.getElementById('txtGPU').innerHTML='nVidia T400 Graphics Card (x2)';
				document.getElementById('optRes').innerHTML='<li>Up to 6 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 6 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 5 4K (3840 x 2160) 60Hz Screens</li>';
				document.getElementById('optPorts').innerHTML='<li>6 x Mini-DisplayPort Outputs</li>';	
				numQuiet = 20;
				numGPU = 10;
				break;
				case '3':
				document.getElementById('txtGPU').innerHTML='nVidia T600 Graphics Card (x2)';
				document.getElementById('optRes').innerHTML='<li>Up to 8 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 8 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 7 4K (3840 x 2160) 60Hz Screens</li><li>Up to 4 5K (5120 x 2880) 60Hz Screens</li><li>Up to 2 8K (7680 x 4320) 60Hz Screens</li>';
				document.getElementById('optPorts').innerHTML='<li>8 x Mini-DisplayPort Outputs</li>';		
				numQuiet = 20;
				numGPU = 13;
				break;
				case '4':
				document.getElementById('txtGPU').innerHTML='nVidia RTX A2000 Graphics Card (x2)';
				document.getElementById('optRes').innerHTML='<li>Up to 8 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 8 QHD (2560 x 1440) 120Hz Screens</li><li>Up to 8 4K (3840 x 2160) 60Hz Screens</li><li>Up to 4 5K (5120 x 2880) 60Hz Screens</li><li>Up to 2 8K (7680 x 4320) 60Hz Screens</li>';
				document.getElementById('optPorts').innerHTML='<li>8 x DisplayPort Outputs</li>';	
				numQuiet = 20;
				numGPU = 20;
				break;
				case '5':
				document.getElementById('txtGPU').innerHTML='nVidia T1000 Graphics Card (x3)';
				document.getElementById('optRes').innerHTML='<li>Up to 12 FHD (1920 x 1080) 120Hz Screens</li><li>Up to 12 QHD (2560 x 1440) 60Hz Screens</li><li>Up to 7 4K (3840 x 2160) 60Hz Screens</li><li>Up to 4 5K (5120 x 2880) 60Hz Screens</li><li>Up to 2 8K (7680 x 4320) 60Hz Screens</li>';
				document.getElementById('optPorts').innerHTML='<li>12 x Mini-DisplayPort Outputs</li>';	
				numQuiet = 20;
				numGPU = 15;
				break;
			} 
		}
	}
	}
	
	var ddlop3 = document.getElementById('idOption12');
	if (ddlop3 != null)
	{
	for (var i = 0; i < ddlop3.options.length; i++) {
		if (ddlop3.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop3.options[i].title);
			if (ddlop3.options[i].id == 4){
				document.getElementById('txtMSO').innerHTML='';
			} else if (ddlop3.options[i].id == 6) {
				document.getElementById('txtMSO').innerHTML='<p id="pMSO">Microsoft Office 2021 Home &amp; Student</p>';
			} else {
				document.getElementById('txtMSO').innerHTML='<p id="pMSO">Microsoft Office 2021 Home &amp; Business</p>';
			}
		}
	}
	}

	var ddlop4 = document.getElementById('idOption13');
	if (ddlop4 != null)
	{
	for (var i = 0; i < ddlop4.options.length; i++) {
		if (ddlop4.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop4.options[i].title);
			if (ddlop4.options[i].id == 7){
				document.getElementById('txtWAR').innerHTML='5 Year Warranty - 1 Yr Onsite / Replace / Collect';
			} else if (ddlop4.options[i].id == 8) {
				document.getElementById('txtWAR').innerHTML='5 Year Warranty - 2 Yr Onsite / Replace / Collect';
			} else {
				document.getElementById('txtWAR').innerHTML='5 Year Warranty - 3 Yr Onsite / Replace / Collect';
			}
		}
	}
	}

	var ddlop5 = document.getElementById('idOption1');
	if (ddlop5 != null)
	{
	for (var i = 0; i < ddlop5.options.length; i++) {
		if (ddlop5.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop5.options[i].title);
			if (ddlop5.options[i].id == 1) {
				document.getElementById('txtCPU').innerHTML='Intel i5 13600KF // 3.6 - 5.1GHz // 14C - 20T';
				document.getElementById('txtMB').innerHTML='Fast Z790 Chipset Motherboard';
				document.getElementById('txtCPUCool').innerHTML='BeQuiet Ultra Low Noise Air CPU Cooler';
				numSpeed = 15;
				numMulti = 20;
				numMulThr = 13;
				numQuiet = 20;
			}   else if (ddlop5.options[i].id == 2) {
				document.getElementById('txtCPU').innerHTML='Intel i7 13700KF // 3.4 - 5.4GHz // 16C - 24T';
				document.getElementById('txtMB').innerHTML='Fast Z790 Chipset Motherboard';
				document.getElementById('txtCPUCool').innerHTML='BeQuiet Ultra Low Noise Air CPU Cooler';
				numSpeed = 17;
				numMulti = 20;
				numMulThr = 15;
				numQuiet = 16;
			}   else if (ddlop5.options[i].id == 3) {
				document.getElementById('txtCPU').innerHTML='Intel i9 13900KF // 3.0 - 5.8GHz // 24C - 32T';
				document.getElementById('txtMB').innerHTML='Fast Z790 Chipset Motherboard';
				document.getElementById('txtCPUCool').innerHTML='Corsair AIO Liquid CPU Cooler';
				numSpeed = 20;
				numMulti = 20;
				numMulThr = 17;
				numQuiet = 16;
			}   else if (ddlop5.options[i].id == 4) {
				document.getElementById('txtCPU').innerHTML='AMD Ryzen 9 7900X // 4.7 - 5.6GHz // 12C - 24T';
				document.getElementById('txtMB').innerHTML='Fast X670 Chipset Motherboard';
				document.getElementById('txtCPUCool').innerHTML='Corsair AIO Liquid CPU Cooler';
				numSpeed = 15;
				numMulti = 20;
				numMulThr = 16;
				numQuiet = 15;
			} else if (ddlop5.options[i].id == 5) {
				document.getElementById('txtCPU').innerHTML='AMD Ryzen 9 7950X // 4.5 - 5.7GHz // 16C - 32T';
				document.getElementById('txtMB').innerHTML='Fast X670 Chipset Motherboard';
				document.getElementById('txtCPUCool').innerHTML='Corsair AIO Liquid CPU Cooler';
				numSpeed = 17;
				numMulti = 20;
				numMulThr = 20;
				numQuiet = 15;
			} 
		}
	}
	}
	
	var ddlop6 = document.getElementById('idOption2');
	if (ddlop6 != null)
	{
	for (var i = 0; i < ddlop6.options.length; i++) {
		if (ddlop6.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop6.options[i].title);
			if (ddlop6.options[i].id == 0){
				document.getElementById('txtRAM').innerHTML='32GB DDR5 5,200MHz RAM';
			} else if (ddlop6.options[i].id == 1) {
				document.getElementById('txtRAM').innerHTML='64GB DDR5 5,200MHz RAM';
			} else if (ddlop6.options[i].id == 2) {
				document.getElementById('txtRAM').innerHTML='128GB DDR5 3,600MHz RAM';
			}
			}
	}
	}

	var ddlop7 = document.getElementById('idOption3');
	if (ddlop7 != null)
	{
	for (var i = 0; i < ddlop7.options.length; i++) {
		if (ddlop7.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop7.options[i].title);
			if (ddlop7.options[i].id == 0){
				document.getElementById('txtSSD').innerHTML='250GB Seagate NVMe M.2 SSD';
			} else if (ddlop7.options[i].id == 1) {
				document.getElementById('txtSSD').innerHTML='500GB Seagate NVMe M.2 SSD';
			} else if (ddlop7.options[i].id == 2) {
				document.getElementById('txtSSD').innerHTML='1TB Adata NVMe M.2 SSD';
			} else if (ddlop7.options[i].id == 3) {
				document.getElementById('txtSSD').innerHTML='2TB Adata NVMe M.2 SSD';
			}
		}
	}
	}


	var ddlop8 = document.getElementById('idOption11');
	if (ddlop8 != null)
	{
	for (var i = 0; i < ddlop8.options.length; i++) {
		if (ddlop8.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop8.options[i].title);
			if (ddlop8.options[i].id == 0){
				document.getElementById('txtBBD').innerHTML='';
			} else if (ddlop8.options[i].id == 1) {
				document.getElementById('txtBBD').innerHTML='<p id="pBBD">Bootable Backup Drive</p>';
				numQuiet = numQuiet - 1;
			} 
		}
	}
	}
	
	var ddlop9 = document.getElementById('idOption10');
	if (ddlop9 != null)
	{
	for (var i = 0; i < ddlop9.options.length; i++) {
		if (ddlop9.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop9.options[i].title);
			if (ddlop9.options[i].id == 1){
				document.getElementById('txtWIN').innerHTML='Windows 10 Home Edition - 64-Bit';
			} else if (ddlop9.options[i].id == 2) {
				document.getElementById('txtWIN').innerHTML='Windows 10 Professional Edition - 64-Bit';
			} else if (ddlop9.options[i].id == 3) {
				document.getElementById('txtWIN').innerHTML='Windows 11 Home Edition - 64-Bit';
			} else if (ddlop9.options[i].id == 4) {
				document.getElementById('txtWIN').innerHTML='Windows 11 Professional Edition - 64-Bit';
			}
		}
	}
	}
	
	var ddlop11 = document.getElementById('idOption9');
	if (ddlop11 != null)
	{
	for (var i = 0; i < ddlop11.options.length; i++) {
		if (ddlop11.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop11.options[i].title);
			if (ddlop11.options[i].id == 0){
				document.getElementById('txtKYB').innerHTML='<del>Wireless Mouse / Keyboard Set</del>';
			} else if (ddlop11.options[i].id == 1) {
				document.getElementById('txtKYB').innerHTML='Wired Mouse / Keyboard Set';
			} else if (ddlop11.options[i].id == 2) {
				document.getElementById('txtKYB').innerHTML='Wireless Mouse / Keyboard Set';
			}  
		}
	}
	}
	
	var ddlop12 = document.getElementById('idOption8');
	if (ddlop12 != null)
	{
	for (var i = 0; i < ddlop12.options.length; i++) {
		if (ddlop12.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop12.options[i].title);
			if (ddlop12.options[i].id == 0){
				document.getElementById('txtWIFI').innerHTML='';
			} else if (ddlop12.options[i].id == 1) {
				document.getElementById('txtWIFI').innerHTML='<p id="pWIFI">Wireless AC 867Mbps Network Card</p>';
			}  
		}
	}
	}
	
	var ddlop14 = document.getElementById('idOption6');
	if (ddlop14 != null)
	{
	for (var i = 0; i < ddlop14.options.length; i++) {
		if (ddlop14.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop14.options[i].title);
			if (ddlop14.options[i].id == 0){
				document.getElementById('txtSPK').innerHTML='';
			} else if (ddlop14.options[i].id == 1) {
				document.getElementById('txtSPK').innerHTML='<p id="pSPK">Logitech Desktop Speakers</p>';
			} 
		}
	}
	}
	
	var ddlop13 = document.getElementById('idOption14');
	if (ddlop13 != null)
	{
	for (var i = 0; i < ddlop13.options.length; i++) {
		if (ddlop13.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop13.options[i].title);
			if (ddlop13.options[i].id == 0){
				document.getElementById('txtHDD2').innerHTML='';
			} else if (ddlop13.options[i].id == 1) {
				document.getElementById('txtHDD2').innerHTML='<p id="pHDD2">1TB Traditional Hard Drive</p>';
				numQuiet = numQuiet - 2;
			} else if (ddlop13.options[i].id == 2) {
				document.getElementById('txtHDD2').innerHTML='<p id="pHDD2">2TB Traditional Hard Drive</p>';
				numQuiet = numQuiet - 2;
			} else if (ddlop13.options[i].id == 3) {
				document.getElementById('txtHDD2').innerHTML='<p id="pHDD2">3TB Traditional Hard Drive</p>';
				numQuiet = numQuiet - 2;
			} else if (ddlop13.options[i].id == 4) {
				document.getElementById('txtHDD2').innerHTML='<p id="pHDD2">4TB Traditional Hard Drive</p>';
				numQuiet = numQuiet - 2;
			} else if (ddlop13.options[i].id == 5) {
				document.getElementById('txtHDD2').innerHTML='<p id="pHDD2">6TB Traditional Hard Drive</p>';
				numQuiet = numQuiet - 2;
			} else if (ddlop13.options[i].id == 6) {
				document.getElementById('txtHDD2').innerHTML='<p id="pHDD2">500GB WD Blue SSD (500MBs/500MBs)</p>';
			} else if (ddlop13.options[i].id == 7) {
				document.getElementById('txtHDD2').innerHTML='<p id="pHDD2">1TB WD Blue SSD (500MBs/500MBs)</p>';
			} else if (ddlop13.options[i].id == 8) {
				document.getElementById('txtHDD2').innerHTML='<p id="pHDD2">2TB WD Blue SSD (500MBs/500MBs)</p>';
			} else if (ddlop13.options[i].id == 9) {
				document.getElementById('txtHDD2').innerHTML='<p id="pHDD2">4TB WD Blue SSD (500MBs/500MBs)</p>';
			} 
		}
	}
	}
	
	var ddlop16 = document.getElementById('idOption15');
	if (ddlop16 != null)
	{
	for (var i = 0; i < ddlop16.options.length; i++) {
		if (ddlop16.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop16.options[i].title);
			if (ddlop16.options[i].id == 0){
				document.getElementById('txtDVD').innerHTML='';
			} else if (ddlop16.options[i].id == 1) {
				document.getElementById('txtDVD').innerHTML='<p id="pDVD">DVD ReWriter Drive</p>';
			} 
		}
	}
	}
	
	var ddlop17 = document.getElementById('idOption16');
	if (ddlop17 != null)
	{
	for (var i = 0; i < ddlop17.options.length; i++) {
		if (ddlop17.options[i].selected == true) {
			numSubTotal = eval(numSubTotal) + eval(ddlop17.options[i].title);
			if (ddlop17.options[i].id == 0){
				document.getElementById('txtBT').innerHTML='';
			} else if (ddlop17.options[i].id == 1) {
				document.getElementById('txtBT').innerHTML='<p id="pBT">USB Bluetooth Adapter</p>';
			} 
		}
	}
	}


	if(numMulti>20){
		numMulti=20;
	}


	switch(numSpeed) {
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
	switch(numMulti) {
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
				case 3:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars1-5.jpg" />';
					break;
				case 4:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars2.jpg" />';
					break;
				case 5:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars2-5.jpg" />';
					break;
				case 6:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars3.jpg" />';
					break;
				case 7:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars3-5.jpg" />';
					break;
				case 8:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars4.jpg" />';
					break;
				case 9:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars4-5.jpg" />';
					break;
				case 10:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars5.jpg" />';
					break;
				case 11:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars5-5.jpg" />';
					break;
				case 12:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars6.jpg" />';
					break;
				case 13:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars6-5.jpg" />';
					break;
				case 14:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars7.jpg" />';
					break;
				case 15:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars7-5.jpg" />';
					break;
				case 16:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars8.jpg" />';
					break;
				case 17:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars8-5.jpg" />';
					break;
				case 18:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars9.jpg" />';
					break;
				case 19:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars9-5.jpg" />';
					break;
				case 20:
					document.getElementById('stars-GPU').innerHTML='<img src="/images/generic/stars10.jpg" />';
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


	numSubTotal = numSubTotal.toFixed(2);

	numPCTotal = eval(numSubTotal) + eval(baseCost);
	numPCTotal = numPCTotal.toFixed(2);
	
	numBunTotal = eval(numPCTotal) + eval(numStandPrice) + eval(numMonitorTotal) - eval(numBunDiscount);
	numBunTotal = numBunTotal.toFixed(2);
		
	vatTotal = eval(numPCTotal) * 1.2;
	vatTotal = vatTotal.toFixed(2);
	
	vatBunTotal = eval(numBunTotal) * 1.2;
	vatBunTotal = vatBunTotal.toFixed(2);

	document.getElementById('vatPrice').innerHTML=vatTotal;
	document.getElementById('pcPrice').innerHTML=numPCTotal;
	if (numMonitorTotal != 0) {
		document.getElementById('txtBunPrice').innerHTML='<p class="uppricefont"><strong class="">Synergy Stand:</strong> <strong class="color">&pound;' + numStandPrice + '</strong> <strong class="pri1">+ VAT</strong></p><p class="uppricefont"><strong class="">Monitors (x' + numMonitorAmount + '):</strong> <strong class="color">&pound;' + numMonitorTotal + '</strong> <strong class="pri1">+ VAT</strong></p><p class="uppricefont"><strong class="">Bundle Discount:</strong> <strong class="color">&pound;' + numBunDiscount + '</strong> <strong class="pri1">+ VAT</strong></p><p class="uppricefont"><strong class="">Total Bundle Price:</strong> <strong class="color">&pound;' + numBunTotal + '</strong> <strong class="pri1">+ VAT</strong></p>';
		document.getElementById('vatPrice').innerHTML=vatBunTotal;
		}
	}

function flashGPU() {
	$("#txtGPU").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
	$("#optScreensTitle").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashCPU() {
	$("#txtCPU").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
	$("#txtMB").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
	$("#txtCPUCool").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashRAM() {
	$("#txtRAM").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashSSD() {
	$("#txtSSD").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashBBD() {
	$("#pBBD").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashDVD() {
	$("#pDVD").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashBT() {
	$("#pBT").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashHDD2() {
	$("#pHDD2").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashSPK() {
	$("#pSPK").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashKYB() {
	$("#pKYB").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashWIFI() {
	$("#pWIFI").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashMSO() {
	$("#pMSO").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashWAR() {
	$("#txtWAR").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

function flashWIN() {
	$("#txtWIN").fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600).fadeOut(600).fadeIn(600);
}

