<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/utilities.asp"-->
<% 
Dim pageTitle, section
pageTitle=dictLanguageCP.Item(Session("language")&"_cpInstPrd_1")
section="products"

'// Determine product type: std, bto, item, app
'// std = "Standard" product
'// bto = "Configurable" product
'// item = "Configurable-Only Item"
'// app = "Apparel" product
pcv_ProductType=lcase(trim(request.Querystring("prdType")))
	'// If not an accepted Product Type, go get it
	if pcv_ProductType="" or (pcv_ProductType<>"std" and pcv_ProductType<>"bto" and pcv_ProductType<>"item") then
		pcv_ProductType="std"
	end if
%>
<!--#include file="AdminHeader.asp"-->

<!-- #include file="../htmleditor/editor.asp" -->

<!--#include file="../includes/javascripts/pcWindowsViewPrd.asp"-->

<% pcv_IdRootCategory=request("idRootCategory")
if NOT isNumeric(pcv_IdRootCategory) or pcv_IdRootCategory="" then
	pcv_IdRootCategory=1
end if

dim f

'if form submitted
if request("catCnt")<>"" then %>
	<script type=text/javascript>
		<%' GGG add-on start%>
		function check_date(field){
			var checkstr = "0123456789";
			var DateField = field;
			var Datevalue = "";
			var DateTemp = "";
			var seperator = "/";
			var day;
			var month;
			var year;
			var leap = 0;
			var err = 0;
			var i;
			err = 0;
			DateValue = DateField.value;
			/* Delete all chars except 0..9 */
			for (i = 0; i < DateValue.length; i++) {
			if (checkstr.indexOf(DateValue.substr(i,1)) >= 0) {
				 DateTemp = DateTemp + DateValue.substr(i,1);
			}
			else
			{
			if (DateTemp.length == 1)
			{
					DateTemp = "0" + DateTemp
			}
			else
			{
				if (DateTemp.length == 3)
				{
				DateTemp = DateTemp.substr(0,2) + '0' + DateTemp.substr(2,1);
				}
			}
			}
			}
			DateValue = DateTemp;
			/* Always change date to 8 digits - string*/
			/* if year is entered as 2-digit / always assume 20xx */
			if (DateValue.length == 6) {
			DateValue = DateValue.substr(0,4) + '20' + DateValue.substr(4,2); }
			if (DateValue.length != 8) {
			return(false);}
			/* year is wrong if year = 0000 */
			year = DateValue.substr(4,4);
			if (year == 0) {
			err = 20;
			}
			/* Validation of month*/
			<%if scDateFrmt="DD/MM/YY" then%>
			month = DateValue.substr(2,2);
			<%else%>
			month = DateValue.substr(0,2);
			<%end if%>
			if ((month < 1) || (month > 12)) {
				err = 21;
			}
			/* Validation of day*/
			<%if scDateFrmt="DD/MM/YY" then%>
			day = DateValue.substr(0,2);
			<%else%>
			day = DateValue.substr(2,2);
			<%end if%>
			if (day < 1) {
			 err = 22;
			}
			/* Validation leap-year / february / day */
			if ((year % 4 == 0) || (year % 100 == 0) || (year % 400 == 0)) {
				leap = 1;
			}
			if ((month == 2) && (leap == 1) && (day > 29)) {
				err = 23;
			}
			if ((month == 2) && (leap != 1) && (day > 28)) {
				err = 24;
			}
			/* Validation of other months */
			if ((day > 31) && ((month == "01") || (month == "03") || (month == "05") || (month == "07") || (month == "08") || (month == "10") || (month == "12"))) {
				err = 25;
			}
			if ((day > 30) && ((month == "04") || (month == "06") || (month == "09") || (month == "11"))) {
				err = 26;
			}
			/* if 00 ist entered, no error, deleting the entry */
			if ((day == 0) && (month == 0) && (year == 00)) {
				err = 0; day = ""; month = ""; year = ""; seperator = "";
			}
			/* if no error, write the completed date to Input-Field (e.g. 13.12.2001) */
			if (err == 0) {
			<%if scDateFrmt="DD/MM/YY" then%>
			DateField.value = day + seperator + month + seperator + year;
			<%else%>
			DateField.value = month + seperator + day + seperator + year;   
			<%end if%>
			return(true);
			}
			/* Error-message if err != 0 */
			else {
			return(false);   
			}
			}
			<%' GGG add-on end%>
			
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
			
			function CountStr(tmpStr,subStr)
			{
			   var substrings = tmpStr.split(subStr);
			   count=substrings.length - 1;
			   return count;
			}
		
			function Form1_Validator(theForm)
			{
				// InnovaStudio HTML Editor Workaround for this keyword
				theForm = document.hForm;
				
				SavePPToFields();
	
				if (theForm.sku.value == "")
				{
					 alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_1")%>");
					return (false);
				}
				if (theForm.description.value == "")
				{
					 alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_2")%>");
					return (false);
				}
				if (theForm.details.value == "")
				{
					 alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_3")%>");
					return (false);
				}
		
				if (theForm.idCategory1.value == "")
				{
					 alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_4")%>");
					return (false);
				}
			
				<%if pcv_ProductType<>"item" then%>
			
				if (theForm.downloadable1.value == "1")
				{
				
					if (theForm.producturl.value == "")
					{
						alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_5")%>");
						return (false);
					}
		
					if (theForm.urlexpire1.value == "1")
					{
				
						if (theForm.expiredays.value == "")
						{
							alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_6")%>");
							return (false);
						}
			
						if (allDigit(theForm.expiredays.value) == false)
						{
							alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_6")%>");
							return (false);
						}
			
						if (theForm.expiredays.value == "0")
						{
							alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_7")%>");
							return (false);
						}
					}
			
					if (theForm.license1.value == "1")
					{
				
						if ((theForm.locallg.value == "") && ((theForm.remotelg.value == "") || (theForm.remotelg.value == "http://")) )
						{
							alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_8")%>");
							return (false);
						}
			
						if ((theForm.locallg.value != "") && (theForm.remotelg.value != "") && (theForm.remotelg.value != "http://") )
						{
							 alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_9")%>");
							return (false);
						}
			
						if ((theForm.licenselabel1.value == "") && (theForm.licenselabel2.value == "") && (theForm.licenselabel3.value == "") && (theForm.licenselabel4.value == "") && (theForm.licenselabel5.value == ""))
						{
							alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_10")%>");
							return (false);
						}
					}
		
				}
				<%end if%>
		
				<%' GGG add-on start
				if pcv_ProductType="std" then %>
		
				if (theForm.GC[0].checked == true)
				{
					if (theForm.GCExp[1].checked == true)
					{
						if (theForm.GCExpDate.value == "")
						{
							alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_11")%>");
							return (false);
						}
						if (check_date(theForm.GCExpDate) == false)
						{
							alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_11")%>");
							return (false);
						}
					}
					if (theForm.GCExp[2].checked == true)
					{
						if (theForm.GCExpDay.value == "")
						{
							alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_6")%>");
							return (false);
						}
			
						if (allDigit(theForm.GCExpDay.value) == false)
						{
							alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_6")%>");
							return (false);
						}
			
						if (theForm.GCExpDay.value == "0")
						{
							alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_7")%>");
							return (false);
						}
		
					}
				
					if (theForm.GCGen[1].checked == true)
					{
						if (theForm.GCGenFile.value == "")
						{
							alert("<%=dictLanguageCP.Item(Session("language")&"_cpAlert_12")%>");
							return (false);
						}
					}
		
				}
			
				<%end if
				' GGG add-on end%>
				
				<% 'Check duplicated Custom Input Fields in Product Page Layout%>
				if (theForm.displayLayout.value=="t")
				{
				var CIF=0;
				if (theForm.ppTop.value != "") CIF=CIF+CountStr(theForm.ppTop.value,"PrdInput");
				if (theForm.ppTopLeft.value != "") CIF=CIF+CountStr(theForm.ppTopLeft.value,"PrdInput");
				if (theForm.ppTopRight.value != "") CIF=CIF+CountStr(theForm.ppTopRight.value,"PrdInput");
				if (theForm.ppMiddle.value != "") CIF=CIF+CountStr(theForm.ppMiddle.value,"PrdInput");
				if (theForm.ppTabs.value != "") CIF=CIF+CountStr(theForm.ppTabs.value,"PrdInput");
				if (theForm.ppBottom.value != "") CIF=CIF+CountStr(theForm.ppBottom.value,"PrdInput");
				
				if (CIF>1)
				{
					alert("You can not add 'Custom Input Fields' into multiple areas of Product Page Layout.");
					return (false);
				}
				}
	
				try
				{
					document.hForm.pcIDDropShipper.disabled=false;
					document.hForm.pcIDSupplier.disabled=false;
				}
				catch(err)
				{
					//Do nothing
				}
				return (true);
			}

		function CheckWindow() {
		options = "toolbar=0,status=0,menubar=0,scrollbars=0,resizable=0,width=600,height=400";
		myloc='testurl.asp?file1=' + document.hForm.producturl.value + '&file2=' + document.hForm.locallg.value + '&file3=' + document.hForm.remotelg.value;
		newcheckwindow=window.open(myloc,"mywindow", options);
		}
		
		function newWindow(file,window) {
			msgWindow=open(file,window,'resizable=no,width=400,height=500');
			if (msgWindow.opener == null) msgWindow.opener = self;
		}

		// Set mouse cursor focus on page load
		function setCursorFocus(){
		document.hForm.sku.focus();
		}
		onload = function() {setCursorFocus()}
	</script>
	
	
	<%
	'// START - Interface for add/selecting the category
	%>


<table class="pcCPcontent">
        <tr> 
            <td colspan="2">
				<div class="cpOtherLinks">
                    <strong>Product type</strong>: you are adding a <u>											
                    <% if pcv_ProductType="std" then %>
                        Standard product
                    <% elseif pcv_ProductType="app" then %>
                        Apparel product
                    <% elseif pcv_ProductType="bto" then %>
                        Configurable product
                    <% else %>
                        Configurable-Only item
                    <% end if %>
                    </u>
                    &nbsp;|&nbsp;												
                    <a href="LocateProducts.asp"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_2")%></a>
                    &nbsp;|&nbsp;
                    <img src="images/pc_required.gif" alt="required field" width="9" height="9" hspace="5"><%=dictLanguageCP.Item(Session("language")&"_cpAlert_0")%>
				</div>
                <h2>Choose one or more categories</h2>
            </td>
        </tr>
        <tr>
            <td colspan="2">
				<script type=text/javascript>
					accor1=0;
				</script>
				<a id="addnewcat" href="#" onClick="javascript: if (accor1==0) {accor1=1; document.getElementById('accor1').style.display='';} else {accor1=0; document.getElementById('accor1').style.display='none';}"><strong>Add a new category &gt;&gt;</strong></a>
				<table class="pcCPcontent">
		        <tr>
				<td>
				<form action="addProduct.asp" id="FormAddCat" class="pcForms">
				<table id="accor1" width="100%" style="padding:5px; border:solid 1px #CCC; display:none;">
				<tr>
					<td width="100">
						<%=dictLanguageCP.Item(Session("language")&"_cpCommon_157")%>:
					</td>
					<td>
						<input id="CategoryName" name="CategoryName" type="text" value="" />
					</td>
				</tr>
				<tr>
					<td>
						<%=dictLanguageCP.Item(Session("language")&"_cpCommon_158")%>:
					</td>
					<td>
						<span id="ParentCatList"></span>
					</td>
				</tr>
				<tr>
					<td colspan="2">
						<input type="button" class="btn btn-default"  value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_106")%>" onclick="javascript:AddNewCat();" class="btn btn-primary" />
					</td>
				</tr>
				<tr>
					<td colspan="2">
						<div id="test"></div>
					</td>
				</tr>
				</table>
				</form>
				<br>
				</td>
				</tr>
				</table>
				
				<script type=text/javascript>
				function AddNewCat()
				{
					$pc("#test").html("<img src=\"images/pc_AjaxLoader.gif\" border=0 align=\"texttop\"> Adding new category...");
					$pc.ajax({
						type: "POST",
						data: encodeURI("CategoryName="+$pc('#CategoryName').val()+"&ParentCatID="+$pc('#ParentCatID').val()),
						url: "pcAddCatAction.asp",
						timeout: 45000,
					}).done(function ( data ) {
					if(data.indexOf("pcCPmessageSuccess")>=0) {
						$pc("#test").html(data);
						LoadCategoryList();
					}
					else
					{
						$pc("#test").html(data);
					}
					}).fail(function() { alert("error"); });
				}
				
				</script>
			</td>
		</tr>
	</table>
	
	<%
	'// END - Interface for add/selecting the category
	%>
	<img src="images/pc_admin.gif" width="85" height="19" alt="Separator between two options" style="padding-left: 10px; margin-top: -10px;">
	<form method="post" name="hForm" action="addProductB.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
		<table class="pcCPcontent">
			<tr>
				<td colspan="2"><strong>Use one or more existing categories</strong>:</td>
			</tr>
			<tr> 
				<td colspan="2" align="left" style="padding: 15px;">
					<span id="CategoryList"></span>
				</td>
			</tr>
	  </table>
	  <script type=text/javascript>
	  	function replaceAll(find, replace, str) {
 			return str.replace(new RegExp(find, 'g'), replace);
		}
	  
	  	function LoadCategoryList()
		{
			$pc("#CategoryList").html("<img src=\"images/pc_AjaxLoader.gif\" border=0 align=\"texttop\"> Loading categories...");
			$pc.ajax({
				type: "GET",
				data: "idRootCategory=<%=pcv_IdRootCategory%>",
				url: "pcRequestCategories.asp",
				timeout: 45000,
	 		}).done(function ( data ) {
			if(data=="NONE") {
				$pc("#CategoryList").html("<b>No categories</b>");
				$pc("#ParentCatList").html("<b>No categories</b>");
			}
			else
			{
				$pc("#CategoryList").html("<select name=\"idCategory1\" size=\"10\" multiple style=\"width: 600px;\">"+data+"</select>");
				$pc("#ParentCatList").html("<select id=\"ParentCatID\" name=\"ParentCatID\">"+replaceAll("selected","",data)+"</select>");
			}
			}).fail(function() { alert("error"); });
		}
		$pc(document).ready(function()
		{
			LoadCategoryList();
		});
	</script>
		
		<%
		'// TABBED PANELS - START NAVIGATION
		%>
	  <div id="TabbedPanels1" class="tabbable-left">
		
		<div class="col-xs-3">
			<ul class="nav nav-tabs tabs-left">
				<li class="active">
					<a href="#tab-1" data-toggle="tab">
						Product Details
						<div class="pcCPextraInfo"><span class="pcSmallText">SKU, Descriptions, and Meta Tags</span></div>
					</a>
				</li>
				<li>
					<a href="#tab-2" data-toggle="tab">
						Prices
						<div class="pcCPextraInfo"><span class="pcSmallText">Online, Retail, &amp; Wholesale Prices</span></div>
					</a>
				</li>	
				<li>
					<a href="#tab-3" data-toggle="tab">
						Images
						<div class="pcCPextraInfo"><span class="pcSmallText">Primary &amp; Additional Images</span></div>
					</a>
				</li>
				<li>
					<a href="#tab-4" data-toggle="tab">
						Product Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Brand, Tax, &amp; Display Options</span></div>
					</a>
				</li>
				<li>
					<a href="#tab-5" data-toggle="tab">
						Shipping Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Inventory, Weight, &amp; Shipping</span></div>
					</a>
				</li>
				<% if pcv_ProductType<>"item" then %>
				<li>
					<a href="#tab-6" data-toggle="tab">
						Product Page Layout
						<div class="pcCPextraInfo"><span class="pcSmallText">Preset, Custom, &amp; Tabbed Layouts</span></div>
					</a>
				</li>
				<li>
					<a href="#tab-7" data-toggle="tab">
						Downloadable Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Make it a Digital Product</span></div>
					</a>
				</li>
				<% end if %>
				
				<% If statusAPP="1" OR scAPP=1 Then %>
				<li>
					<a href="#tabs-12" data-toggle="tab">
						Apparel Product Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Make it an Apparel Product</span></div>
					</a>
				</li>
				<% End If %>

				<% if pcv_ProductType="std" then %>
				<li>
					<a href="#tab-8" data-toggle="tab">
						Gift Certificate Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Make it a Gift Certificate</span></div>
					</a>
				</li>
				<% end if %>
				<% if pcv_ProductType<>"item" then %>
				<li>
					<a href="#tab-9" data-toggle="tab">
						Custom Fields
						<div class="pcCPextraInfo"><span class="pcSmallText">Manage Custom Search Fields</span></div>
					</a>
				</li>
				<li>
					<a href="#tab-10" data-toggle="tab">
						Google Shopping Settings
						<div class="pcCPextraInfo"><span class="pcSmallText">Setup with Google Shopping</span></div>
					</a>
				</li>
				<% end if %>
				<li>
					<div style="height:40px; margin-top:10px; text-align: center">
					<input type="hidden" name="idsupplier" value="10">
					<input type="hidden" name="prdType" value="<%=pcv_ProductType%>">
					<input type="submit" name="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_100")%>" value="Add Product" class="btn btn-primary" onClick="SavePPToFields();">
					</div>
				</li>
			</ul>
		</div>
		<%
		'// TABBED PANELS - END NAVIGATION
		
		'// TABBED PANELS - START PANELS
		%>
        <div class="col-xs-9">
            <div class="tab-content">
		
			<%
			'// =========================================
			'// FIRST PANEL - START - Name, SKU, descriptions
			'// =========================================
			%>
				<div id="tab-1" class="tab-pane active">
				
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_4")%></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td nowrap><%=dictLanguageCP.Item(Session("language")&"_cpCommon_26")%>:<img src="images/pc_required.gif" alt="required field" width="9" height="9"></td>
							<td>
								<input type="text" name="sku" size="30" tabindex="101"> 
							</td>
						</tr>
						<tr>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_28")%>:<img src="images/pc_required.gif" alt="required field" width="9" height="9"></td>
							<td>  
								<input type="text" name="description" size="40" tabindex="102">
							</td>
						</tr>
						<tr>
							<td valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_27")%>:<img src="images/pc_required.gif" alt="required field" width="9" height="9"></td>
							<td>  
								<textarea class="htmleditor" name="details" id="details" rows="6" cols="56" tabindex="103"><%=pDetails%></textarea>
							</td>
						</tr>
						<% if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item %>
						<tr>
							<td valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_29")%>:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=401"></td>
							<td>
								<textarea name="sdesc" rows="6" cols="60" tabindex="104"></textarea>
							</td>
						</tr>
						<% end if %>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
            <tr>
							<th colspan="2">SEO/Meta Tags&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=204"></a></th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td align="right" valign="top">Title: </td>
							<td><textarea name="PrdMetaTitle" cols="50" rows="3" tabindex="1001"><%=pStrPrdMetaTitle%></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Description: </td>
							<td><textarea name="PrdMetaDesc" cols="50" rows="6" tabindex="1002"><%=pStrPrdMetaDesc%></textarea>
						</tr>
						<tr>
							<td align="right" valign="top">Keywords: </td>
							<td><textarea name="PrdMetaKeywords" cols="50" rows="4" tabindex="1003"><%=pStrPrdMetaKeywords%></textarea>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
					</table>
				
				</div>
			<%
			'// =========================================
			'// FIRST PANEL - END
			'// =========================================
			
			'// =========================================
			'// SECOND PANEL - START - Prices
			'// =========================================
			%>
				<div id="tab-2" class="tab-pane">
				
					<table class="pcCPcontent">

						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_5")%></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<% if pcv_ProductType="std" then %>
							<td width="30%">Online Price:</td>
							<% else %>
							<td>Base Price:&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=500"></a></td>
							<% end if %>
							<td width="70%"><%=scCurSign%>&nbsp;<input type="text" name="price" value="0" size="10" tabindex="201"></td>
						</tr>
						<% if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item %>
						<tr> 
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_31")%>:</td>
							<td><%=scCurSign%> <input type="text" name="listPrice" value="0" size="10" tabindex="202"></td>
						</tr>
						<tr> 
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_34")%>:</td>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%> <input type="checkbox" name="listhidden" value="-1" tabindex="203" class="clearBorder">
							</td>
						</tr>
						<% end if ' Hide if it's a Configurable-Only Item

						'START CT ADD
						'if there are customer type categories - List them here 
						dim intATBExists
						intATBExists=0
						query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories;"
						SET rs=Server.CreateObject("ADODB.RecordSet")
						SET rs=conntemp.execute(query)
						if NOT rs.eof then 
							do until rs.eof 
								intIdcustomerCategory=rs("idcustomerCategory")
								strpcCC_Name=rs("pcCC_Name")
								strpcCC_CategoryType=rs("pcCC_CategoryType")
								%>
								<tr>
									<td><%=strpcCC_Name%></td>
									<td><%=scCurSign%><input type="text" name="pcCC_<%=intIdcustomerCategory%>" value="0" size="10">
									<% if strpcCC_CategoryType="ATB" then %>
									<br /><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_6")%>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=308"></a>
			
									<%	end if %>
									</td>
								</tr>
							<% rs.moveNext
							loop
						end if
						SET rs=nothing
						'END CT ADD %>
						<tr> 
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_32")%>:</td>
							<td><%=scCurSign%> <input type="text" name="bToBprice" value="0" size="10" tabindex="204"></td>
						</tr>
						<%'Start SDBA%>
						<tr>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_33")%>:</td>
							<td><%=scCurSign%> <input type="text" name="cost" value="0" size="10" tabindex="205"></td>
						</tr>
						<%'End SDBA%>
					
						<% if pcv_ProductType="bto" then %>
						<tr>
							<td colspan="2"><hr></td>
						</tr>
						<tr> 
							<td>Hide Default Price:</td>
							<td>Yes <input type="checkbox" name="hidebtoprice" value="1" class="clearBorder">&nbsp;<font color="#666666">When the defaut price is very small, use this option to hide it</font></td>
						</tr>
						<tr> 
							<td>Hide default configuration:</td>
							<td>Yes <input type="checkbox" name="hidedefconfig" value="1" class="clearBorder"></td>
						</tr>
						<tr> 
							<td valign="bottom">Skip Product Details Page:</td>
							<td>Yes <input type="checkbox" name="pcv_intSkipDetailsPage" value="1" class="clearBorder">
							</td>
						</tr>
						<tr>
							<td valign="top">Disallow purchasing<br />(quoting only):</td>
							<td>
							<input type="radio" name="noprices" value="0" checked class="clearBorder">No&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="noprices" value="1" <%if cint(pnoprices)=1 then%>checked<%end if%> class="clearBorder">Yes - Show Prices&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="noprices" value="2" <%if cint(pnoprices)=2 then%>checked<%end if%> class="clearBorder">Yes - Hide Prices
						</td>
						</tr>
						<% If statusCM="0" OR scCM=1 Then %>
						<tr> 
							<td>Show Configurator+ conflict alert messages<br />(Conflict Management):</td>
							<td>Yes <input type="checkbox" name="showBtoCmMsg" value="1" class="clearBorder" checked></td>
						</tr>
						<% End If %>
						<tr>
							<td valign="top">Maximum number of selections:</td>
							<td>
							<input type="text" size="5" name="maxselect" value="0"><br>
							<i>(The number of total items selected on the product configuration page)</i>
							</td>
						</tr>
						<% end if %>
						
					</table>
				
				</div>
			<%
			'// =========================================
			'// SECOND PANEL - END
			'// =========================================

			'// =========================================
			'// THIRD PANEL - START - Product images
			'// =========================================
			%>
				<div id="tab-3" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_7")%></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2">
						<!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->
						<img src="images/sortasc_blue.gif" alt="Upload your images">&nbsp;
						<%If HaveImgUplResizeObjs=1 then%>
							<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_8")%><a href="javascript:;" onClick="pcCPWindow('uploadresize/productResizea.asp', 400, 450); return false;"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_303")%></a>.
						<% Else %>
							<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_9")%><a href="javascript:;" onClick="pcCPWindow('imageuploada_popup.asp', 400, 360)"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_303")%></a>.
						<% End If %>
							</td>
						</tr>
						<tr>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_10")%>:</td>
							<td>  
							<input type="text" name="smallImageUrl" value="" size="30" tabindex="401"><a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=smallImageUrl&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>
							&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=446"></a>
							</td>
						</tr>
						<tr> 
							<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_11")%>:</td>
							<td>
								<input type="text" name="imageUrl" value="" size="30" tabindex="402"><a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=imageUrl&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=446"></a>
							</td>
						</tr>
						<tr>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_12")%>:</td>
							<td>
								<input type="text" name="largeImageUrl" value="" size="30" tabindex="403"><a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=largeImageUrl&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=446"></a>
							</td>
						</tr>
                        <tr> 
							<td>Alt Tag Text (optional) :</td>
							<td>
								<input type="text" name="altTagText" value="" size="30" tabindex="403">
							</td>
						</tr>
                        <tr>
							<td colspan="2"><hr></td>
						</tr>
                        <tr>
							<td>Enable Image Magnifier:</td>
							<td>
								<input type="checkbox" name="MojoZoom" value="1" class="clearBorder" tabindex="404">
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=467"></a>
							</td>
						</tr>
					</table>
				</div>
			<%
			'// =========================================
			'// THIRD PANEL - END
			'// =========================================
			%>

			<%
			'// =========================================			
			'// FOURTH PANEL - START - Product Settings
			'// =========================================
			%>
			
				<div id="tab-4" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_22")%></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						
						<% if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item
						
						'// Brands - Start
						query="Select IDBrand, BrandName from Brands order by BrandName asc"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						if not rs.eof then%>
						<tr>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_35")%>:</td>
							<td>
							<select name="IDBrand" tabindex="701">
							<option value="0" selected></option>
							<% do while not rs.eof
								intIDBrand=rs("IDBrand")
								strBrandName=rs("BrandName") %>
							<option value="<%=intIDBrand%>"><%=strBrandName%></option>
										<%
										rs.MoveNext
									loop
									set rs=nothing
									%>
								</select>
							</td>
						</tr>
						<%
							else
							set rs=nothing
						%>
						<tr> 
							<td colspan="2">
								<input type=hidden name=IDBrand value="0">
							</td>
						</tr>
						<% end if
						'// Brands - End
						end if ' Hide if it's a Configurable-Only Item %>
						
						<tr> 
							<td width="30%"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_37")%>:</td>
							<td width="70%"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="checkbox" name="active" value="-1" checked class="clearBorder" tabindex="702"></td>
						</tr>

						<% if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item %>
						<tr> 
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_38")%>:</td>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="checkbox" name="hotDeal" value="-1" class="clearBorder" tabindex="703"></td>
						</tr>
						<tr>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_39")%>:</td>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="checkbox" name="showInHome" value="-1" class="clearBorder" tabindex="704"></td>
						</tr>
						<% end if ' Hide if it's a Configurable-Only Item %>

						<% 'RP ADDON-S
						If RewardsActive <> 0 Then %>
							<tr>
								<td><%=RewardsLabel%>:</td>
								<td><input type="text" name="iRewardPoints" width="10" size="20" tabindex="705"></td>
							</tr>
						<% End If
						'RP ADDON-E %>
						<tr> 
							<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_26")%>:</td>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="checkbox" name="notax" value="-1" class="clearBorder" tabindex="706"></td>
						</tr>
                        <% If ptaxAvalara = 1 Then %>
                        <tr>
							<td>Avalara Tax Code:</td>
							<td><input type="text" name="AvalaraTaxCode" value=""></td>
						</tr>
                        <% End If %>
						<% if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item %>
						<tr>
							<td nowrap><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_96")%>:</td>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="checkbox" name="hideSKU" value="1" class="clearBorder" tabindex="707"></td>
						</tr>
						<tr> 
							<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_24")%>:</td>
							<td><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="checkbox" name="formQuantity" value="-1" class="clearBorder" tabindex="708"></td>
						</tr>
						<tr>
							<td valign="top">Not for Sale Message:<br /><span class="pcSmallText">(e.g. &quot;Coming Soon&quot;)</span></td>
							<td>
								<textarea name="emailText" rows="4" cols="40" tabindex="709" onKeyUp="javascript:testchars(this,'1',250); javascript:document.getElementById('emailTextCounter').style.display='';"></textarea>
                <div id="emailTextCounter" style="margin-top: 5px; display: none; color:#666;">There are <span id="countchar1" name="countchar1" style="font-weight: bold"><%=maxlength%></span> characters left.</div>
							</td>
						</tr>
						<% end if ' Hide if it's a Configurable-Only Item %>
						<tr>
							<td valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_87")%></td>
							<td><textarea name="prdnotes" rows="6" cols="60" tabindex="710"><%=pcv_prdnotes%></textarea></td>
						</tr>
					</table>
				
				</div>
			<%
			'// =========================================			
			'// FOURTH PANEL - END - Product Settings
			'// =========================================
			%>

			<%
			'// =========================================			
			'// FIFTH PANEL - START - Inventory settings
			'// =========================================
			%>
				<div id="tab-5" class="tab-pane">
				
					<table class="pcCPcontent">

						<%'Start SDBA%>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Inventory Settings</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<script type=text/javascript>
							$pc(document).ready(function() {
								$pc("#noStock").click(function(e) {
									if ($pc(this).is(":checked")) {
										$pc("#stockOptions").hide();
									} else {
										$pc("#stockOptions").show();
									}
								});
							});	
						</script>
						<tr>
							<td>Disregard Stock:</td>
							<td><input type="checkbox" name="noStock" id="noStock" value="-1" <% if pNoStock<>0 then response.write "checked" %> class="clearBorder" tabindex="501"></td>
						</tr>
						<tr> 
							<td>Stock:</td>
							<td>  
								<%'SHW-S
								if SHWSync=1 then%><b><%=pStock%></b>&nbsp;&nbsp;&nbsp;<font color=blue><i>(Has been synchronized with SHIPWIRE)</i></font>
									<input type="hidden" name="stock" value="<%=pStock%>">
								<%else%>
									<input type="text" name="stock" id="stock" value="<%=pStock%>" size="4" tabindex="502">
								<%end if
								'SHW-E%>
								<input type="hidden" name="deliveringTime" value="<%response.write pDeliveringTime%>"> 
							</td>
						</tr>
						<%
							pcv_strDisplayStyle = ""
							If pNoStock <> 0 Then
								pcv_strDisplayStyle	= "style='display: none'"
							End If
						%>
						<tbody id="stockOptions" <%= pcv_strDisplayStyle %>>
							<tr> 
								<td valign="top">Allow back-ordering:</td>
								<td>
									<input type="radio" name="pcbackorder" value="1" <%if pcbackorder="1" then%>checked<%end if%> class="clearBorder" tabindex="503"> Yes 
									&nbsp;<input type="radio" name="pcbackorder" value="0" <%if pcbackorder<>"1" then%>checked<%end if%> class="clearBorder" tabindex="503"> No<br>
									When back-ordered, typically ships within <input type="text" size="5" value="<%=pcShipNDays%>" name="pcShipNDays" tabindex="504"> days </td>
							</tr>
							<tr> 
								<td>Low inventory notification:</td>
								<td><input type="radio" name="pcnotifystock" value="1" <%if pcnotifystock="1" then%>checked<%end if%> class="clearBorder" tabindex="505"> Yes 
									&nbsp;<input type="radio" name="pcnotifystock" value="0" <%if pcnotifystock="0" then%>checked<%end if%> class="clearBorder" tabindex="505"> No 
								<span color="#666666"><i>(Store admin is notified when inventory drops below the Reorder Level)</i></span></td>
							</tr>
							<tr> 
								<td>Reorder Level:</td>
								<td>
								<input name="pcreorderlevel" type="text" value="<%=pcreorderlevel%>" size="5" maxlength="10" tabindex="506"></td>
							</tr>
						</tbody>
						<tr> 
							<td>Minimum Quantity to Buy:</td>
							<td><input name="minimumqty" type="text" value="<%=pcv_lngMinimumQty%>" size="5" maxlength="10" tabindex="507">
								&nbsp;&nbsp;&nbsp;&nbsp;          
									<input type="checkbox" name="qtyvalidate" value="1" <%if pcv_intQtyValidate=1 then%>checked<%end if%> class="clearBorder" tabindex="508"> Force purchase of multiples of:&nbsp;<input name="multiQty" type="text" value="<%=pcv_multiQty%>" size="5" maxlength="10" tabindex="509">
							</td>
						</tr>
						<%'End SDBA%>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Weight Settings</th>
						</tr>
						
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<%
						'// WEIGHTS - Start
						if scShipFromWeightUnit="KGS" then %>
						<tr> 
							<td width="30%">Weight:</td>
							<td width="70%">
								<input type="text" name="weight_kg" value="<%=pWeight_kg%>" size="4" tabindex="520"> kg 
								<input type="text" name="weight_g" value="<%=pWeight_g%>" size="4" tabindex="521"> g
							</td>
						</tr>
						<tr>
							<td colspan="2">If this product weighs less than one gram, use the field below to specify how many units of this product it takes to weigh 1 KG. For more information, see the User Guide.</td>
						</tr>
						<tr>
							<td>Units to make 1 KG:</td>
							<td><input name="QtyToPound" type="text" id="QtyToPound" value="<%=pcv_QtyToPound%>" size="10" maxlength="10" tabindex="522"></td>
						</tr>
						<% else %>
						<tr> 
							<td width="30%">Weight:</td>
							<td width="70%">
								<input type="text" name="weight" value="<%=pWeight%>" size="4" tabindex="520"> lbs. 
								<input type="text" name="weight_oz" value="<%=pWeight_oz%>" size="4" tabindex="521"> ozs.</td>
						</tr>
						<tr>
							<td colspan="2">If this product weighs less than one ounce, use the field below to specify how many units of this product it takes to weigh 1 pound. For more information, see the User Guide.</td>
						</tr>
						<tr>
							<td>Units to make 1 lb:</td>
							<td><input name="QtyToPound" type="text" id="QtyToPound" value="<%=pcv_QtyToPound%>" size="10" maxlength="10" tabindex="522"></td>
						</tr>
						<% end if
						'// WEIGHTS - End
						%>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Shipping Settings</th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<%
						if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item
						%>
						<tr> 
							<td>Non-Shipping Item:</td>
							<td> 
								<input type="checkbox" name="noshipping" id="noShipping" value="-1" class="clearBorder" tabindex="530">
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=449"></a>
							</td>
						</tr>
						<tbody id="noShippingSettings" style="display: none">
							<tr>
								<td>Display Non-Shipping Text:</td>
								<td>
									<input type="checkbox" name="noshippingtext" value="-1" class="clearBorder" tabindex="531">
								</td>
							</tr>
						</tbody>
						<tr>
							<td colspan="2"><hr></td>
						</tr>
						<%

						end if ' Hide if it's a Configurable-Only Item
						%>
						
						<tr> 
							<td colspan="2"><strong>Oversized</strong> products shipped via <strong>UPS, USPS, or FedEx</strong></td>
						</tr>
						<tr>
							<td colspan="2">This product will be shipped as an oversized product.
								<input name="OverSizeSpec" type="radio" value="YES" class="clearBorder" tabindex="532">&nbsp;Yes 
								<input name="OverSizeSpec" type="radio" value="NO" checked class="clearBorder" tabindex="532">&nbsp;No 
								<br>
								If 'Yes', set the size below in inches. NOTE: Oversized products will always be shipped separately.
							</td>
						</tr>
						<tr> 
														<td colspan="2">
								<table class="pcCPcontent">
									<tr> 
										<td>Length:</td>
										<td width="15%"> 
											<input name="os_length" type="text" id="os_length" size="3" maxlength="3" tabindex="533">
										</td>
										<td rowspan="3" align="left" valign="top">
												<table width="100%" border="0" cellpadding="6" cellspacing="0">
													<tr>
														<td>
															Notes about shipping oversized packages with UPS, USPS, or FedEx:
															<ul>
																<li><strong>Length</strong> should always be the longest side.</li>
																<li><strong>Girth</strong> is defined as (width * 2) + (height * 2).</li>
															</ul>
														</td>
													</tr>
													<tr>
														<td><strong>Please refer to the <a href="http://wiki.productcart.com/productcart/shipping-oversized_items" target="_blank">Wiki</a> or the shipping provider's documentation for more information on oversized packages.</strong></td>
													</tr>
												</table>
                                       	</td>
                                    </tr>
                                    <tr> 
                                        <td>Width:</td>
                                        <td width="15%"> 
                                            <input name="os_width" type="text" id="os_width" size="3" maxlength="3" tabindex="534"></td>
                                    </tr>
                                    <tr> 
														<td width="11%">Height:</td>
														<td width="15%"> 
															<input name="os_height" type="text" id="os_height" size="3" maxlength="3" tabindex="535">
														</td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
						<tr>
							<td colspan="2"><hr></td>
						</tr>
						<%
						if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item

						'Start SDBA
						'Get Suppliers List
						query="Select pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName from pcSuppliers order by pcSupplier_Company asc"
						set rs=connTemp.execute(query)
						if not rs.eof then
							pcArray=rs.getRows()
							intCount=ubound(pcArray,2)
							%>
						<tr>
							<td>Supplier:</td>
							<td>
							<select name="pcIDSupplier" onChange="javascript:TestDropShipper();" tabindex="536">
							<option value="0" selected></option>
							<%For i=0 to intCount%>
								<option value="<%=pcArray(0,i)%>" <%if clng(pcIDSupplier)=clng(pcArray(0,i)) then%>selected<%end if%>><%=pcArray(1,i)%>&nbsp;<%if pcArray(2,i) & pcArray(3,i)<>"" then%>(<%=pcArray(2,i) & " " & pcArray(3,i)%>)<%end if%></option>
							<%Next%>
							</select>
							</td>
						</tr>
						<%else%>
						<tr> 
							<td colspan="2">
								<input type=hidden name="pcIDSupplier" value="0">
							</td>
						</tr>
						<%end if
						set rs=nothing

						'Get Drop-Shippers List
						query="SELECT pcDropShipper_ID,pcDropShipper_Company,pcDropShipper_FirstName,pcDropShipper_LastName,0 FROM pcDropShippers UNION (SELECT pcSupplier_ID,pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName,1 FROM pcSuppliers WHERE pcSupplier_IsDropShipper=1) ORDER BY pcDropShipper_Company ASC"
						set rs=connTemp.execute(query)
						dim pcv_ShowDSFields
						pcv_ShowDSFields=0
						if not rs.eof then
							pcv_ShowDSFields=1
						'// Allow selection only if drop-shippers exist
						%>
						
						<tr>
							<td>This product is drop-shipped:</td>
							<td> 
								<input type="radio" name="pcIsdropshipped" value="1" <%if pcIsdropshipped="1" then%>checked<%end if%> class="clearBorder" onClick="javascript:TurnOnDropShipper();" tabindex="537"> Yes 
								&nbsp;<input type="radio" name="pcIsdropshipped" value="0" <%if pcIsdropshipped<>"1" then%>checked<%end if%> class="clearBorder" onClick="javascript:TurnOffDropShipper();" tabindex="537"> No
							</td>
						</tr>
						
						<%
						'// Get list of drop-shippers
						
							pcArray=rs.getRows()
							intCount=ubound(pcArray,2)
							set rs=nothing
							%>
						<tr>
							<td>Drop-Shipper:</td>
							<td>
							<select name="pcIDDropShipper" onChange="javascript:TestSupplier()" tabindex="538">
							<option value="0" selected></option>
							<%For i=0 to intCount%>
								<option value="<%=pcArray(0,i)%>_<%=pcArray(4,i)%>" <%if (clng(pcIDDropShipper)=clng(pcArray(0,i))) AND (clng(pcArray(4,i))=pcDropShipperSupplier) then%>selected<%end if%>><%=pcArray(1,i)%>&nbsp;<%if pcArray(2,i) & pcArray(3,i)<>"" then%>(<%=pcArray(2,i) & " " & pcArray(3,i)%>)<%end if%></option>
							<%Next%>
							</select>
							</td>
						</tr>
						<tr>
							<td colspan="2"><hr></td>
						</tr>
						<%else%>
								<input type="hidden" name="pcIDDropShipper" value="0">
						<%end if
						set rs=nothing%>
						<script type=text/javascript>
							function TestDropShipper()
							{
								var tmp1=document.hForm.pcIDSupplier.value;
								try
								{
									var j=document.hForm.pcIDDropShipper.length;
									var i=0;
									var test=0;
									do
									{
										i=j-1;
										if (tmp1 + "_1" == document.hForm.pcIDDropShipper.options[i].value)
										{
											document.hForm.pcIDDropShipper.options[i].selected=true;
											document.hForm.pcIDDropShipper.disabled=true;
											document.hForm.pcIsdropshipped[0].checked=true;
											test=1;
											break;
										}
									}
									while (--j);
									if (test==0)
									{
										if (document.hForm.pcIsdropshipped[0].checked==true)
										{
											document.hForm.pcIDDropShipper.disabled=false;
										}
										var tmp1=document.hForm.pcIDDropShipper.value;
										var tmp2=tmp1.split("_");
										if (tmp2[1]==1)
										{
											document.hForm.pcIDDropShipper.options[0].selected=true;
										}
									}
								}
								catch(err)
								{
									return(true);
								}
							}
							function TestSupplier()
							{
								var tmp1=document.hForm.pcIDDropShipper.value;
								var tmp2=tmp1.split("_");
								try
								{
									var test=0;
									if (tmp2[1]=="1")
									{
										var j=document.hForm.pcIDSupplier.length;
										var i=0;
									
										do
										{
											i=j-1;
											if (tmp2[0] == document.hForm.pcIDSupplier.options[i].value)
											{
												document.hForm.pcIDSupplier.options[i].selected=true;
												document.hForm.pcIDSupplier.disabled=true;
												test=1;
												break;
											}
										}
										while (--j);
									}
									if (test==0)
									{
										if (document.hForm.pcIDSupplier.disabled==true)
										{
											document.hForm.pcIDSupplier.disabled=false;
											document.hForm.pcIDSupplier.options[0].selected=true;
										}
									}
								}
								catch(err)
								{
									return(true);
								}
					
							}
						
							function TurnOnDropShipper()
							{
								try
								{
									document.hForm.pcIDDropShipper.disabled=false;
									document.hForm.pcIDSupplier.disabled=false;
								}
								catch(err)
								{
									//Do nothing
								}
							
							}
						
							function TurnOffDropShipper()
							{
								try
								{
									document.hForm.pcIDDropShipper.disabled=true;
									document.hForm.pcIDSupplier.disabled=false;
									var tmp1=document.hForm.pcIDDropShipper.value;
									if (tmp1!="0")
									{
										var tmp2=tmp1.split("_");
										if (tmp2[1]=="1")
										{
											document.hForm.pcIDSupplier.options[0].selected=true;
										}
									}
									document.hForm.pcIDDropShipper.options[0].selected=true;
								}
								catch(err)
								{
									//Do nothing
								}
							
							}
							<% if pcv_ShowDSFields=1 then %>
							TestDropShipper();
							if (document.hForm.pcIsdropshipped[1].checked==true) TurnOffDropShipper();
							<% end if %>
						</script>
						<%
						'End SDBA

						end if ' Hide if it's a Configurable-Only Item
						%>
						<tr> 
							<td colspan="2"><strong>Shipping Surcharge</strong>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=463"></a></td>
						</tr>
						<tr>
							<td>First Unit Surcharge:</td>
							<td><input name="surcharge1" type="text" id="surcharge1" value="<%=money(pcv_Surcharge1)%>" size="10" maxlength="10" tabindex="539"></td>
						</tr>
						<tr>
							<td>Additional Unit(s) Surcharge:</td>
							<td><input name="surcharge2" type="text" id="surcharge2" value="<%=money(pcv_Surcharge2)%>" size="10" maxlength="10" tabindex="540"></td>
						</tr>
					</table>
	
				</div>
			<%
			'// =========================================
			'// FIFTH PANEL - END
			'// =========================================
			%>

			<%
			'// =========================================			
			'// SIXTH PANEL - START - Product Page Layout
			'// =========================================
			%>
				<%
				if pcv_ProductType<>"item" then ' Hide if it's a Configurable-Only Item
				%>
				<div id="tab-6" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>  
						<tr>
							<th colspan="2">Product Page Layout</th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">
								Choose a product details page layout below. 
							</td>
						</tr>
							<tr> 
								<td>Page Layout:</td>
								<td>
										<select name="displayLayout" id="displayLayout" tabindex="706">
											<option value="">Use Default</option>
											<option value="c">Two Columns-Image on Right</option>
											<option value="l">Two Columns-Image on Left</option>
											<option value="o">One-Column</option>
											<option value="t">Custom Layout</option>
										</select>
								</td>
							</tr>
							<%
								pcv_strDisplayStyle = ""
								pcv_strCustPrdDisplayStyle = "style='display: none'"
								pcv_strCustTabDisplayStyle = "style='display: none'"
							%>
							<tr>
								<td></td>
								<td>
									<ul id="customizeButtons" style="list-style-type: none; padding: 0px; margin: 0px; <%= pcv_strDisplayStyle %>">
										<li style="padding: 5px"><a href="#" id="customizeLayout"><span class="glyphicon glyphicon-cog"></span>&nbsp;Customize this Layout</a></li>
										<li style="padding: 5px"><a href="#" id="addTabsToLayout"><span class="glyphicon glyphicon-cog"></span>&nbsp;Customize this Layout w/ Tabs</a></li>
									</ul>
								</td>
							</tr>
							<!--#include file="inc_CustomPrdPage.asp"-->
					</table>
				</div>
				<%end if%>
			<%
			'// =========================================			
			'// SIXTH PANEL - END - Product Page Layout
			'// =========================================
			%>

			<%
			'// SEVENTH PANEL - START - Downloadable product
			if pcv_ProductType<>"item" then	 ' Hide for Configurable-Only Items
			%>
				<div id="tab-7" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_36")%></th>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer">
							<input type=hidden name="downloadable1" value="0">
							<input type=hidden name="urlexpire1" value="0">
							<input type=hidden name="license1" value="0">
							</td>
						</tr>
						<tr> 
							<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_37")%>&nbsp; 
							<% If statusAPP="1" OR scAPP=1 Then %>
								<input name="downloadable" type="radio" value="1" onClick="<% if pcv_ProductType="std" then %>document.hForm.apparel[1].checked='true'; document.hForm.GC[1].checked='true'; <% end if %>document.hForm.downloadable1.value='1'; document.getElementById('show_19').style.display='';<% if pcv_ProductType="std" then %> document.getElementById('show_20').style.display='none'; document.getElementById('show_21').style.display='none' <% end if %>" class="clearBorder" tabindex="801">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
							<% Else %>
								<input name="downloadable" type="radio" value="1" onClick="<% if pcv_ProductType="std" then %>document.hForm.GC[1].checked='true'; <% end if %>document.hForm.downloadable1.value='1'; document.getElementById('show_19').style.display='';<% if pcv_ProductType="std" then %> document.getElementById('show_20').style.display='none'<% end if %>" class="clearBorder" tabindex="801">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
							<% End If %>
							<input name="downloadable" type="radio" value="0" checked onClick="document.hForm.downloadable1.value='0'; document.hForm.urlexpire1.value='0'; document.hForm.license1.value='0'; document.getElementById('show_19').style.display='none';" class="clearBorder" tabindex="802">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
							</td>
						</tr>
						<tr>
							<td align="center" colspan="2">                       
							<table id="show_19" style="display:none" class="pcCPcontent">
								<tr>
									<td colspan="2"><p><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_38")%></p>
										<ul>
										<li><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_39")%><br>
										<img src="images/spacer.gif" height="15" width="1"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_40")%><%=Server.MapPath("/")%></li>
										<li><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_41")%></li>
										</ul>
									</td>
								</tr>
								<tr>
									<td colspan="2"><input type="text" name="producturl" size="70" tabindex="803"></td>
								</tr>
								<tr>
									<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_42")%>&nbsp;
										<input name="urlexpire" type="radio" value="1" onClick="document.hForm.urlexpire1.value='1';" class="clearBorder" tabindex="804">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%> 
										<input name="urlexpire" type="radio" value="0" checked onClick="document.hForm.urlexpire1.value='0'; document.hForm.expiredays.value='';" class="clearBorder" tabindex="805">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%> 
									</td>
								</tr>
								<tr>
									<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_43")%><input type="text" name="expiredays" size="5" tabindex="806">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_304")%></td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_44")%>&nbsp;
										<input name="license" type="radio" value="1" onClick="document.hForm.license1.value='1';" class="clearBorder">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%> 
										<input name="license" type="radio" value="0" checked onClick="document.hForm.license1.value='0'; document.hForm.locallg.value=''; document.hForm.remotelg.value='http://';" class="clearBorder">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%> 
									</td>
								</tr>
								<tr>
									<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_45")%></td>
								</tr>
								<tr>
									<td colspan="2"><input type="text" name="locallg" size="70" tabindex="809"></td>
								</tr>
								<tr>
									<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_46")%></td>
								</tr>
								<tr>
									<td colspan="2"> 
										<input type="text" name="remotelg" value="http://" size="70" tabindex="810"></td>
								</tr>
									<tr>
										<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_47")%></td>
									</tr>
									<tr>
										<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_48")%> (1):&nbsp;
										<input type="text" name="licenselabel1" size="36" value="" tabindex="811"></td>
									</tr>
									<tr>
										<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_48")%> (2):&nbsp;
										<input type="text" name="licenselabel2" size="36" value="" tabindex="812"></td>
									</tr>
									<tr>
										<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_48")%> (3):&nbsp;
										<input type="text" name="licenselabel3" size="36" value="" tabindex="813"></td>
									</tr>
									<tr>
										<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_48")%> (4):&nbsp;
										<input type="text" name="licenselabel4" size="36" value="" tabindex="814"></td>
									</tr>
									<tr>
										<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_48")%> (5):&nbsp;
										<input type="text" name="licenselabel5" size="36" value="" tabindex="815"></td>
									</tr>
									<tr>
										<td colspan="2" class="pcCPspacer"></td>
									</tr>
									<tr>
										<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_49")%></td>
									</tr>
								<tr>
									<td colspan="2"><textarea name="addtomail" rows="9" cols="65" tabindex="816"></textarea></td>
								</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
								<tr>
									<td colspan="2" align="center">
									<input type="button" class="btn btn-default"  name="checkbutton" value=" Verify Download URL " onClick="javascript:CheckWindow();" tabindex="817"></td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
				
				</div>
			<%
			end if ' Hide for Configurable-Only Items

			'// SEVENTH PANEL - END
			
			'// EIGHTH PANEL - START - Gift certificate
			if pcv_ProductType="std" then ' Hide if this is not a standard product

			%>
				<div id="tab-8" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_50")%></th>
						</tr>
								<tr>
									<td colspan="2" class="pcCPspacer"></td>
								</tr>
						<tr> 
							<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_51")%>&nbsp;
							<% If statusAPP="1" OR scAPP=1 Then %>
								<input name="GC" type="radio" value="1" onClick="document.hForm.apparel[1].checked='true'; document.hForm.downloadable[1].checked='true'; document.hForm.downloadable1.value='0'; document.hForm.urlexpire1.value='0'; document.hForm.license1.value='0'; document.getElementById('show_19').style.display='none'; document.getElementById('show_20').style.display=''; document.getElementById('show_21').style.display='none';" class="clearBorder" tabindex="901">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
							<% Else %>
								<input name="GC" type="radio" value="1" onClick="document.hForm.downloadable[1].checked='true'; document.hForm.downloadable1.value='0'; document.hForm.urlexpire1.value='0'; document.hForm.license1.value='0'; document.getElementById('show_19').style.display='none';document.getElementById('show_20').style.display=''" class="clearBorder" tabindex="901">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;
							<% End If %>
							<input name="GC" type="radio" value="0" checked onClick="document.getElementById('show_20').style.display='none'" class="clearBorder" tabindex="902">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
							</td>
						</tr>
						<tr>
							<td colspan="2">                       
							<table id="show_20" style="display:none" class="pcCPcontent">
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_52")%>:</td>
							</tr>
							<tr>
								<td align="right">
									<input name="GCExp" type="radio" value="0" checked class="clearBorder" tabindex="903">
								</td>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_53")%></td>
							</tr>
							<tr>
								<td align="right" valign="top">
									<input name="GCExp" type="radio" value="1" class="clearBorder" tabindex="904">
								</td>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_54")%>&nbsp;<input type="text" name="GCExpDate" size="25" tabindex="905">&nbsp;(<i></i><%=dictLanguageCP.Item(Session("language")&"_cpCommon_233")%>: <%if scDateFrmt="DD/MM/YY" then%><%=dictLanguageCP.Item(Session("language")&"_cpCommon_234")%><%else%><%=dictLanguageCP.Item(Session("language")&"_cpCommon_235")%><%end if%></i>)
								</td>
							</tr>
							<tr>
								<td align="right" valign="top">
									<input name="GCExp" type="radio" value="2" class="clearBorder" tabindex="906">
								</td>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_55")%><input type="text" name="GCExpDay" size="5" tabindex="907">
								</td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_56")%>:&nbsp;<input name="GCEOnly" type="checkbox" value="1" checked class="clearBorder" tabindex="908">
								</td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_57")%></td>
							</tr>
							<tr>
								<td align="right">
									<input name="GCGen" type="radio" value="0" checked class="clearBorder" tabindex="909">
								</td>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_58")%></td>
							</tr>
							<tr>
								<td align="right" valign="top">
									<input name="GCGen" type="radio" value="1" class="clearBorder" tabindex="910">
								</td>
								<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_59")%><input type="text" name="GCGenFile" size="53" tabindex="911">
									<div class="pcCPnotes"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_60")%></div>
								</td>
							</tr>
							</table>
							</td>
						</tr>
					</table>
				
				</div>
			<%
			end if ' Hide if this is not a standard product
			'// EIGHTH PANEL - END		


			'// APPAREL PANEL - START
			If statusAPP="1" OR scAPP=1 Then
				if pcv_ProductType="std" then ' Hide if this is not a standard product
				%>
					<div id="tabs-12" class="tab-pane">
					
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr> 
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_70")%></th>
							</tr>
							<tr> 
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_71")%>&nbsp; 
									<input name="apparel" type="radio" value="1" onClick="document.hForm.downloadable[1].checked='true'; document.hForm.GC[1].checked='true'; document.hForm.downloadable1.value='0'; document.hForm.urlexpire1.value='0'; document.hForm.license1.value='0'; document.getElementById('show_19').style.display='none'; document.getElementById('show_20').style.display='none';document.getElementById('show_21').style.display='';" class="clearBorder">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>
									<input name="apparel" type="radio" value="0" checked onClick="document.getElementById('show_21').style.display='none';" class="clearBorder">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								</td>
							</tr>
							<tr>
								<td colspan="2">                       
								<table id="show_21" style="display:none" class="pcCPcontent">
									<tr>
										<td width="100%" nowrap colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_72")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top">
											<input type="radio" name="showstockmsg" value="0" style="float: right" checked class="clearBorder">
										</td>
										<td width="70%" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_73")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top">
											<input type="radio" name="showstockmsg" value="2" style="float: right" class="clearBorder">
										</td>
										<td width="70%" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_73a")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top">
											<input type="radio" name="showstockmsg" value="3" style="float: right" class="clearBorder">
										</td>
										<td width="70%" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_73b")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top">
											<input type="radio" name="showstockmsg" value="1" style="float: right" class="clearBorder">
										</td>
										<td width="70%" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_74")%><br>
											<input type="text" name="stockmsg" size="40"><br>
											<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_75")%>
										</td>
									</tr>
									<tr>
										<td colspan="2"><hr /></td>
									</tr>
									<tr>
										<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_76")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="middle">
											<input type="radio" name="pcv_ApparelRadio" value="0" style="float: right" checked class="clearBorder">
										</td>
										<td width="70%" valign="middle"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_77")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="middle">
											<input type="radio" name="pcv_ApparelRadio" value="1" style="float: right" class="clearBorder">
										</td>
										<td width="70%" valign="middle"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_78")%></td>
									</tr>
									<tr>
										<td colspan="2"><hr /></td>
									</tr>
									<tr>
										<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_79")%></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top" align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_80")%></td>
										<td width="70%" valign="top"><input type="text" name="sizelink" cols="40" value="Size Chart"></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top" align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_81")%></td>
										<td width="70%" valign="top"><textarea rows="9" name="sizeinfo" cols="40"></textarea></td>
									</tr>
									<tr>
										<td width="30%" nowrap valign="top" align="right"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_82")%></td>
										<td width="70%" valign="top"><input type="text" name="sizeimg" size="40"><br>
										<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_83")%>&nbsp;<a href="#" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_84")%></a></td>
									</tr>
									<tr>
										<td width="30%" nowrap align="right" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_85")%></td>
										<td width="70%" valign="top">
											<input type="text" name="sizeurl" size="40" value="http://">
											<br><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_86")%>
										</td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
						
					</div>
				<%
				end if
			End If
			'// APPAREL PANEL - END


			'// NINTH PANEL - START - Custom fields
			if pcv_ProductType<>"item" then	 ' Hide for Configurable-Only Items
			%>
				<div id="tab-9" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>  
						<tr>
							<th colspan="2">Custom Search Fields</th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2">This tab will allow the store manager to view, add, and edit custom search fields associated with this product.</td>
						</tr>
						<tr>
							<td colspan="2">
								<%tmpJSStr=""
								tmpJSStr=tmpJSStr & "var SFID=new Array();" & vbcrlf
								tmpJSStr=tmpJSStr & "var SFNAME=new Array();" & vbcrlf
								tmpJSStr=tmpJSStr & "var SFVID=new Array();" & vbcrlf
								tmpJSStr=tmpJSStr & "var SFVALUE=new Array();" & vbcrlf
								tmpJSStr=tmpJSStr & "var SFVORDER=new Array();" & vbcrlf
								intCount=-1
								tmpJSStr=tmpJSStr & "var SFCount=" & intCount & ";" & vbcrlf%>
								<script type=text/javascript>
									<%=tmpJSStr%>
									function CreateTable()
									{
										var tmp1="";
										var tmp2="";
										var i=0;
										var found=0;
										tmp1='<table class="pcCPcontent"><tr><td></td><td nowrap><strong>Text to display</strong></td><td><strong>Value</strong></td></tr>';
										for (var i=0;i<=SFCount;i++)
										{
											found=1;
											tmp1=tmp1 + '<tr><td align="right"><a href="javascript:ClearSF(SFID['+i+']);"><img src="../pc/images/minus.jpg" alt="Remove" border="0"></a></td><td width="275" nowrap>'+SFNAME[i]+'</td><td width="100%">'+SFVALUE[i]+'</td></tr>';
											if (tmp2=="") tmp2=tmp2 + "||";
											tmp2=tmp2 + "^^^" + SFID[i] + "^^^" + SFVID[i] + "^^^" + SFVALUE[i] + "^^^" + SFVORDER[i] + "^^^||"
										}
										tmp1=tmp1+'</table>';
										if (found==0) tmp1="<br><b>No search fields are assigned to this product</b><br><br>";
										document.getElementById("stable").innerHTML=tmp1;
										document.hForm.SFData.value=tmp2;
									}
									function ClearSF(tmpSFID)
									{
										var i=0;
										for (var i=0;i<=SFCount;i++)
										{
											if (SFID[i]==tmpSFID)
											{
												removedArr = SFID.splice(i,1);
												removedArr = SFNAME.splice(i,1);
												removedArr = SFVID.splice(i,1);
												removedArr = SFVALUE.splice(i,1);
												removedArr = SFVORDER.splice(i,1);
												SFCount--;
												break;
											}
										}
										CreateTable();
									}
					
									function AddSF(tmpSFID,tmpSFName,tmpSVID,tmpSValue,tmpSOrder)
									{
										if (tmpSValue!="")
										{
											var i=0;
											var found=0;
											for (var i=0;i<=SFCount;i++)
											{
												if (SFID[i]==tmpSFID)
												{
													SFVID[i]=tmpSVID;
													SFVALUE[i]=tmpSValue;
													SFVORDER[i]=tmpSOrder;
													found=1;
													break;
												}
											}
											if (found==0)
											{
												SFCount++;
												SFID[SFCount]=tmpSFID;
												SFNAME[SFCount]=tmpSFName;
												SFVID[SFCount]=tmpSVID;
												SFVALUE[SFCount]=tmpSValue;
												SFVORDER[SFCount]=tmpSOrder;
											}
											CreateTable();
										}
									}
								</script>
								<span id="stable" name="stable"></span>
								<input type="hidden" name="SFData" value="">
								<%query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields WHERE pcSearchFieldCPShow=1 ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
								set rs=Server.CreateObject("ADODB.Recordset")
								set rs=conntemp.execute(query)
								if not rs.eof then
									set pcv_tempFunc = new StringBuilder
									pcv_tempFunc.append "<script type=text/javascript>" & vbcrlf
									pcv_tempFunc.append "function CheckList(cvalue) {" & vbcrlf
									pcv_tempFunc.append "if (cvalue==0) {" & vbcrlf
									pcv_tempFunc.append "var SelectA = document.hForm.SearchValues;" & vbcrlf
									pcv_tempFunc.append "SelectA.options.length = 0; }" & vbcrlf
					
									set pcv_tempList = new StringBuilder
									pcv_tempList.append "<select name=""customfield"" onchange=""javascript:document.hForm.newvalue.value='';document.hForm.neworder.value='0';CheckList(document.hForm.customfield.value);"">" & vbcrlf
					
									pcArray=rs.getRows()
									intCount=ubound(pcArray,2)
									set rs=nothing
					
									For i=0 to intCount
										pcv_tempList.append "<option value=""" & pcArray(0,i) & """>" & replace(pcArray(1,i),"""","&quot;") & "</option>" & vbcrlf
										query="SELECT idSearchData,pcSearchDataName FROM pcSearchData WHERE idSearchField=" & pcArray(0,i) & " ORDER BY pcSearchDataOrder ASC,pcSearchDataName ASC;"
										set rs=connTemp.execute(query)
										if not rs.eof then
											tmpArr=rs.getRows()
											LCount=ubound(tmpArr,2)
											pcv_tempFunc.append "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
											pcv_tempFunc.append "var SelectA = document.hForm.SearchValues;" & vbcrlf
											pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
											For j=0 to LCount
												pcv_tempFunc.append "SelectA.options[" & j & "]=new Option(""" & replace(tmpArr(1,j),"""","\""") & """,""" & tmpArr(0,j) & """);" & vbcrlf
											Next
											pcv_tempFunc.append "}" & vbcrlf
										else
											pcv_tempFunc.append "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
											pcv_tempFunc.append "var SelectA = document.hForm.SearchValues;" & vbcrlf
											pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
											pcv_tempFunc.append "SelectA.options[" & 0 & "]=new Option("""",""""); }" & vbcrlf
										end if
									Next
			
									pcv_tempList.append "</select>" & vbcrlf
									pcv_tempFunc.append "}" & vbcrlf
									pcv_tempFunc.append "</script>" & vbcrlf
									
									pcv_tempList=pcv_tempList.toString
									pcv_tempFunc=pcv_tempFunc.toString
									%>
									<br><br>
									<hr>
									<table class="pcCPcontent" style="width:auto;">
										<tr>
											<td colspan="2"><a name="2"></a><b><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_91")%></b></td>
										</tr>
										<tr>
											<td width="20%"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_92")%></td>
											<td width="80%">
											<%=pcv_tempList%>&nbsp;<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_93")%>&nbsp;
											<select name="SearchValues">
											</select>
											<%=pcv_tempFunc%>
											<script type=text/javascript>
												CheckList(document.hForm.customfield.value);
											</script>
											&nbsp;<a href="javascript:AddSF(document.hForm.customfield.value,document.hForm.customfield.options[document.hForm.customfield.selectedIndex].text,document.hForm.SearchValues.value,document.hForm.SearchValues.options[document.hForm.SearchValues.selectedIndex].text,0);"><img src="../pc/images/plus.jpg" alt="Add" border="0"></a>
											</td>
										</tr>
										<tr>
											<td><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_94")%></td>
											<td>
												<input type="text" value="" name="newvalue" size="30">
                        <input type="hidden" value="0" name="neworder">
												&nbsp;<a href="javascript:AddSF(document.hForm.customfield.value,document.hForm.customfield.options[document.hForm.customfield.selectedIndex].text,-1,document.hForm.newvalue.value,document.hForm.neworder.value);"><img src="../pc/images/plus.jpg" alt="Add" border="0"></a>
											</td>
										</tr>
										<tr>
											<td colspan="2">
												<b><u><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_88")%></u></b> <i><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_89")%></i>
												<br><br>
											</td>
										</tr>
										</table>
								<%else
									query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
									set rs=Server.CreateObject("ADODB.Recordset")
									set rs=conntemp.execute(query)
									if not rs.eof then%>
										<a href="ManageSearchFields.asp">Click here</a> to manage custom search fields.</a>
									<%else%>
										<a href="ManageSearchFields.asp">Click here</a> to add new product custom search field.</a>
									<%end if
									set rs=nothing%>
								<%end if%>
								<script type=text/javascript>CreateTable();</script>
							</td>
						</tr>
					</table>
				
				</div>
				<% '// NINTH PANEL - END - Custom fields %>

				<% '// TENTH PANEL - START - Google Shopping Settings %>
				<div id="tab-10" class="tab-pane">
				
					<table class="pcCPcontent">
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>  
						<tr>
							<th colspan="2">Google Shopping Settings</th>
						</tr>
						<tr>
							<td class="pcCPspacer" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2"><b>Google Product Category</b></td>
						</tr>
						<tr>
							<td><input type="radio" name="pcv_GPC" value="0" checked class="clearBorder"></td>
							<td>Use the Product’s current category assignment for Google Shopping. (Set by default)</td>
						</tr>
						<tr>
							<td><input type="radio" name="pcv_GPC" value="1" class="clearBorder"></td>
							<td>Use a Google Product Category Attribute</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td>
								<select name="pcv_GCat">
									<option value="" selected>Select one... </option>
									<option value="Apparel &amp; Accessories">Apparel &amp; Accessories</option>
									<option value="Apparel &amp; Accessories &gt; Clothing">Apparel &amp; Accessories &gt; Clothing</option>
									<option value="Apparel &amp; Accessories &gt; Shoes">Apparel &amp; Accessories &gt; Shoes</option>
									<option value="Media &gt; Books">Media &gt; Books</option>
									<option value="Media &gt; DVDs &amp; Movies">Media &gt; DVDs &amp; Movies</option>
									<option value="Media &gt; Music">Media &gt; Music</option>
									<option value="Software &gt; Video Game Software">Software &gt; Video Game Software</option>
								</select>
						<tr>
							<td>&nbsp;</td>
							<td>
								Or using other: <input type="text" name="pcv_GCatO" size="35" value=""><br>
								<i><u>Note:</u> To get correct Google's Product Taxonomy, <a href="http://support.google.com/merchants/bin/answer.py?hl=en&answer=1705911" target="_blank">click here</a></i>
						 	</td>
						</tr>
						<tr>
							<td colspan="2"><hr width="95%"></td>
						</tr>
						<tr>
							<td colspan="2"><b>Google Apparel Product Attributes</b></td>
						</tr>
						<tr>
							<td>Gender:</td>
							<td>
								<select name="pcv_GGen">
									<option value="" selected>Select one... </option>
									<option value="male">Male</option>
									<option value="female">Female</option>
									<option value="unisex">Unisex</option>
								</select>
							</td>
						</tr>
						<tr>
							<td>Age Group:</td>
							<td>
								<select name="pcv_GAge">
									<option value="" selected>Select one... </option>
									<option value="adult">Adult</option>
									<option value="kids">Kids</option>
								</select>
							</td>
						</tr>
						<tr>
							<td>Size:</td>
							<td>
								<input type="text" name="pcv_GSize" size="35" value="">
							</td>
						</tr>
						<tr>
							<td>Color:</td>
							<td>
								<input type="text" name="pcv_GColor" size="35" value="">
							</td>
						</tr>
						<tr>
							<td>Pattern:</td>
							<td>
								<input type="text" name="pcv_GPat" size="35" value="">
							</td>
						</tr>
						<tr>
							<td>Material:</td>
							<td>
								<input type="text" name="pcv_GMat" size="35" value="">
							</td>
						</tr>
					</table>
				
				</div>
				<% '// TENTH PANEL - END - Google Shopping Settings %>
				
			<%
			end if	 ' Hide for Configurable-Only Items
			'// TENTH PANEL - END
			%>
			
			</div>
		
		<%
		'// TABBED PANELS - MAIN DIV END
		%>
        </div>
    </div>
	<div style="clear: both;">&nbsp;</div>
		<script type=text/javascript>
			var tab = window.location.hash;
			$pc('.nav-tabs').on('click', 'a', function() {
				window.location.hash = $pc(this).attr('href');
			});

			$pc('#TabbedPanels1 a[href="' + tab + '"]').tab('show');
		</script>
	
</form>    
        
<% else
	'count categories, if too many are present, show alternate page
	dim iCatCnt
	iCatCnt=0
	query="SELECT Count(*) As CatTotal FROM categories WHERE tier=3 or tier=4;"
	set rstemp=conntemp.execute(query)
	if not rstemp.eof then
		iCatCnt=rstemp("CatTotal")
	end if
	set rstemp=nothing
	
	if iCatCnt<200 then
		call closeDb()
		response.redirect "addProduct.asp?catCnt=100&prdType="&pcv_ProductType
	else %>
		<form method="post" name="RootCatForm" action="addProduct.asp?prdType=<%=pcv_ProductType%>" class="pcForms">
			<input type="hidden" name="catCnt" value="200">
			<table class="pcCPcontent">
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr> 
					<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_61")%></th>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr> 
					<td> 
						<select name="idRootCategory">
						<%' get leaf categories
						query="SELECT idCategory, categoryDesc, idparentCategory FROM categories WHERE idparentCategory=1 ORDER BY categoryDesc"
						set rstemp=conntemp.execute(query)
						if err.number <> 0 then
							
							call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&Err.Description) 
						end if
						if  rstemp.eof then 
							
							call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("<b><i>No categories defined:</i></b><br><br>Before you can add products to your store, you need to add at least one category.<br><br><a href=instCata.asp>Click here to add categories to your store.</a>") 
						end if
						dim parent
						do until rstemp.eof 
							idcategory=rstemp("idCategory")
							idparentCategory=rstemp("idparentCategory")
							categoryDesc=rstemp("categoryDesc") %>
							<option value='<%=idcategory%>'><%=categoryDesc%></option>
							<% rstemp.movenext
						loop
						%>
						</select>
					</td>
				</tr>
				<tr> 
					<td><hr></td>
				</tr>
				<tr> 
					<td>
						<input type="submit" name="Submit" value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_105")%>" class="btn btn-primary">
						&nbsp; 
						<input type="button" class="btn btn-default"  value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_103")%>" onClick="location.href='LocateProducts.asp?cptype=0'">
						&nbsp; 
						<input type="button" class="btn btn-default"  value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_104")%>" onClick="location.href='manageCategories.asp'">
					</td>
				</tr>
			</table>			
			
		</form>
	<% end if
end if %><!--#include file="AdminFooter.asp"-->