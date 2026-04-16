<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="CustLIv.asp"-->
<%
pIdCustomer=session("idCustomer")

if (request("action")="add") and (request("rewrite")="0") then
	getype=getUserInput(request("etype"),0)
	gename=getUserInput(request("ename"),0)
	gedate=getUserInput(request("edate"),0)
	if gedate="" then
		gedate="01/01/1900"
	end if
	gedelivery=getUserInput(request("edelivery"),0)
	if gedelivery="" then
		gedelivery="0"
	end if
	gemyaddr=getUserInput(request("emyaddr"),0)
	if gemyaddr="" then
		gemyaddr="0"
	end if
	gehide=getUserInput(request("ehide"),0)
	if gehide="" then
		gehide="0"
	end if
	geHideAddress=getUserInput(request("eHideAddress"),0)
	if geHideAddress="" then
		geHideAddress="0"
	end if
	genotify=getUserInput(request("enotify"),0)
	if genotify="" then
		genotify="0"
	end if
	geincgc=getUserInput(request("eincgc"),0)
	if geincgc="" then
		geincgc="0"
	end if	
	geactive=getUserInput(request("eactive"),0)
	if geactive="" then
		geactive="0"
	end if
		
	Do while mytest=0
		myTest=0
		Tn1=""
		For w=1 to 16
			Randomize
			myC=Fix(3*Rnd)
			Select Case myC
			Case 0: 
				Randomize
				Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
			Case 1: 
				Randomize
				Tn1=Tn1 & Cstr(Fix(10*Rnd))
			Case 2: 
				Randomize
				Tn1=Tn1 & Chr(Fix(26*Rnd)+97)		
			End Select
		Next
		query="select pcEv_IDEvent from pcEvents where pcEv_Code='" & Tn1 & "'"
		set rstemp=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		if rstemp.eof then
			myTest=1
		end if
		
		set rstemp=nothing
	Loop
	
	geCode=Tn1
	
	if SQL_Format="1" then
		geDate=(day(geDate)&"/"&month(geDate)&"/"&year(geDate))
	else
		geDate=(month(geDate)&"/"&day(geDate)&"/"&year(geDate))
	end if
	query="INSERT INTO pcEvents (pcEv_IDCustomer,pcEv_Type,pcEv_Name,pcEv_Date,pcEv_Delivery,pcEv_MyAddr,pcEv_Hide,pcEv_Notify,pcEv_IncGcs,pcEv_Active,pcEv_Code,pcEv_HideAddress) values (" & pIDCustomer & ",'" & getype & "',N'" & gename & "','" & gedate & "'," & gedelivery & "," & gemyaddr & "," & gehide & "," & genotify & "," & geincgc & "," & geactive & ",'" & gecode & "'," & geHideAddress & ")"
	set rstemp=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	set rstemp=nothing
	
	query="select pcEv_IDEvent from pcEvents where pcEv_IDCustomer=" & pIDCustomer & " order by pcEv_IDEvent desc"
	set rstemp=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	gIDEvent=rstemp("pcEv_IDEvent")
	
	set rstemp=nothing
	
	if geincgc="1" then
	
		query="select IDProduct from Products where pcprod_GC=1"
		set rstemp=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

		do while not rstemp.eof
			IDProduct=rstemp("IDProduct")
			query="insert into pcEvProducts (pcEP_IDEvent,pcEP_IDProduct,pcEP_GC) values (" & gIDEvent & "," & IDProduct & ",1)"
			set rs1=conntemp.execute(query)
			rstemp.MoveNext
		loop
		
		set rstemp=nothing

	end if	
	
	'*******************************************************************************************************
	' Start - Send Email Notification to the Store Administrator only when a customer creates new Gift Registry
	'*******************************************************************************************************
	'The flag below can be used to Turn ON/OFF this feature.
	'Set the flag to Y to Turn ON and N to Turn OFF.
	pc_SendNewGREmailToStoreAdmin = "Y"
	if (pIdCustomer <> 0 and pIdCustomer <> "") then
		if pc_SendNewGREmailToStoreAdmin = "Y" then
			query="SELECT name, lastname, email from customers WHERE idCustomer=" & pIdCustomer & ";"			
			set rsNewGR=server.CreateObject("ADODB.RecordSet")
			set rsNewGR=conntemp.execute(query)
			if not rsNewGR.eof then
				pcStrName=rsNewGR("name")
				pcStrLastName=rsNewGR("lastname")
				pcStrCustomerName = pcStrName & " " & pcStrLastName
				pcStrEmail=rsNewGR("email")
				pcDate = ShowDateFrmt(gedate)
				
				pcStrSubject=dictLanguage.Item(Session("language")&"_NewGREmailToStoreAdmin_Subject") & " - " & pcStrCustomerName
				pcStrSubject=replace(pcStrSubject,"''","'")				

				pcStrBody=dictLanguage.Item(Session("language")&"_NewGREmailToStoreAdmin_Body") & VBcrlf & VBcrlf
				pcStrBody=pcStrBody & dictLanguage.Item(Session("language")&"_NotifyGR_3") & " " & pcStrCustomerName & VBcrlf
				pcStrBody=pcStrBody & dictLanguage.Item(Session("language")&"_NotifyGR_4") & " " & pcStrEmail & VBcrlf
				pcStrBody=pcStrBody & dictLanguage.Item(Session("language")&"_GRDetails_1b") & " " & pcDate & VBcrlf
				pcStrBody=pcStrBody & dictLanguage.Item(Session("language")&"_GRDetails_1c") & " " & gename & VBcrlf
				pcStrBody=replace(pcStrBody,"''","'")
				
				call sendmail (pcStrCustomerName, pcStrEmail, scFrmEmail, pcStrSubject, pcStrBody)
			end if
			rsNewGR.Close
			SET rsNewGR = NOTHING
		end if
	end if
	'*******************************************************************************************************
	' End - Send Email Notification to the Store Administrator only when a customer creates new Gift Registry
	'*******************************************************************************************************	
	
	response.redirect "ggg_manageGRs.asp"

end if

IF request("rewrite")="1" then
	getype=getUserInput(request("etype"),0)
	gename=getUserInput(request("ename"),0)
	gedate=getUserInput(request("edate"),0)
	if gedate="" then
		gedate="01/01/1900"
	end if
	gedelivery=getUserInput(request("edelivery"),0)
	if gedelivery="" then
		gedelivery="0"
	end if
	gemyaddr=getUserInput(request("emyaddr"),0)
	if gemyaddr="" then
		gemyaddr="0"
	end if
	gehide=getUserInput(request("ehide"),0)
	if gehide="" then
		gehide="0"
	end if
	geHideAddress=getUserInput(request("eHideAddress"),0)
	if geHideAddress="" then
		geHideAddress="0"
	end if
	genotify=getUserInput(request("enotify"),0)
	if genotify="" then
		genotify="0"
	end if
	geincgc=getUserInput(request("eincgc"),0)
	if geincgc="" then
		geincgc="0"
	end if	
	geactive=getUserInput(request("eactive"),0)
	if geactive="" then
		geactive="0"
	end if

ELSE
	geactive="1"
END IF

%>
<!--#include file="header_wrapper.asp"-->
<script type=text/javascript>
function win(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=600,height=550')
	myFloater.location.href=fileName;
	checkwin();
	}
function checkwin()
{

if (myFloater.closed)
{
document.Form1.submit();
}
else
{
setTimeout('checkwin()',500);
}
}

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
   if ((err == 0) && (day != "") && (month != "") && (year != "") && (seperator != ""))
   {
		var EDate=new Date(year, month-1, day);
		var NDate=new Date();
		if (EDate<NDate) err=1;
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
	
function Form1_Validator(theForm)
{
	if (theForm.ename.value == "")
  	{
			alert("Please enter a value for this field.");
		    theForm.ename.focus();
		    return (false);
	}

	if (theForm.edate.value == "")
  	{
			alert("Please enter a valid date for this field.");
		    theForm.edate.focus();
		    return (false);
	}
	
	if (check_date(theForm.edate) == false)
  	{
		alert("Please enter a valid date for this field.");
	    theForm.edate.focus();
	    return (false);
	}
	
return (true);
}
</script>
<div id="pcMain">
	<div class="pcMainContent">
		<form method="post" name="Form1" action="ggg_instGR.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
			<h1><%= dictLanguage.Item(Session("language")&"_instGR_1")%></h1>

			<% If msg<>"" then %>
				<div class="pcErrorMessage"><%=msg%></div>
			<% end if %>

			<% '// Event Type %>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_instGR_2")%></div>
				<div class="pcFormField"><input type=text name="etype" size="30" value="<%=getype%>"></div>
			</div>

			<% '// Event Name %>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_instGR_3")%></div>
				<div class="pcFormField"><input type=text name="ename" size="30" value="<%=gename%>"></div>
			</div>

			<% '// Event Date %>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_instGR_4")%></div>
				<div class="pcFormField">
					<input type=text name="edate" class="datepicker" size="30" value="<%=gedate%>"> (<i><%= dictLanguage.Item(Session("language")&"_instGR_4a")%>
					<%if scDateFrmt="DD/MM/YY" then%>DD/MM/YY<%else%>MM/DD/YY<%end if%></i>)
				</div>
			</div>

			<% '// Preferred Delivery %>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_instGR_5")%></div>
				<div class="pcFormField">
					<input type=radio name="edelivery" value="1" class="clearBorder" <%if gedelivery="1" then%>checked<%end if%>>
					<%= dictLanguage.Item(Session("language")&"_instGR_6")%>
					&nbsp;<select name="emyaddr">
					<%
					myTest=0

					query="SELECT address,city,state,statecode,zip,countrycode,shippingAddress, shippingCity, shippingState, shippingStateCode, shippingZip, shippingCountryCode, shippingCompany, shippingAddress2 FROM customers WHERE idCustomer=" &session("idCustomer")
					set rstemp=conntemp.execute(query)

					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
		
					pshippingAddress=rstemp("shippingAddress")
					pshippingZip=rstemp("shippingZip")
					pshippingState=rstemp("shippingState")
					pshippingStateCode=rstemp("shippingStateCode") 
					pShippingCity=rstemp("shippingCity")
					pshippingCountryCode=rstemp("shippingCountryCode")
					pshippingCompany=rstemp("shippingCompany")
					pshippingAddress2=rstemp("shippingAddress2")
				
					myTest=1
					session("paddress")=ucase(rstemp("address"))
					session("pcity")=ucase(rstemp("city"))
					session("pstate")=ucase(rstemp("state") & rstemp("statecode"))
					set rstemp=nothing
				
					session("pshipadd")=pshippingAddress
					session("pshipZip")=pshippingZip
					session("pshipState")=pshippingState
					session("pshipStateCode")=pshippingStateCode 
					session("pshipCity")=pShippingCity
					session("pshipCountryCode")=pshippingCountryCode
					session("pshipCom")=pshippingCompany
					session("pshipadd2")=pshippingAddress2
					%>
					<option value="0" <%if gemyaddr="0" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_CustSAmanage_10")%></option>
					<%
					query="SELECT idRecipient, recipient_NickName,recipient_Address,recipient_City,recipient_State,recipient_StateCode FROM recipients WHERE idCustomer=" &session("idCustomer")
					set rstemp=conntemp.execute(query)
	
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
		
					do while not rstemp.eof
						myTest=1
						IDre=rstemp("idRecipient")
						reFullName=trim(rstemp("recipient_NickName"))
						reShipAddr=ucase(rstemp("recipient_Address"))
						reShipCity=ucase(rstemp("recipient_City"))
						reShipState=ucase(rstemp("recipient_State") & rstemp("recipient_StateCode"))
	
						myTest1=0

						if (reShipAddr=session("pAddress")) and (reShipState=session("pState")) and (reShipCity=session("pCity")) and (reFullName="") then 
							myTest1=1
						end if
						if (reShipAddr=ucase(session("pShipAdd"))) and (reShipState=Ucase(session("pShipState")&session("pShipStateCode")) ) and (reShipCity=ucase(session("pShipCity"))) and (reFullName="") then 
							myTest1=1
						end if			
						if MyTest1=0 then
							if trim(reFullName)<>"" then
							else
								reFullName="No shipping name specified"
							end if%>
							<option value="<%=IDre%>" <%if clng(gemyaddr)=clng(IDre) then%>selected<%end if%>><%=reFullName%></option>
						<%end if
						rstemp.movenext
					loop
					set rstemp=nothing%>
					</select>
					<br>
					<a href="javascript:win('CustAddShipPop.asp');"><%= dictLanguage.Item(Session("language")&"_instGR_8")%></a>
					<br>
					<input type=radio name="edelivery" value="0" class="clearBorder" <%if gedelivery<>"1" then%>checked<%end if%>>
					<%= dictLanguage.Item(Session("language")&"_instGR_9")%>
				</div>
			</div>

			<div class="pcSpacer"></div>

			<% '// Other Settings %>
			<div class="pcFormItem">
				<div class="pcFormLabel">
					<%= dictLanguage.Item(Session("language")&"_instGR_5b")%>
				</div>
				<div class="pcFormField">
					<input type=checkbox name="ehide" value="1" class="clearBorder" <%if gehide="1" then%>checked<%end if%>>
					<%= dictLanguage.Item(Session("language")&"_instGR_10")%>
					<br />
					<i><%= dictLanguage.Item(Session("language")&"_instGR_10a")%></i>

					<br /><br />

					<input type=checkbox name="eHideAddress" value="1" class="clearBorder" <%if geHideAddress="1" then%>checked<%end if%>>
					<%= dictLanguage.Item(Session("language")&"_instGR_17")%>

					<br />

					<input type=checkbox name="enotify" value="1" class="clearBorder" <%if genotify="1" then%>checked<%end if%>>
					<%= dictLanguage.Item(Session("language")&"_instGR_11")%>

					<br />

					<input type=checkbox name="eincgc" value="1" class="clearBorder" <%if geincgc="1" then%>checked<%end if%>>
					<%= dictLanguage.Item(Session("language")&"_instGR_12")%>

					<br />

					<input type=checkbox name="eactive" value="1" class="clearBorder" <%if geactive="1" then%>checked<%end if%>>
					<%= dictLanguage.Item(Session("language")&"_instGR_13")%>
				</div>
			</div>
				
			<div class="pcSpacer"></div>

			<div class="pcFormButtons">
				<button class="pcButton pcButtonContinue" id="submit" name="submit" value="<%= dictLanguage.Item(Session("language")&"_instGR_12")%>" onClick="document.Form1.rewrite.value='0';">
					<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>">
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
				</button>

				<a class="pcButton pcButtonBack" href="ggg_manageGRs.asp">
					<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
				</a>

				<input type=hidden name="rewrite" value="1">
			</div>
		</form>

	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
