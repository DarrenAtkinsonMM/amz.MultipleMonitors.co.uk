<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/UPSconstants.asp"-->
<!--#include file="../includes/pcUPSClass.asp"-->

<% Dim objUPSXmlDoc, objUPSStream, strFileName, GraphicXML
Dim UPS_AccessRequest, UPS_PostData, objUPSClass, objOutputXMLDoc, srvUPSXmlHttp, UPS_result, UPS_URL, pcv_strErrorMsg, pcv_strAction

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
%>
<!--#include file="inc_headerV5.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->

<body id="pcPopup">
<% pcPageName="pcUPSTimeInTransit.asp"

'// SET THE UPS OBJECT
set objUPSClass = New pcUPSClass

pcv_isCountryCodeRequired=false
%>
<div id="pcMain">
	<div class="pcMainContent">
		<% if len(session("pcSFCity"))<1 then
			session("pcSFCity")=getUserInput(request("sCity"),0)
		end if 
		if len(session("pcSFCountryCode"))<1 then
				session("pcSFCountryCode")=getUserInput(request("sCountry"),0)
		end if 
		if len(session("pcSFPostalCode"))<1 then
				session("pcSFPostalCode")=getUserInput(request("sPC"),0)
		end if 
		if len(session("pcSFCity"))<1 then
				session("pcSFCity")=getUserInput(request("sCity"),0)
		end if 
		if len(session("pcSFState"))<1 then
				session("pcSFState")=getUserInput(request("sState"),0)
		end if 
		if len(session("pcSFWeight"))<1 then
			session("pcSFWeight")=getUserInput(request("sWeight"),0)
		end if
		if len(session("pcSFPackageNum"))<1 then
			session("pcSFPackageNum")=getUserInput(request("sPackageCnt"),0)
		end if
				
		If request("Submit")<>"" then	
			
			'// UPS CREDENTIALS
			query="SELECT active, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=3;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
		
			if NOT rs.eof then
				UPS_Active=rs("active")
				UPS_UserId=trim(rs("userID"))
				UPS_Password=trim(rs("password"))
				UPS_LicenseKey=trim(rs("AccessLicense"))
			end if
		
			set rs=nothing

			'// set errors to none
			pcv_intErr=0
					
			if request("Candidate")="Y" then
				k=request("intCandidateSelection")

				ArryPoliticalDivision1=request("ArryPoliticalDivision1")
				ArryPoliticalDivision2=request("ArryPoliticalDivision2")
				ArryPoliticalDivision3=request("ArryPoliticalDivision3")
				ArryCountry=request("ArryCountry")
				ArryCountryCode=request("ArryCountryCode")
				ArryPostcodePrimaryLow=request("ArryPostcodePrimaryLow")
				ArryPostcodePrimaryHigh=request("ArryPostcodePrimaryHigh")

				splitPoliticalDivision1=split(ArryPoliticalDivision1,",")
				splitPoliticalDivision2=split(ArryPoliticalDivision2,",")
				splitPoliticalDivision3=split(ArryPoliticalDivision3,",")
				splitCountry=split(ArryCountry,",")
				splitCountryCode=split(ArryCountryCode,",")
				splitPostcodePrimaryLow=split(ArryPostcodePrimaryLow,",")
				splitPostcodePrimaryHigh=split(ArryPostcodePrimaryHigh,",")
			
				session("pcSFCountryCode")=splitCountryCode(k)
				session("pcSFPostalCode")=splitPostcodePrimaryLow(k)
				session("pcSFCity")=splitPoliticalDivision2(k)
				session("pcSFState")=splitPoliticalDivision1(k)
				'session("pcSFResidential")=splitCountryCode(k)


			else
				'/////////////////////////////////////////////////////
				'// Validate Fields and Set Sessions	
				'/////////////////////////////////////////////////////
			
				'// generic error for page
				pcv_strGenericPageError = "One of more fields were not filled in correctly."
			
				pcv_isPostalCodeRequired=false
				pcv_isCityRequired=false
				pcv_isStateRequired=false
				pcv_isResidentialRequired=false

				'// Clear error string
				pcv_strErrorMsg = ""
				pcs_ValidateTextField	"CountryCode", pcv_isCountryCodeRequired, 2
				if session("pcSFCountryCode")="US" then
					'pcv_isPostalCodeRequired=true
				end if
				pcs_ValidateTextField	"PostalCode", pcv_isPostalCodeRequired, 10
				pcs_ValidateTextField	"City", pcv_isCityRequired, 30
				pcs_ValidateTextField	"State", pcv_isStateRequired, 30
				pcs_ValidateTextField	"Residential", pcv_isResidentialRequired, 1
				pcs_ValidateTextField	"Weight", true, 0
				pcs_ValidateTextField	"PackageNum", true, 0   
			end if
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			' Check for Validation Errors. Do not proceed if there are errors.
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			If pcv_intErr>0 Then
				session("ErrMsg") = pcv_strGenericPageError
				response.redirect pcPageName & "?sub=1"
			Else				

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Build Our Transaction.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				objUPSClass.NewXMLTransaction UPS_LicenseKey, UPS_UserId, UPS_Password

				objUPSClass.NewXMLShipmentTimeInTransitRequest "Customer Data"

				'/////////////////////////////////////////////////////
				'// Set Local Variables for Setting
				'/////////////////////////////////////////////////////

				pcStrCountryCode = Session("pcSFCountryCode")
				pcStrPostalCode = Session("pcSFPostalCode")
				pcStrCity = Session("pcSFCity")
				pcStrState = Session("pcSFState")
				pcStrResidential = Session("pcSFResidential")

				objUPSClass.WriteParent "TransitFrom", ""
				objUPSClass.WriteParent "AddressArtifactFormat", ""
					objUPSClass.AddNewNode "PoliticalDivision2", UPS_CITY
					objUPSClass.AddNewNode "PoliticalDivision1", UPS_STATE
					objUPSClass.AddNewNode "CountryCode", UPS_COUNTRY
					objUPSClass.AddNewNode "PostcodePrimaryLow", UPS_POSTALCODE
				objUPSClass.WriteParent "AddressArtifactFormat", "/"
				objUPSClass.WriteParent "TransitFrom", "/"
			
				objUPSClass.WriteParent "TransitTo", ""
				objUPSClass.WriteParent "AddressArtifactFormat", ""
					objUPSClass.AddNewNode "PoliticalDivision2", session("pcSFCity")
					objUPSClass.AddNewNode "PoliticalDivision1", session("pcSFState")
					objUPSClass.AddNewNode "CountryCode", session("pcSFCountryCode")
					objUPSClass.AddNewNode "PostcodePrimaryLow", session("pcSFPostalCode")
				objUPSClass.WriteParent "AddressArtifactFormat", "/"
				objUPSClass.WriteParent "TransitTo", "/"


				'New Code for Date Format - NO SATURDAY DELIVERIES
				UPSTodayDate=Now()+1
				strWeekdayName=WeekdayName(Weekday(UPSTodayDate))
				If instr(ucase(strWeekdayName),"SATURDAY") then
					UPSTodayDate=Now()+3
				end if
				If instr(ucase(strWeekdayName),"SUNDAY") then
					UPSTodayDate=Now()+2
				end if
				dtShipMonth=Month(UPSTodayDate)
				if int(dtShipMonth)<10 then
					dtShipMonth="0"&dtShipMonth
				end if
				dtShipDay=Day(UPSTodayDate)
				if int(dtShipDay)<10 then
					dtShipDay="0"&dtShipDay
				end if
				dtShipYear=Year(UPSTodayDate)
				dtShipDate=dtShipYear&dtShipMonth&dtShipDay


				objUPSClass.WriteParent "ShipmentWeight", ""
					objUPSClass.WriteParent "UnitOfMeasurement", ""
							objUPSClass.AddNewNode "Code", "LBS"
							objUPSClass.AddNewNode "Description", "Pounds"
					objUPSClass.WriteParent "UnitOfMeasurement", "/"
					objUPSClass.AddNewNode "Weight", session("pcSFWeight")/16
				objUPSClass.WriteParent "ShipmentWeight", "/"
				objUPSClass.AddNewNode "TotalPackagesInShipment", session("pcSFPackageNum")
				objUPSClass.WriteParent "InvoiceLineTotal", ""
						objUPSClass.AddNewNode "CurrencyCode", "USD"
						objUPSClass.AddNewNode "MonetaryValue", "250.00"
				objUPSClass.WriteParent "InvoiceLineTotal", "/"
				objUPSClass.AddNewNode "PickupDate", dtShipDate
						'objUPSClass.WriteParent "DocumentsOnlyIndicator", ""
						'objUPSClass.WriteParent "DocumentsOnlyIndicator", "/"

				objUPSClass.EndXMLTransaction "TimeInTransitRequest"

				'//Clear illegal ampersand characters from XML
				'UPS_postdata=replace(UPS_postdata, "&", "and")
				'UPS_postdata=replace(UPS_postdata, "&amp;", "and")
		
				UPS_data = UPS_AccessRequest & UPS_PostData

				'// Print out our newly formed request xml
				'response.clear
				'response.contenttype = "text/xml"
				'response.write UPS_PostData
				'response.end
							
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Send Our Transaction.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				UPS_URL="https://onlinetools.ups.com/ups.app/xml/TimeInTransit"
				call objUPSClass.SendXMLRequest(UPS_data, UPS_URL)
			
				'// Print out our response
				'response.write UPS_result
				'response.end
			
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Load Our Response.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				call objUPSClass.LoadXMLResults(UPS_result)
														
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Check for errors from UPS.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
		
				'//SOME ERROR CHECKING HERE							
				call objUPSClass.XMLResponseVerify(ErrPageName)
			
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Redirect with a Message OR complete some task.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				'if NOT len(pcv_strErrorMsg)>0 AND UCASE(pcv_strErrorSeverity)<>"WARNING" then
							

				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				' Set Our Response Data to Local.
				'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
				dim intCandidate
				intCandidate=0
			
				ArryCandidate = objUPSClass.ReadResponseasArray("//Candidate", "PoliticalDivision2")	
				if ArryCandidate<>"" then
					ArryPoliticalDivision1 = objUPSClass.ReadResponseasArray("//AddressArtifactFormat", "PoliticalDivision1")	
					ArryPoliticalDivision2 = objUPSClass.ReadResponseasArray("//AddressArtifactFormat", "PoliticalDivision2")	
					ArryPoliticalDivision3 = objUPSClass.ReadResponseasArray("//AddressArtifactFormat", "PoliticalDivision3")	
					ArryCountry = objUPSClass.ReadResponseasArray("//AddressArtifactFormat", "Country")	
					ArryCountryCode = objUPSClass.ReadResponseasArray("//AddressArtifactFormat", "CountryCode")	
					ArryPostcodePrimaryLow = objUPSClass.ReadResponseasArray("//AddressArtifactFormat", "PostcodePrimaryLow")
					ArryPostcodePrimaryHigh = objUPSClass.ReadResponseasArray("//AddressArtifactFormat", "PostcodePrimaryHigh")
				
					'// Create Flag to show drop down (or radio buttons)
					intCandidate=1
				else				
					ArryDescription = objUPSClass.ReadResponseasArray("//Service", "Description")
					ArryBusinessTransitDays = objUPSClass.ReadResponseasArray("//EstimatedArrival", "BusinessTransitDays")
					ArryTime = objUPSClass.ReadResponseasArray("//EstimatedArrival", "Time")
				end if
				%>
				<form method="POST" name="orderform" action="pcUPSTimeInTransit.asp" class="pcForms">   
					<h2>Time-In-Transit: Calculate estimated transit time</h2>
					<div class="pcPageDesc">Time-in-transit does not include the time to process order, the day it ships, weekends, and holidays.</div>

					<div class="pcShowContent">
							
						<% if intCandidate=1 then
							splitPoliticalDivision1=split(ArryPoliticalDivision1,",") %>
								<div class="pcErrorMessage">There was a problem encountered with the destination address.</div>

								<h3>Select Destination Address</h3>
										
								<div class="pcFormItem">
									<select name="intCandidateSelection">
									<% for j=0 to ubound(splitPoliticalDivision1)-1
										'response.write "<BR>"&ArryPoliticalDivision1&"<BR>"
										splitPoliticalDivision2=split(ArryPoliticalDivision2,",")
										splitPoliticalDivision3=split(ArryPoliticalDivision3,",")
										splitCountry=split(ArryCountry,",")
										splitCountryCode=split(ArryCountryCode,",")
										splitPostcodePrimaryLow=split(ArryPostcodePrimaryLow,",")
										splitPostcodePrimaryHigh=split(ArryPostcodePrimaryHigh,",")
										'//Create String
										strOption=""
										if len(splitPoliticalDivision2(j))>1 then
											strOption=strOption&splitPoliticalDivision2(j)&", "
										end if
										response.write ubound(splitPoliticalDivision3)
										if len(splitPoliticalDivision3(j))>1 then
											strOption=strOption&splitPoliticalDivision3(j)&", "
										end if
										if len(splitPoliticalDivision1(j))>1 then
											strOption=strOption&splitPoliticalDivision1(j)&" "
										end if
										if len(splitPostcodePrimaryLow(j))>1 then
											strOption=strOption&splitPostcodePrimaryLow(j)&" "
										end if
										if len(splitCountryCode(j))>1 then
											strOption=strOption&splitCountryCode(j)&" "
										end if
										%>
										<option value="<%=j%>"><%=strOption%></option>
									<% next %>
									</select>
									<input type="hidden" name="Weight" value="<%=session("pcSFWeight")%>">
									<input type="hidden" name="PackageNum" value="<%=session("pcSFPackageNum")%>">   
									<input type="hidden" name="ArryPoliticalDivision1" value="<%=ArryPoliticalDivision1%>">
									<input type="hidden" name="ArryPoliticalDivision2" value="<%=ArryPoliticalDivision2%>">
									<input type="hidden" name="ArryPoliticalDivision3" value="<%=ArryPoliticalDivision3%>">
									<input type="hidden" name="ArryCountry" value="<%=ArryCountry%>">
									<input type="hidden" name="ArryCountryCode" value="<%=ArryCountryCode%>">
									<input type="hidden" name="ArryPostcodePrimaryLow" value="<%=ArryPostcodePrimaryLow%>">
									<input type="hidden" name="ArryPostcodePrimaryHigh" value="<%=ArryPostcodePrimaryHigh%>">
									<input type="hidden" name="Candidate" value="Y">&nbsp;
								</div>

								<div class="pcFormButtons">
									<button class="pcButtonSubmit" name="Submit" value="Submit">Submit</button>
								</div>
						<% else %>

							<%
								col_MethodClass = "pcCol-6"
								col_DaysClass		= "pcCol-3"
								col_TimeClass		= "pcCol-3"	
							%>
							<div class="pcTable">
								<div class="pcTableHeader">
									<div class="<%= col_MethodClass %>">Method</div>
									<div class="<%= col_DaysClass %>">Business Days</div>
									<div class="<%= col_TimeClass %>">Time</div>
								</div>
								<% SplitDescription=split(ArryDescription ,"," )
								SplitBusinessTransitDays=split(ArryBusinessTransitDays ,"," )
								SplitTime=split(ArryTime ,"," )
								for i=0 to ubound(SplitDescription) %>
								<div class="pcTableRow">
									<div class="<%= col_MethodClass %>"><%=SplitDescription(i)%></div>
									<div class="<%= col_DaysClass %>"><%=SplitBusinessTransitDays(i)%></div>
									<div class="<%= col_TimeClass %>"><%=SplitTime(i)%></div>
								</div>
								<%next %>
							</div>
						<% end if %>
					</div>
					<div class="pcSpacer"></div>
					<div class="pcFormButtons">
						<a class="pcButton pcButtonBack" href="javascript:window.history.back();">
							<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
							<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
						</a>
					</div>
					<div class="pcClear"></div>
				</form>	
			<% end if %>
		<% else %>
			<h2>Time-In-Transit: Calculate estimated transit time</h2>
			<div class="pcPageDesc">
				Use this form to determine the scheduled time-in-transit for packages shipped by UPS. Simply enter the destination of your shipment below, we will let you know the approximate delivery time from the time the shipment leaves the warehouse. 
				<br><br>
				<strong>Please note:</strong> Time-in-transit does not include the time to process order, the day it ships, weekends, and holidays.
			</div>

			<% msg=session("ErrMsg")
			session("ErrMsg")=""
			if msg&""<>"" then %>
				<div class="pcErrorMessage"><%=msg%></div>
			<% end if %>

			<form method="POST" name="orderform" action="pcUPSTimeInTransit.asp" class="pcForms">
				<input type="hidden" name="Weight" value="<%=session("pcSFWeight")%>">
				<input type="hidden" name="PackageNum" value="<%=session("pcSFPackageNum")%>">
				<div class="pcShowContent">
					<h3>Shipping Address</h3>

					<div class="pcFormItem">
						<div class="pcFormLabel">City:</div>
						<div class="pcFormField"><input type="text" name="City" value="<%=session("pcSFCity")%>"></div>
					</div>

					<div class="pcFormItem">
						<div class="pcFormLabel">State:</div>
						<div class="pcFormField"><input type="text" name="State" value="<%=session("pcSFState")%>"></div>
					</div>

					<div class="pcFormItem">
						<div class="pcFormLabel">Postal Code:</div>
						<div class="pcFormField"><input type="text" name="PostalCode" value="<%=session("pcSFPostalCode")%>"></div>
					</div>

					<div class="pcFormItem">
						<div class="pcFormLabel">Country:</div>
						<div class="pcFormField"><input type="text" name="CountryCode" value="<%=session("pcSFCountryCode")%>"></div>
					</div>

					<div class="pcFormItem">
						<div class="pcFormLabel">&nbsp;</div>
						<div class="pcFormField"><input type="checkbox" name="checkbox" value="checkbox" class="clearBorder">&nbsp;This is a Residential Address</div>
					</div>

					<div class="pcFormButtons">
						<button class="pcButtonSubmit" name="Submit" value="Submit">Submit</button>
					</div>
				</div>
			</form>
			<% end if %>
	
			<div class="pcFormItem">
				<div class="pcFormItemFull"><hr /></div>
			</div>
			<div id="pcUPSFooter">
				<div class="pcUPSLogo"><img src="<%=pcf_getImagePath("../UPSLicense","LOGO_S2.png")%>" alt="UPS Logo" style="width: 42px; height: 50px;" /></div>
				<div class="pcUPSTerms">
					<span class="pcSmallText">UPS, the UPS Shield trademark, the UPS Ready mark, <br />the UPS Developer Kit mark and the Color Brown are trademarks of <br />United Parcel Service of America, Inc. All Rights Reserved.</span>
				</div>
			</div>
		</div>
	</div>
</body>
<%conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing%>
