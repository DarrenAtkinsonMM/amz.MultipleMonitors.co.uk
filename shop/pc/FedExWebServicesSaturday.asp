<%
pcPageName = "ShipRates.asp"

pcv_strMethodNameWS = "RateRequest"
pcv_strMethodReplyWS = "RateResponse"
pcv_customerTransactionID = "ProductCart Rates"
pcv_strEnvironment = FEDEXWS_Environment
pcv_strFedExBaselineLogging = false

'// FEDEX CREDENTIALS
query = "SELECT ShipmentTypes.userID, ShipmentTypes.password, ShipmentTypes.AccessLicense, ShipmentTypes.FedExKey, ShipmentTypes.FedExPwd "
query = query & "FROM ShipmentTypes "
query = query & "WHERE (((ShipmentTypes.idShipment)=9));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if NOT rs.eof then
	FedExAccountNumber=rs("userID")
	FedExMeterNumber=rs("password")
	pcv_strEnvironment=rs("AccessLicense")
	FedExkey=rs("FedExKey")
	FedExPassword=rs("FedExPwd")
end if
set rs=nothing

if (FedEXWS_active=true or FedExWS_active="-1") AND FedEXWS_AccountNumber<>"" then

	iFedExWSActive=1
	arryFedExWSService=""
	arryFedExWSRate=""
	arrFedExWSDeliveryDate=""
	arryFedExWSRate2 = ""

	pcv_TmpListRate = FEDEXWS_LISTRATE
	pcv_TmpSaturdayDelivery = FEDEXWS_SATURDAYDELIVERY
	'// Override List Rates for International addresses
	'If Universal_destination_country<>"US" Then
	'	pcv_TmpListRate = "0"
	'End If

	pcv_strVersion = FedExWS_RateVersion

	'// FedEx EXPRESS RATES
	set objFedExWSClass = New pcFedExWSClass

	fedex_postdataWS=""
	FEDEXWS_result=""

	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// START: FedEx Express
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	
	pcv_strCarrierCode = "FDXE"
  pcv_strShippingMethod = ""

	objFedExWSClass.NewXMLSubscription pcv_strMethodNameWS, FedExKey, FedExPassword, FedExWS_RateVersion, "rate"

		objFedExWSClass.WriteParent "ClientDetail", ""
			objFedExWSClass.AddNewNode "AccountNumber", FedExAccountNumber
			objFedExWSClass.AddNewNode "MeterNumber", FedExMeterNumber
			objFedExWSClass.AddNewNode "ClientProductId", pcv_strClientProductID
			objFedExWSClass.AddNewNode "ClientProductVersion", pcv_strClientProductVersion
		objFedExWSClass.WriteParent "ClientDetail", "/"

		'// Transaction ID
		objFedExWSClass.WriteParent "TransactionDetail", ""
			objFedExWSClass.AddNewNode "CustomerTransactionId", pcv_customerTransactionID
		objFedExWSClass.WriteParent "TransactionDetail", "/"

		objFedExWSClass.WriteParent "Version", ""
			objFedExWSClass.AddNewNode "ServiceId", "crs"
			objFedExWSClass.AddNewNode "Major", FedExWS_RateVersion
			objFedExWSClass.AddNewNode "Intermediate", "0"
			objFedExWSClass.AddNewNode "Minor", "0"
		objFedExWSClass.WriteParent "Version", "/"
  
		objFedExWSClass.AddNewNode "ReturnTransitAndCommit", "1"
    If pcv_strShippingMethod = "" Then
		  objFedExWSClass.WriteSingleParent "CarrierCodes", pcv_strCarrierCode
    End If
	
		If FEDEXWS_ONERATE = "-1" Then
			objFedExWSClass.AddNewNode "VariableOptions", "FEDEX_ONE_RATE"
		End If

		objFedExWSClass.WriteParent "RequestedShipment", ""

			'// Ship Date
			pcAddDay = 	FEDEXWS_ADDDAY
			if pcAddDay&""="" then
				pcAddDay = 0
			end if

			pcFedExShipDate = DateAdd("D", pcAddDay, Date)
			pcWeekDay = WeekDay(pcFedExShipDate)
			pcAddDay = 	FEDEXWS_ADDDAY
			if pcAddDay&""="" then
				pcAddDay = 0
			end if

			if pcWeekDay = "1" then
				'Add a day to the current date
				pcFedExShipDate = DateAdd("D", 1, Date)
			end if
			'if saturday pickup=0 then we need to shift a saturday date to monday
			if pcWeekDay = "7" AND FEDEXWS_SATURDAYPICKUP = "0" then
				'Add 2 days to the current date
				pcFedExShipDate = DateAdd("D", 2, Date)
			end if
			FwsYear = Year(pcFedExShipDate)
			FwsMonth = Month(pcFedExShipDate)
			if len(FwsMonth)=1 Then
				FwsMonth = "0"&FwsMonth
			end if
			FwsDay = Day(pcFedExShipDate)
			if len(FwsDay)=1 Then
				FwsDay = "0"&FwsDay
			end if

			pcFedExFormatedShipDate = FwsYear&"-"&FwsMonth&"-"&FwsDay&"T14:33:57+05:30"

			objFedExWSClass.AddNewNode "ShipTimestamp", pcFedExFormatedShipDate

			objFedExWSClass.AddNewNode "DropoffType", FEDEXWS_DROPOFF_TYPE
		  objFedExWSClass.AddNewNode "ServiceType", pcv_strShippingMethod
			objFedExWSClass.AddNewNode "PackagingType", FEDEXWS_FEDEX_PACKAGE

      '// Add total weight
      totalPounds = 0
      totalOunces = 0
			for q=1 to pcv_intPackageNum
				tmpPounds = session("FEDEXWSPackPounds"&q)
				tmpOunces = session("FEDEXWSPackOunces"&q)
				if NOT isNumeric(tmpPounds) then
					tmpPounds = 0
				end if
				if NOT isNumeric(tmpOunces) then
					tmpOunces = 0
				end if
				tmpOuncesDec = Cint(tmpOunces)/16

        totalPounds = totalPounds + tmpPounds
        totalOunces = totalOunces + tmpOuncesDec
      next

			objFedExWSClass.WriteParent "TotalWeight", ""
				If scShipFromWeightUnit="LBS" Then
					objFedExWSClass.WriteSingleParent "Units", "LB"
				Else
					objFedExWSClass.WriteSingleParent "Units", "KG"
				End If
				objFedExWSClass.WriteSingleParent "Value", tmpPounds + tmpOuncesDec
			objFedExWSClass.WriteParent "TotalWeight", "/"

			objFedExWSClass.WriteParent "Shipper", ""
				objFedExWSClass.WriteParent "Contact", ""
					objFedExWSClass.AddNewNode "PersonName", scOriginPersonName
					objFedExWSClass.AddNewNode "CompanyName", scShipFromName
					objFedExWSClass.AddNewNode "PhoneNumber", scOriginPhoneNumber
				objFedExWSClass.WriteParent "Contact", "/"

				objFedExWSClass.WriteParent "Address", ""
					objFedExWSClass.AddNewNode "StreetLines", scShipFromAddress1
					objFedExWSClass.AddNewNode "City", scShipFromCity
                    If FedExRequiresStateProvince(scShipFromPostalCountry) Then
						objFedExWSClass.AddNewNode "StateOrProvinceCode", FedExCorrectStateProvince(scShipFromState)
					End If
					objFedExWSClass.AddNewNode "PostalCode", scShipFromPostalCode
					objFedExWSClass.AddNewNode "CountryCode", scShipFromPostalCountry
				objFedExWSClass.WriteParent "Address", "/"


			objFedExWSClass.WriteParent "Shipper", "/"

			objFedExWSClass.WriteParent "Recipient", ""
				objFedExWSClass.WriteParent "Contact", ""
					objFedExWSClass.AddNewNode "PersonName", pcShippingNickName
					objFedExWSClass.AddNewNode "CompanyName", pcShippingCompany
          objFedExWSClass.AddNewNode "PhoneNumber", pcShippingPhone
          objFedExWSClass.AddNewNode "EMailAddress", pcShippingEmail
				objFedExWSClass.WriteParent "Contact", "/"

				objFedExWSClass.WriteParent "Address", ""
					objFedExWSClass.AddNewNode "StreetLines", Universal_destination_address
					objFedExWSClass.AddNewNode "City", Universal_destination_city
                    If FedExRequiresStateProvince(Universal_destination_country) Then
						objFedExWSClass.AddNewNode "StateOrProvinceCode", FedExCorrectStateProvince(Universal_destination_provOrState)
					end if
					objFedExWSClass.AddNewNode "PostalCode", Universal_destination_postal
					objFedExWSClass.AddNewNode "CountryCode", Universal_destination_country
					if pResidentialShipping="-1" or pResidentialShipping="1" then
						objFedExWSClass.AddNewNode "Residential", "true"
          else
            objFedExWSClass.AddNewNode "Residential", "false"
					end if
				objFedExWSClass.WriteParent "Address", "/"

			objFedExWSClass.WriteParent "Recipient", "/"

			objFedExWSClass.WriteParent "ShippingChargesPayment", ""
				objFedExWSClass.AddNewNode "PaymentType", "SENDER"
        
			  objFedExWSClass.WriteParent "Payor", ""
			    objFedExWSClass.WriteParent "ResponsibleParty", ""
				    objFedExWSClass.AddNewNode "AccountNumber", FedExAccountNumber
			    objFedExWSClass.WriteParent "ResponsibleParty", "/"
			  objFedExWSClass.WriteParent "Payor", "/"
			objFedExWSClass.WriteParent "ShippingChargesPayment", "/"


			If pcv_TmpSaturdayDelivery<>"0" Then
				objFedExWSClass.WriteParent "SpecialServicesRequested", ""
					objFedExWSClass.AddNewNode "SpecialServiceTypes", "FUTURE_DAY_SHIPMENT"

					'// Saturday Delivery
					objFedExWSClass.AddNewNode "SpecialServiceTypes", "SATURDAY_DELIVERY"

				objFedExWSClass.WriteParent "SpecialServicesRequested", "/"
			End If
	
	    If pcv_TmpListRate = "-1" Or pcv_TmpListRate = -1 Then
		    pcv_strRateRequestType = "LIST"
	    elseIf pcv_TmpListRate = "-2" Or pcv_TmpListRate = -2 Then
		    pcv_strRateRequestType = "PREFERRED"
	    Else
		    pcv_strRateRequestType = "NONE"
	    End If
			objFedExWSClass.AddNewNode "RateRequestTypes", pcv_strRateRequestType
			objFedExWSClass.AddNewNode "PackageCount", pcv_intPackageNum


			for q=1 to pcv_intPackageNum
				'//FEDEXWSWEIGHTCHANGE///////////////////////////////////////
				tmpPounds = session("FEDEXWSPackPounds"&q)
				tmpOunces = session("FEDEXWSPackOunces"&q)
				if NOT isNumeric(tmpPounds) then
					tmpPounds = 0
				end if
				if NOT isNumeric(tmpOunces) then
					tmpOunces = 0
				end if
				tmpOuncesDec = Cint(tmpOunces)/16

				objFedExWSClass.WriteParent "RequestedPackageLineItems", ""

					objFedExWSClass.AddNewNode "SequenceNumber", q
					objFedExWSClass.AddNewNode "GroupNumber", q
					objFedExWSClass.AddNewNode "GroupPackageCount", q

					pcv_TempPackPrice=session("FedEXWSPackPrice"&q)
					If pcv_TempPackPrice="" Then
						pcv_TempPackPrice="100.00"
					End If

					objFedExWSClass.WriteParent "InsuredValue", ""
						objFedExWSClass.AddNewNode "Currency", FEDEXWS_CURRENCY
						objFedExWSClass.AddNewNode "Amount", replace(money(pcv_TempPackPrice),",","")
					objFedExWSClass.WriteParent "InsuredValue", "/"

					objFedExWSClass.WriteParent "Weight", ""
						If scShipFromWeightUnit="LBS" Then
							objFedExWSClass.WriteSingleParent "Units", "LB"
						Else
							objFedExWSClass.WriteSingleParent "Units", "KG"
						End If
						objFedExWSClass.WriteSingleParent "Value", tmpPounds + tmpOuncesDec
					objFedExWSClass.WriteParent "Weight", "/"

					if ((FEDEXWS_FEDEX_PACKAGE="YOUR_PACKAGING") AND (session("FedEXWSPackLength"&q)<>"" AND session("FedEXWSPackWidth"&q)<>"" AND session("FedEXWSPackHeight"&q)<>"")) then
						pcv_strDimUnit = FEDEXWS_DIM_UNIT
						if pcv_strDimUnit="" then
							pcv_strDimUnit = "IN"
						end if
						objFedExWSClass.WriteParent "Dimensions", ""
							objFedExWSClass.AddNewNode "Length", Int(session("FedEXWSPackLength"&q))
							objFedExWSClass.AddNewNode "Width", Int(session("FedEXWSPackWidth"&q))
							objFedExWSClass.AddNewNode "Height", Int(session("FedEXWSPackHeight"&q))
							objFedExWSClass.AddNewNode "Units", FEDEXWS_DIM_UNIT
						objFedExWSClass.WriteParent "Dimensions", "/"
					end if

          ' FEDEX CERT ONLY
				  objFedExWSClass.WriteParent "CustomerReferences", ""
					  objFedExWSClass.AddNewNode "CustomerReferenceType", "CUSTOMER_REFERENCE"
					  objFedExWSClass.AddNewNode "Value", "CUSTOMER_REFERENCE"
				  objFedExWSClass.WriteParent "CustomerReferences", "/"

				  objFedExWSClass.WriteParent "CustomerReferences", ""
					  objFedExWSClass.AddNewNode "CustomerReferenceType", "INVOICE_NUMBER"
					  objFedExWSClass.AddNewNode "Value", pcv_customerTransactionID
				  objFedExWSClass.WriteParent "CustomerReferences", "/"

				objFedExWSClass.WriteParent "RequestedPackageLineItems", "/"

			next

		objFedExWSClass.WriteParent "RequestedShipment", "/"

	objFedExWSClass.EndXMLTransaction pcv_strMethodNameWS

'--------------------------------------------------------------------------------------------------------


	'// Print out our newly formed request xml
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(fedex_postdataWS)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Send Our Transaction.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.SendXMLRequest(fedex_postdataWS)

	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(FEDEXWS_result)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Load Our Response.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.LoadXMLResults(FEDEXWS_result)


	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for errors from FedEx.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//Notifications", "Severity")
	if pcv_strErrorMsgWS<>"SUCCESS" AND  pcv_strErrorMsgWS<>"NOTE" AND pcv_strErrorMsgWS<>"WARNING"  then
		pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//Notifications", "Message")
	else
		pcv_strErrorMsgWS = ""
	end if
	
	'/////////////////////////////////////////////////////////////
	'// BASELINE LOGGING
	'/////////////////////////////////////////////////////////////
	'// Log our Request
	'call objFedExWSClass.pcs_LogTransaction(fedex_postdataWS, "Rate_" & pcv_strCarrierCode & "_" & pcv_strLogID & "_Sat_Req.xml", pcv_strFedExBaselineLogging)
	'// Log our Response
	'call objFedExWSClass.pcs_LogTransaction(FEDEXWS_result, "Rate_" & pcv_strCarrierCode & "_" & pcv_strLogID & "_Sat_Res.xml", pcv_strFedExBaselineLogging)
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Redirect with a Message OR complete some task.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	arrFedExWSDeliveryDate = ""
	arrFedExWSDeliveryTime = ""
	If NOT len(pcv_strErrorMsgWS) > 0 Then
		arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "ServiceType")

		if scHideEstimateDeliveryTimes="1" then
			arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
			arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
		else
      deliveryDates = objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "DeliveryDayOfWeek")
      deliveryTimes = objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "DeliveryTimestamp")

			arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & deliveryDates
			arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & deliveryTimes
		end if

    '// Use net charge (no tax) for CA rates
    If scShipFromPostalCountry="CA" Then
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/TotalNetFedExCharge/Amount")
    Else
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/TotalNetCharge/Amount")
    End If
	End If
	set objFedExWSClass = nothing
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// END: FedEx Express
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////

	'Split Arrays
	dim intRateIndexWSS
	dim pcFedExWSSMultiArry(50,5)
	for z=0 to 50
		pcFedExWSSMultiArry(z,1)=0
	next

	pcStrTempFedExService=split(arryFedExWSService,",")
	pcStrTempFexExRate=split(arryFedExWSRate,",")
	pcStrTempFedExDeliveryDate=split(arrFedExWSDeliveryDate,",")
	pcStrTempFedExDeliveryTime=split(arrFedExWSDeliveryTime,",")

	for t=0 to (ubound(pcStrTempFedExService)-1)
		For idx = 0 To UBound(FedExWS_ShipmentTypes) - 1
			Service = FedExWS_ShipmentTypes(idx)
			If pcStrTempFedExService(t) = Service Then
				If pcStrTempFedExService(t) = "GROUND_HOME_DELIVERY" Then
					If Universal_destination_country="US" Then
						intRateIndexWSS = idx
						pcFedExWSSMultiArry(intRateIndexWSS,2)=FedExWS_ShipmentName(pcStrTempFedExService(t)) & " (Saturday Delivery)"
						pcFedExWSSMultiArry(intRateIndexWSS,3)=pcStrTempFedExService(t)
						pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
						pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
					End If
				Else
					intRateIndexWSS = idx
					pcFedExWSSMultiArry(intRateIndexWSS,2)=FedExWS_ShipmentName(pcStrTempFedExService(t)) & " (Saturday Delivery)"
					pcFedExWSSMultiArry(intRateIndexWSS,3)=pcStrTempFedExService(t)
					pcFedExWSSMultiArry(intRateIndexWSS,4)=pcStrTempFedExDeliveryDate(t)
					pcFedExWSSMultiArry(intRateIndexWSS,5)=pcStrTempFedExDeliveryTime(t)
				End If
			End if
		Next

		tempRate=pcFedExWSSMultiArry(intRateIndexWSS,1)
		pcFedExWSSMultiArry(intRateIndexWSS,1)=cdbl(tempRate)+cdbl(pcStrTempFexExRate(t))
	next

	for z=0 to 50
		if pcFedExWSSMultiArry(z,1)>0 then
			pcv_strFormattedDate = ""

			tmpDeliveryDayOfWeek = pcFedExWSSMultiArry(z,4)
			select case tmpDeliveryDayOfWeek
				case "MON"
					tmpDeliveryDayOfWeek = "Monday"
				case "TUE"
					tmpDeliveryDayOfWeek = "Tuesday"
				case "WED"
					tmpDeliveryDayOfWeek = "Wednesday"
				case "THU"
					tmpDeliveryDayOfWeek = "Thursday"
				case "FRI"
					tmpDeliveryDayOfWeek = "Friday"
				case "SAT"
					tmpDeliveryDayOfWeek = "Saturday"
			end select

			tmpDeliveryTimestamp = pcFedExWSSMultiArry(z,5)
			arrDeliveryTimestamp = split(tmpDeliveryTimestamp, "T")
			tmpDeliveryTime = arrDeliveryTimestamp(1)
			arrTimeFormat = split(tmpDeliveryTime,":")
			tmpTimeHour = Cint(arrTimeFormat(0))
			tmpTimeMinutes = arrTimeFormat(1)
			tmpTimeSeconds = arrTimeFormat(2)
			'//Format hour and check for AM/PM
			if tmpTimeHour < 12 then
				tmpAMPM = "AM"
				tmpHour = Cint(tmpTimeHour)
			else
				tmpAMPM = "PM"
				tmpHour = Cint(tmpTimeHour) - 12
			end if
			tmpDeliveryDate = arrDeliveryTimestamp(0)
			arrDeliveryDate = split(tmpDeliveryDate,"-")
			tmpDeliveryDay = arrDeliveryDate(2)
			tmpDeliveryMonth = arrDeliveryDate(1)
			select case tmpDeliveryMonth
				case "01"
					tmpDeliveryMonth = "January"
				case "02"
					tmpDeliveryMonth = "February"
				case "03"
					tmpDeliveryMonth = "March"
				case "04"
					tmpDeliveryMonth = "April"
				case "05"
					tmpDeliveryMonth = "May"
				case "06"
					tmpDeliveryMonth = "June"
				case "07"
					tmpDeliveryMonth = "July"
				case "08"
					tmpDeliveryMonth = "August"
				case "09"
					tmpDeliveryMonth = "September"
				case "10"
					tmpDeliveryMonth = "October"
				case "11"
					tmpDeliveryMonth = "November"
				case "12"
					tmpDeliveryMonth = "December"
			end select
			tmpDeliveryYear = arrDeliveryDate(0)

			pcv_strFormattedDate = tmpDeliveryDayOfWeek&", "&tmpDeliveryMonth&" "&tmpDeliveryDay&", "&tmpDeliveryYear&" "&tmpTimeHour&":"&tmpTimeMinutes&" "&tmpAMPM

			availableShipStr=availableShipStr&"|?|FedExWS|"&pcFedExWSSMultiArry(z,3)&"|"&pcFedExWSSMultiArry(z,2)&"|"&pcFedExWSSMultiArry(z,1)&"|"&pcv_strFormattedDate
			iFedExWSFlag=1
	
		end if
	next
end if 'if fedex is active
%>