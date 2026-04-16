<%
pcPageName = "ShipRates.asp"

pcv_strMethodNameWS = "RateRequest"
pcv_strMethodReplyWS = "RateResponse"
pcv_customerTransactionID = "ProductCart Rates"
pcv_strEnvironment = FEDEXWS_Environment
pcv_strFedExBaselineLogging = true
If (FedEXWS_active=true Or FedExWS_active="-1") Then

  '// FEDEX CREDENTIALS
  query = "SELECT ShipmentTypes.userID, ShipmentTypes.password, ShipmentTypes.AccessLicense, ShipmentTypes.FedExKey, ShipmentTypes.FedExPwd "
  query = query & "FROM ShipmentTypes "
  query = query & "WHERE (((ShipmentTypes.idShipment)=9));"
  Set rs = server.CreateObject("ADODB.RecordSet")
  Set rs = conntemp.execute(query)
  If Not rs.Eof Then
      FedExAccountNumber=rs("userID")
      FedExMeterNumber=rs("password")
      pcv_strEnvironment=rs("AccessLicense")
      FedExkey=rs("FedExKey")
      FedExPassword=rs("FedExPwd")
  End If
  Set rs = Nothing

  If FedEXWS_AccountNumber<>"" Then

	iFedExWSActive=1
	dim arryFedExWSService
	dim arryFedExWSRate
	dim arrFedExWSDeliveryDate
	arryFedExWSService=""
	arryFedExWSRate=""
	arrFedExWSDeliveryDate=""
	arryFedExWSRate2 = ""
	
	arryFedExWSRateType=""
	arryFedExWSSpecialRatingType=""

	pcv_TmpListRate = FEDEXWS_LISTRATE
	pcv_TmpSaturdayDelivery = 0

	'// Get packaging type based on oversized
	pcv_strPackaging = FEDEXWS_FEDEX_PACKAGE
	'for q=1 to pcv_intPackageNum
		'if session("OSFlaggedPackage"&q) = "YES" then
			'pcv_strPackaging = "YOUR_PACKAGING"
		'end if
	'next

	If pcv_TmpListRate = "-1" Or pcv_TmpListRate = -1 Then
		pcv_strRateRequestType = "LIST"
	elseIf pcv_TmpListRate = "-2" Or pcv_TmpListRate = -2 Then
		pcv_strRateRequestType = "PREFERRED"
	Else
		pcv_strRateRequestType = "NONE"
	End If

	'// FedEx EXPRESS RATES
	set objFedExWSClass = New pcFedExWSClass

	fedex_postdataWS=""
	FEDEXWS_result=""

	pcv_strLogID = objFedExWSClass.RandomNumber(999999999)

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

    '// Only include general carrier code if specific shipping method is not specified
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
			'//Identifies the date and time the package is tendered to FedEx. Both the date and time portions of the string are expected to be used. The date should not be a past date or a date more than 10 days in the future. The time is the local time of the shipment based on the shipper's time zone. The date component must be in the format: YYYY-MM-DD (e.g. 2006-06-26). The time component must be in the format: HH:MM:SS using a 24 hour clock (e.g. 11:00 a.m. is 11:00:00, whereas 5:00 p.m. is 17:00:00). The date and time parts are separated by the letter T (e.g. 2006-06-26T17:00:00). There is also a UTC offset component indicating the number of hours/mainutes from UTC (e.g 2006-06-26T17:00:00-0400 is defined form June 26, 2006 5:00 pm Eastern Time).</xs:documentation>

			objFedExWSClass.AddNewNode "DropoffType", FEDEXWS_DROPOFF_TYPE
			objFedExWSClass.AddNewNode "ServiceType", pcv_strShippingMethod
			objFedExWSClass.AddNewNode "PackagingType", pcv_strPackaging

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
				objFedExWSClass.WriteSingleParent "Value", totalPounds + totalOunces
			objFedExWSClass.WriteParent "TotalWeight", "/"

			If pcv_strRateRequestType = "PREFERRED" Then
				objFedExWSClass.AddNewNode "PreferredCurrency", FEDEXWS_PREFERRED_CURRENCY
			End If

			objFedExWSClass.WriteParent "Shipper", ""
				objFedExWSClass.WriteParent "Contact", ""
					objFedExWSClass.AddNewNode "PersonName", scOriginPersonName
					objFedExWSClass.AddNewNode "CompanyName", scShipFromName
					objFedExWSClass.AddNewNode "PhoneNumber", scOriginPhoneNumber
					objFedExWSClass.AddNewNode "EMailAddress", scOriginEmailAddress
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
					End If
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
					'// Saturday Delivery
					objFedExWSClass.AddNewNode "SpecialServiceTypes", "SATURDAY_DELIVERY"
					'// Saturday Pickup
					'objFedExWSClass.AddNewNode "SpecialServiceTypes", "SATURDAY_PICKUP"
				objFedExWSClass.WriteParent "SpecialServicesRequested", "/"
			End If

			objFedExWSClass.AddNewNode "RateRequestTypes", pcv_strRateRequestType
			objFedExWSClass.AddNewNode "PackageCount", pcv_intPackageNum

			for q=1 to pcv_intPackageNum
				tmpPounds = session("FEDEXWSPackPounds"&q)
				tmpOunces = session("FEDEXWSPackOunces"&q)
				
				if NOT isNumeric(tmpPounds) then
					tmpPounds = 0
				end if
				if NOT isNumeric(tmpOunces) then
					tmpOunces = 0
				end if

				If scShipFromWeightUnit="LBS" Then
		      tmpOuncesDec = CDbl(tmpOunces)/16
        Else
          tmpOuncesDec = tmpOunces
        End If

				objFedExWSClass.WriteParent "RequestedPackageLineItems", ""

					objFedExWSClass.AddNewNode "SequenceNumber", q
					objFedExWSClass.AddNewNode "GroupNumber", q
					objFedExWSClass.AddNewNode "GroupPackageCount", 1

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

					if ((pcv_strPackaging="YOUR_PACKAGING") AND (session("FedEXWSPackLength"&q)<>"" AND session("FedEXWSPackWidth"&q)<>"" AND session("FedEXWSPackHeight"&q)<>"")) then
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
	'call objFedExWSClass.pcs_LogTransaction(fedex_postdataWS, "Rate_" & pcv_strCarrierCode & "_" & pcv_strLogID & "_Req.xml", pcv_strFedExBaselineLogging)
	'// Log our Response
	'call objFedExWSClass.pcs_LogTransaction(FEDEXWS_result, "Rate_" & pcv_strCarrierCode & "_" & pcv_strLogID & "_Res.xml", pcv_strFedExBaselineLogging)

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Redirect with a Message OR complete some task.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
		end if
	End If

	set objFedExWSClass = nothing
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// END: FedEx Express
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////


	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// START: FedEx Smart Post
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	set objFedExWSClass = New pcFedExWSClass

	fedex_postdataWS=""
	FEDEXWS_result=""

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
		objFedExWSClass.WriteParent "RequestedShipment", ""

			'// Ship Date
			pcAddDay = 	FEDEXWS_ADDDAY
			if pcAddDay&""="" then
				pcAddDay = 0
			end if

			pcFedExShipDate = DateAdd("D", pcAddDay, Date)
			pcWeekDay = WeekDay(pcFedExShipDate)
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
			objFedExWSClass.AddNewNode "ServiceType", "SMART_POST"
			objFedExWSClass.AddNewNode "PackagingType", pcv_strPackaging

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
				objFedExWSClass.WriteSingleParent "Value", totalPounds + totalOunces
			objFedExWSClass.WriteParent "TotalWeight", "/"
	
			If pcv_strRateRequestType = "PREFERRED" Then
				objFedExWSClass.AddNewNode "PreferredCurrency", FEDEXWS_PREFERRED_CURRENCY
			End If

			objFedExWSClass.WriteParent "Shipper", ""
				objFedExWSClass.WriteParent "Contact", ""
					objFedExWSClass.AddNewNode "PersonName", scOriginPersonName
					objFedExWSClass.AddNewNode "CompanyName", scShipFromName
					objFedExWSClass.AddNewNode "PhoneNumber", scOriginPhoneNumber
					objFedExWSClass.AddNewNode "EMailAddress", scOriginEmailAddress
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
					if FedExRequiresStateProvince(Universal_destination_country) then
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
			objFedExWSClass.WriteParent "SmartPostDetail", ""
				objFedExWSClass.AddNewNode "Indicia", "PARCEL_SELECT"
				objFedExWSClass.AddNewNode "HubId", FEDEXWS_SMHUBID
				objFedExWSClass.AddNewNode "CustomerManifestId", "String"
			objFedExWSClass.WriteParent "SmartPostDetail", "/"

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
				If scShipFromWeightUnit="LBS" Then
				    tmpOuncesDec = CDbl(tmpOunces)/16
                Else
                    tmpOuncesDec = tmpOunces
                End If
			
				objFedExWSClass.WriteParent "RequestedPackageLineItems", ""
				
				  objFedExWSClass.AddNewNode "SequenceNumber", q
				  objFedExWSClass.AddNewNode "GroupNumber", q
				  objFedExWSClass.AddNewNode "GroupPackageCount", pcv_intPackageNum
					
				  objFedExWSClass.WriteParent "Weight", ""
				  If scShipFromWeightUnit="LBS" Then
					  objFedExWSClass.WriteSingleParent "Units", "LB"
				  Else
					  objFedExWSClass.WriteSingleParent "Units", "KG"
				  End If
				  objFedExWSClass.WriteSingleParent "Value", tmpPounds + tmpOuncesDec
				  objFedExWSClass.WriteParent "Weight", "/"
				  if ((pcv_strPackaging="YOUR_PACKAGING") AND (session("FedEXWSPackLength"&q)<>"" AND session("FedEXWSPackWidth"&q)<>"" AND session("FedEXWSPackHeight"&q)<>"")) then
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

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Load Our Response.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Print out our newly formed request xml
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(FEDEXWS_result)
	'response.End()
	call objFedExWSClass.LoadXMLResults(FEDEXWS_result)
	
	'/////////////////////////////////////////////////////////////
	'// BASELINE LOGGING
	'/////////////////////////////////////////////////////////////
	'// Log our Request
	'call objFedExWSClass.pcs_LogTransaction(fedex_postdataWS, "Rate_SP_" & pcv_strLogID & "_Req.xml", pcv_strFedExBaselineLogging)
	'// Log our Response
	'call objFedExWSClass.pcs_LogTransaction(FEDEXWS_result, "Rate_SP_" & pcv_strLogID & "_Res.xml", pcv_strFedExBaselineLogging)

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for errors from FedEx.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'call objFedExWSClass.XMLResponseVerify(ErrPageName)
	pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//Notifications", "Severity")
	if pcv_strErrorMsgWS<>"SUCCESS" AND  pcv_strErrorMsgWS<>"NOTE" AND pcv_strErrorMsgWS<>"WARNING"  then
		pcv_strErrorMsgWS = objFedExWSClass.ReadResponseNode("//Notifications", "Message")
	else
		pcv_strErrorMsgWS = ""
	end if

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Redirect with a Message OR complete some task.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	If NOT len(pcv_strErrorMsgWS)>0 Then
		arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "ServiceType")
		If scHideEstimateDeliveryTimes="1" Then
			arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
			arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
		Else
            
				pcv_minimumDeliverDays = objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "CommitDetails/TransitTime") 'TWO_DAYS
                
				pcv_maximumDeliverDays = objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "CommitDetails/MaximumTransitTime") 'EIGHT_DAYS
                
				arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & "SPTransitTime:"&objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "TransitTime")
                
				arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & "SPTransitTime:"&objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "CommitDetails/MaximumTransitTime")     
		End If
		arryFedExWSRateType = arryFedExWSRateType & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/RateType")
		arryFedExWSSpecialRatingType = arryFedExWSSpecialRatingType & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/SpecialRatingApplied")
		arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/TotalNetFedExCharge/Amount")
	End If
	set objFedExWSClass = nothing
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// END: FedEx SmartPost
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////

	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// START: FedEx Ground
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	set objFedExWSClass = New pcFedExWSClass

	fedex_postdataWS=""
	FEDEXWS_result=""
	
	pcv_strCarrierCode = "FDXG"
  pcv_strShippingMethod = ""

	objFedExWSClass.NewXMLSubscription pcv_strMethodNameWS, FedExkey, FedExPassword, FedExWS_RateVersion, "rate"

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

    if pcv_strShippingMethod = "" then
	  	objFedExWSClass.WriteSingleParent "CarrierCodes", pcv_strCarrierCode
    end if

		objFedExWSClass.WriteParent "RequestedShipment", ""

			'// Ship Date
			pcAddDay = 	FEDEXWS_ADDDAY
			if pcAddDay&""="" then
				pcAddDay = 0
			end if

			pcFedExShipDate = DateAdd("D", pcAddDay, Date)
			pcWeekDay = WeekDay(pcFedExShipDate)
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
			objFedExWSClass.AddNewNode "PackagingType", pcv_strPackaging

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
				objFedExWSClass.WriteSingleParent "Value", totalPounds + totalOunces
			objFedExWSClass.WriteParent "TotalWeight", "/"

			If pcv_strRateRequestType = "PREFERRED" Then
				objFedExWSClass.AddNewNode "PreferredCurrency", FEDEXWS_PREFERRED_CURRENCY
			End If

			objFedExWSClass.WriteParent "Shipper", ""
				objFedExWSClass.WriteParent "Contact", ""
					objFedExWSClass.AddNewNode "PersonName", scOriginPersonName
					objFedExWSClass.AddNewNode "CompanyName", scShipFromName
					objFedExWSClass.AddNewNode "PhoneNumber", scOriginPhoneNumber
					objFedExWSClass.AddNewNode "EMailAddress", scOriginEmailAddress
				objFedExWSClass.WriteParent "Contact", "/"

				objFedExWSClass.WriteParent "Address", ""
					objFedExWSClass.AddNewNode "StreetLines", scShipFromAddress1
					objFedExWSClass.AddNewNode "City", scShipFromCity
					if FedExRequiresStateProvince(scShipFromPostalCountry) then
						objFedExWSClass.AddNewNode "StateOrProvinceCode", FedExCorrectStateProvince(scShipFromState)
					end if
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
					if FedExRequiresStateProvince(Universal_destination_country) then
						objFedExWSClass.AddNewNode "StateOrProvinceCode",  FedExCorrectStateProvince(Universal_destination_provOrState)
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

			'// ShipDate
			'objFedExWSClass.AddNewNode "ShipDate", "2011-08-03T14:33:57+05:30"
			'objFedExWSClass.WriteParent "ShipDate", ""
			'	objFedExWSClass.AddNewNode "ShipDate", "2011-08-03"
			'objFedExWSClass.WriteParent "ShipDate", "/"

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start: FDXE Special Services
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			If pcv_TmpSaturdayDelivery<>"0" Then
				objFedExWSClass.WriteParent "SpecialServicesRequested", ""
					'// Saturday Delivery
					objFedExWSClass.AddNewNode "SpecialServiceTypes", "SATURDAY_DELIVERY"
					'// Saturday Pickup
					'objFedExWSClass.AddNewNode "SpecialServiceTypes", "SATURDAY_PICKUP"
				objFedExWSClass.WriteParent "SpecialServicesRequested", "/"
			End If
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// End: FDXE Special Services
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
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
				If scShipFromWeightUnit="LBS" Then
				    tmpOuncesDec = CDbl(tmpOunces)/16
                Else
                    tmpOuncesDec = tmpOunces
                End If

				objFedExWSClass.WriteParent "RequestedPackageLineItems", ""

				  objFedExWSClass.AddNewNode "SequenceNumber", q
				  objFedExWSClass.AddNewNode "GroupNumber", q
				  objFedExWSClass.AddNewNode "GroupPackageCount", 1

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

				  if ((pcv_strPackaging="YOUR_PACKAGING") AND (session("FedEXWSPackLength"&q)<>"" AND session("FedEXWSPackWidth"&q)<>"" AND session("FedEXWSPackHeight"&q)<>"")) then
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

	'// Print out our newly formed request xml
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(fedex_postdataWS)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Send Our Transaction.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.SendXMLRequest(fedex_postdataWS)

	'// Print out our response
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(FEDEXWS_result)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Load Our Response.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.LoadXMLResults(FEDEXWS_result)

	pcv_strErrorMsgWS = cSTR("")
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for errors from FedEx.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'call objFedExWSClass.XMLResponseVerify(ErrPageName)
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
	'call objFedExWSClass.pcs_LogTransaction(fedex_postdataWS, "Rate_" & pcv_strCarrierCode & "_" & pcv_strLogID & "_Req.xml", pcv_strFedExBaselineLogging)
	'// Log our Response
	'call objFedExWSClass.pcs_LogTransaction(FEDEXWS_result, "Rate_" & pcv_strCarrierCode & "_" & pcv_strLogID & "_Res.xml", pcv_strFedExBaselineLogging)

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Redirect with a Message OR complete some task.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	If NOT len(pcv_strErrorMsgWS)>0 Then
		arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "ServiceType")
		If scHideEstimateDeliveryTimes="1" Then
			arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
			arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
		Else
			arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & "TransitTime:"&objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "TransitTime")
			arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
		End If
		arryFedExWSRateType = arryFedExWSRateType & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/RateType")
		arryFedExWSSpecialRatingType = arryFedExWSSpecialRatingType & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/SpecialRatingApplied")

    '// Use net charge (no tax) for CA rates
    If scShipFromPostalCountry="CA" Then
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/TotalNetFedExCharge/Amount")
    Else
			arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/TotalNetCharge/Amount")
    End If

	End If

	set objFedExWSClass = nothing
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// END: FedEx Ground
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////

	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// START: FedEx Freight
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	set objFedExWSClass = New pcFedExWSClass

	fedex_postdataWS=""
	FEDEXWS_result=""
  
  pcv_strCarrierCode = "FXFR"

	objFedExWSClass.NewXMLSubscription pcv_strMethodNameWS, FedExkey, FedExPassword, FedExWS_RateVersion, "rate"

		objFedExWSClass.WriteParent "ClientDetail", ""
			objFedExWSClass.AddNewNode "AccountNumber", FedExAccountNumber
			objFedExWSClass.AddNewNode "MeterNumber", FedExMeterNumber
			objFedExWSClass.AddNewNode "ClientProductId", FedExWS_RateVersion
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
		objFedExWSClass.WriteSingleParent "CarrierCodes", pcv_strCarrierCode

		objFedExWSClass.WriteParent "RequestedShipment", ""

			'// Ship Date
			pcAddDay = 	FEDEXWS_ADDDAY
			if pcAddDay&""="" then
				pcAddDay = 0
			end if

			pcFedExShipDate = DateAdd("D", pcAddDay, Date)
			pcWeekDay = WeekDay(pcFedExShipDate)
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

			objFedExWSClass.AddNewNode "ServiceType", ""
			objFedExWSClass.AddNewNode "DropoffType", FEDEXWS_DROPOFF_TYPE
			objFedExWSClass.AddNewNode "PackagingType", pcv_strPackaging

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
				objFedExWSClass.WriteSingleParent "Value", totalPounds + totalOunces
			objFedExWSClass.WriteParent "TotalWeight", "/"
	
			If pcv_strRateRequestType = "PREFERRED" Then
				objFedExWSClass.AddNewNode "PreferredCurrency", FEDEXWS_PREFERRED_CURRENCY
			End If

			objFedExWSClass.WriteParent "Shipper", ""
				objFedExWSClass.WriteParent "Contact", ""
					objFedExWSClass.AddNewNode "PersonName", scOriginPersonName
					objFedExWSClass.AddNewNode "CompanyName", scShipFromName
					objFedExWSClass.AddNewNode "PhoneNumber", scOriginPhoneNumber
					objFedExWSClass.AddNewNode "EMailAddress", scOriginEmailAddress
				objFedExWSClass.WriteParent "Contact", "/"

				objFedExWSClass.WriteParent "Address", ""
					objFedExWSClass.AddNewNode "StreetLines", scShipFromAddress1
					objFedExWSClass.AddNewNode "City", scShipFromCity
					if FedExRequiresStateProvince(scShipFromPostalCountry) then
						objFedExWSClass.AddNewNode "StateOrProvinceCode", FedExCorrectStateProvince(scShipFromState)
					end if
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
					if FedExRequiresStateProvince(Universal_destination_country) then
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
			objFedExWSClass.WriteParent "ShippingChargesPayment", "/"

			'// ShipDate
			'objFedExWSClass.AddNewNode "ShipDate", "2011-08-03T14:33:57+05:30"
			'objFedExWSClass.WriteParent "ShipDate", ""
			'	objFedExWSClass.AddNewNode "ShipDate", "2011-08-03"
			'objFedExWSClass.WriteParent "ShipDate", "/"

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start: FDXE Special Services
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			If pcv_TmpSaturdayDelivery<>"0" Then
				objFedExWSClass.WriteParent "SpecialServicesRequested", ""
					'// Saturday Delivery
					objFedExWSClass.AddNewNode "SpecialServiceTypes", "SATURDAY_DELIVERY"
					'// Saturday Pickup
					'objFedExWSClass.AddNewNode "SpecialServiceTypes", "SATURDAY_PICKUP"
				objFedExWSClass.WriteParent "SpecialServicesRequested", "/"
			End If
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// End: FDXE Special Services
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// Start: FDXE International
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'objFedExWSClass.WriteParent "CustomsClearanceDetail", ""

			'// Custom Value
			'objFedExWSClass.WriteParent "CustomsValue", ""
			'		objFedExWSClass.AddNewNode "Currency", "USD"
			'		objFedExWSClass.AddNewNode "Amount", "250"
			'objFedExWSClass.WriteParent "CustomsValue", "/"

			'objFedExWSClass.WriteParent "CustomsClearanceDetail", "/"
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// End: FDXE International
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
				If scShipFromWeightUnit="LBS" Then
				    tmpOuncesDec = CDbl(tmpOunces)/16
                Else
                    tmpOuncesDec = tmpOunces
                End If
				objFedExWSClass.WriteParent "RequestedPackageLineItems", ""

				objFedExWSClass.AddNewNode "SequenceNumber", q
				objFedExWSClass.AddNewNode "GroupNumber", q
				objFedExWSClass.AddNewNode "GroupPackageCount", pcv_intPackageNum

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

				if ((pcv_strPackaging="YOUR_PACKAGING") AND (session("FedEXWSPackLength"&q)<>"" AND session("FedEXWSPackWidth"&q)<>"" AND session("FedEXWSPackHeight"&q)<>"")) then
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

				objFedExWSClass.WriteParent "RequestedPackageLineItems", "/"

			next

		objFedExWSClass.WriteParent "RequestedShipment", "/"

	objFedExWSClass.EndXMLTransaction pcv_strMethodNameWS


	'// Print out our newly formed request xml
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(fedex_postdataWS)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Send Our Transaction.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.SendXMLRequest(fedex_postdataWS)

	'// Print out our response
	'response.Clear()
	'response.ContentType="text/xml"
	'response.Write(FEDEXWS_result)
	'response.End()

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Load Our Response.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	call objFedExWSClass.LoadXMLResults(FEDEXWS_result)

	pcv_strErrorMsgWS = cSTR("")

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for errors from FedEx.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'call objFedExWSClass.XMLResponseVerify(ErrPageName)
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
	'call objFedExWSClass.pcs_LogTransaction(fedex_postdataWS, "Rate_" & pcv_strCarrierCode & "_" & pcv_strLogID & "_Req.xml", pcv_strFedExBaselineLogging)
	'// Log our Response
	'call objFedExWSClass.pcs_LogTransaction(FEDEXWS_result, "Rate_" & pcv_strCarrierCode & "_" & pcv_strLogID & "_Res.xml", pcv_strFedExBaselineLogging)

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Redirect with a Message OR complete some task.
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



	If NOT len(pcv_strErrorMsgWS)>0 Then
		arryFedExWSService = arryFedExWSService & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "ServiceType")
		If scHideEstimateDeliveryTimes="1" Then
			arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & ""
			arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
		Else
			arrFedExWSDeliveryDate = arrFedExWSDeliveryDate & "TransitTime:"&objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "TransitTime")
			arrFedExWSDeliveryTime = arrFedExWSDeliveryTime & ""
		End If
		arryFedExWSRateType = arryFedExWSRateType & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/RateType")
		arryFedExWSSpecialRatingType = arryFedExWSSpecialRatingType & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/SpecialRatingApplied")
		arryFedExWSRate = arryFedExWSRate & objFedExWSClass.ReadResponsesArray("//RateReplyDetails", "RatedShipmentDetails/ShipmentRateDetail/TotalNetFedExCharge/Amount")
	End If

	set objFedExWSClass = nothing
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// END: FedEx Freight
	'/////////////////////////////////////////////////////////////////////////////////////////////////////////////



	' trim the last comma if there is one
	'xStringLength = len(ReadResponsesArray)
	'if xStringLength>0 then
	'	ReadResponsesArray = left(ReadResponsesArray,(xStringLength-1))
	'end if

	'Split Arrays
	dim intRateIndexWS
	dim pcFedExWSMultiArry(50,5)
	for z=0 to 50
		pcFedExWSMultiArry(z,1)=0
	next

	pcStrTempFedExService=split(arryFedExWSService,",")
	pcStrTempFexExRate=split(arryFedExWSRate,",")
	pcStrTempFexExRateType=split(arryFedExWSRateType,",")
	pcStrTempFexExSpecialRatingType=split(arryFedExWSSpecialRatingType,",")
	pcStrTempFedExDeliveryDate=split(arrFedExWSDeliveryDate,",")
	pcStrTempFedExDeliveryTime=split(arrFedExWSDeliveryTime,",")

	For t=0 To ubound(pcStrTempFedExService) - 1
		For idx = 0 To UBound(FedExWS_ShipmentTypes) - 1
			Service = FedExWS_ShipmentTypes(idx)
			If pcStrTempFedExService(t) = Service Then
				rateType = ""
				
				if t < ubound(pcStrTempFexExRateType) then
					if InStr(pcStrTempFexExRateType(t), "ACCOUNT") > 0 then
						rateType = "Account"
					elseif InStr(pcStrTempFexExRateType(t), "LIST") > 0 then
						rateType = "List"
					end if
	
					if pcStrTempFexExRateType(t) = "PREFERRED" then
						rateType = "Preferred " & rateType
					end if
				end if

				if t < ubound(pcStrTempFexExSpecialRatingType) then
					if pcStrTempFexExSpecialRatingType(t) = "FEDEX_ONE_RATE" then
						rateType = "One Rate"
					end if
				end if

				If pcStrTempFedExService(t) = "GROUND_HOME_DELIVERY" Then
					If Universal_destination_country="US" Then
						intRateIndexWS = idx
						pcFedExWSMultiArry(intRateIndexWS,2)=FedExWS_ShipmentName(pcStrTempFedExService(t))
						pcFedExWSMultiArry(intRateIndexWS,3)=pcStrTempFedExService(t)
						pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
						pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
					End If
				Else
					intRateIndexWS = idx
					pcFedExWSMultiArry(intRateIndexWS,2)=FedExWS_ShipmentName(pcStrTempFedExService(t))
					pcFedExWSMultiArry(intRateIndexWS,3)=pcStrTempFedExService(t)
					pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
					pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
				End If
                If pcStrTempFedExService(t) = "FEDEX_GROUND" Then
                    If Universal_destination_country="CA" Then
                        intRateIndexWS = idx
                        pcFedExWSMultiArry(intRateIndexWS,2)="FedEx International Ground<sup>&reg;</sup>"
                        pcFedExWSMultiArry(intRateIndexWS,3)="INTERNATIONAL_GROUND"
                        pcFedExWSMultiArry(intRateIndexWS,4)=pcStrTempFedExDeliveryDate(t)
                        pcFedExWSMultiArry(intRateIndexWS,5)=pcStrTempFedExDeliveryTime(t)
                    End If
                End If
			End If
		Next

		tempRate=pcFedExWSMultiArry(intRateIndexWS,1)
		pcFedExWSMultiArry(intRateIndexWS,1)=cdbl(tempRate)+cdbl(pcStrTempFexExRate(t))
	Next

	for z=0 to 50
		if pcFedExWSMultiArry(z,1)>0 then
			intNoTime = 0
			pcv_strFormattedDate = ""

			tmpDeliveryDayOfWeek = pcFedExWSMultiArry(z,4)

			If instr(tmpDeliveryDayOfWeek,"TransitTime") Then
				If instr(tmpDeliveryDayOfWeek,"SPTransitTime") Then
					tmpDeliveryDayOfWeek = replace(tmpDeliveryDayOfWeek, "SPTransitTime:","")
					intNoTime = 2
					select case tmpDeliveryDayOfWeek
						case "EIGHTEEN_DAYS"
							tmpDeliveryDayOfWeek = "18"
						case "EIGHT_DAYS"
							tmpDeliveryDayOfWeek = "8"
						case "ELEVEN_DAYS"
							tmpDeliveryDayOfWeek = "11"
						case "FIFTEEN_DAYS"
							tmpDeliveryDayOfWeek = "15"
						case "FIVE_DAYS"
							tmpDeliveryDayOfWeek = "5"
						case "FOURTEEN_DAYS"
							tmpDeliveryDayOfWeek = "14"
						case "FOUR_DAYS"
							tmpDeliveryDayOfWeek = "4"
						case "NINETEEN_DAYS"
							tmpDeliveryDayOfWeek = "19"
						case "NINE_DAYS"
							tmpDeliveryDayOfWeek = "9"
						case "ONE_DAY"
							tmpDeliveryDayOfWeek = "1"
						case "SEVENTEEN_DAYS"
							tmpDeliveryDayOfWeek = "17"
						case "SEVEN_DAYS"
							tmpDeliveryDayOfWeek = "7"
						case "SIXTEEN_DAYS"
							tmpDeliveryDayOfWeek = "16"
						case "SIX_DAYS"
							tmpDeliveryDayOfWeek = "6"
						case "TEN_DAYS"
							tmpDeliveryDayOfWeek = "10"
						case "THIRTEEN_DAYS"
							tmpDeliveryDayOfWeek = "13"
						case "THREE_DAYS"
							tmpDeliveryDayOfWeek = "3"
						case "TWELVE_DAYS"
							tmpDeliveryDayOfWeek = "12"
						case "TWENTY_DAYS"
							tmpDeliveryDayOfWeek = "20"
						case "TWO_DAYS"
							tmpDeliveryDayOfWeek = "2"
						case else
							tmpDeliveryDayOfWeek = "Unknown"
					end select
					tmpDeliveryBetween = pcFedExWSMultiArry(z,5)
					tmpDeliveryBetween = replace(tmpDeliveryBetween, "SPTransitTime:","")
					select case tmpDeliveryBetween
						case "EIGHTEEN_DAYS"
							tmpDeliveryBetween = "18"
						case "EIGHT_DAYS"
							tmpDeliveryBetween = "8"
						case "ELEVEN_DAYS"
							tmpDeliveryBetween = "11"
						case "FIFTEEN_DAYS"
							tmpDeliveryBetween = "15"
						case "FIVE_DAYS"
							tmpDeliveryBetween = "5"
						case "FOURTEEN_DAYS"
							tmpDeliveryBetween = "14"
						case "FOUR_DAYS"
							tmpDeliveryBetween = "4"
						case "NINETEEN_DAYS"
							tmpDeliveryBetween = "19"
						case "NINE_DAYS"
							tmpDeliveryBetween = "9"
						case "ONE_DAY"
							tmpDeliveryBetween = "1"
						case "SEVENTEEN_DAYS"
							tmpDeliveryBetween = "17"
						case "SEVEN_DAYS"
							tmpDeliveryBetween = "7"
						case "SIXTEEN_DAYS"
							tmpDeliveryBetween = "16"
						case "SIX_DAYS"
							tmpDeliveryBetween = "6"
						case "TEN_DAYS"
							tmpDeliveryBetween = "10"
						case "THIRTEEN_DAYS"
							tmpDeliveryBetween = "13"
						case "THREE_DAYS"
							tmpDeliveryBetween = "3"
						case "TWELVE_DAYS"
							tmpDeliveryBetween = "12"
						case "TWENTY_DAYS"
							tmpDeliveryBetween = "20"
						case "TWO_DAYS"
							tmpDeliveryBetween = "2"
						case else
							tmpDeliveryBetween = "Unknown"
					end select
				Else
				tmpDeliveryDayOfWeek = replace(tmpDeliveryDayOfWeek, "TransitTime:","")
				intNoTime = 1
				select case tmpDeliveryDayOfWeek
					case "EIGHTEEN_DAYS"
						tmpDeliveryDayOfWeek = "18"
					case "EIGHT_DAYS"
						tmpDeliveryDayOfWeek = "8"
					case "ELEVEN_DAYS"
						tmpDeliveryDayOfWeek = "11"
					case "FIFTEEN_DAYS"
						tmpDeliveryDayOfWeek = "15"
					case "FIVE_DAYS"
						tmpDeliveryDayOfWeek = "5"
					case "FOURTEEN_DAYS"
						tmpDeliveryDayOfWeek = "14"
					case "FOUR_DAYS"
						tmpDeliveryDayOfWeek = "4"
					case "NINETEEN_DAYS"
						tmpDeliveryDayOfWeek = "19"
					case "NINE_DAYS"
						tmpDeliveryDayOfWeek = "9"
					case "ONE_DAY"
						tmpDeliveryDayOfWeek = "1"
					case "SEVENTEEN_DAYS"
						tmpDeliveryDayOfWeek = "17"
					case "SEVEN_DAYS"
						tmpDeliveryDayOfWeek = "7"
					case "SIXTEEN_DAYS"
						tmpDeliveryDayOfWeek = "16"
					case "SIX_DAYS"
						tmpDeliveryDayOfWeek = "6"
					case "TEN_DAYS"
						tmpDeliveryDayOfWeek = "10"
					case "THIRTEEN_DAYS"
						tmpDeliveryDayOfWeek = "13"
					case "THREE_DAYS"
						tmpDeliveryDayOfWeek = "3"
					case "TWELVE_DAYS"
						tmpDeliveryDayOfWeek = "12"
					case "TWENTY_DAYS"
						tmpDeliveryDayOfWeek = "20"
					case "TWO_DAYS"
						tmpDeliveryDayOfWeek = "2"
					case else
						tmpDeliveryDayOfWeek = "Unknown"
				end select
				End If
			End If

			If intNoTime = 0 Then
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

				tmpDeliveryTimestamp = pcFedExWSMultiArry(z,5)
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
					tmpHour = Cint(tmpTimeHour) - Cint(12)
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
			End If
			If intNoTime = 1 or intNoTime = 2 Then
				if tmpDeliveryDayOfWeek = "Unknown" then
					pcv_strFormattedDate = tmpDeliveryDayOfWeek
				else
					pcAddDay = 	FEDEXWS_ADDDAY
					if pcAddDay&""="" then
						pcAddDay = 0
					end if
					DatePlus = Date()+Cint(pcAddDay)
                    pcv_strDeliveryDate = DateAdd("D", Cint(tmpDeliveryDayOfWeek), DatePlus)
                    pcWeekDay = WeekDay(pcv_strDeliveryDate)
                    if pcWeekDay = "1" then
                        'Add a day to the current date
                        DatePlus = DatePlus + 1
                    end if
                    'if saturday pickup=0 then we need to shift a saturday date to monday
                    if pcWeekDay = "7" AND FEDEXWS_SATURDAYPICKUP = "0" then
                        'Add 2 days to the current date
                        DatePlus = DatePlus + 2
                    end if
					pcv_strFormattedDate = FormatDateTime(DateAdd("D", Cint(tmpDeliveryDayOfWeek), DatePlus), 1)&" 7:00 PM"
					if intNoTime = 2 then
						tmpFromSPDate1 = FormatDateTime(DateAdd("D", Cint(tmpDeliveryDayOfWeek), DatePlus), 1)
						if instr(tmpFromSPDate1, "Sunday") then
							tmpDeliveryDayOfWeek = int(tmpDeliveryDayOfWeek)+1
							tmpFromSPDate1 = FormatDateTime(DateAdd("D", Cint(tmpDeliveryDayOfWeek), DatePlus), 1)
						end if
						tmpToSPDate1 = FormatDateTime(DateAdd("D", Cint(tmpDeliveryBetween), DatePlus), 1)
						
						if instr(tmpToSPDate1, "Sunday") then
							tmpDeliveryBetween = int(tmpDeliveryBetween)+1
							tmpToSPDate1 = FormatDateTime(DateAdd("D", Cint(tmpDeliveryBetween), DatePlus), 1)
						end if
					
						pcv_strFormattedDate = "Between: "&tmpFromSPDate1&" <br>AND "&tmpToSPDate1&" "
					end if
				end if
			Else
				pcv_strFormattedDate = tmpDeliveryDayOfWeek&", "&tmpDeliveryMonth&" "&tmpDeliveryDay&", "&tmpDeliveryYear&" "&tmpHour&":"&tmpTimeMinutes&" "&tmpAMPM
			End If

			availableShipStr=availableShipStr&"|?|FedExWS|"&pcFedExWSMultiArry(z,3)&"|"&pcFedExWSMultiArry(z,2)&"|"&pcFedExWSMultiArry(z,1)&"|"&pcv_strFormattedDate
			iFedExWSFlag=1
			
		end if
	next
    
End If '// If (FedEXWS_active=true Or FedExWS_active="-1") Then

%>