<% if ups_active=true or ups_active="-1" then
	Dim iUPSFlag
	iUPSFlag=0
	iUPSActive=1
	'//UPS Rates
	
	'// Generate Access Request XML
	ups_accessrequest=""
	ups_accessrequest="<?xml version=""1.0""?>"
	ups_accessrequest=ups_accessrequest&"<AccessRequest xml:lang=""en-US"">"
	ups_accessrequest=ups_accessrequest&"<AccessLicenseNumber>"&ups_license_key&"</AccessLicenseNumber>"
	ups_accessrequest=ups_accessrequest&"<UserId>"&ups_userid&"</UserId>"
	ups_accessrequest=ups_accessrequest&"<Password>"&ups_password&"</Password>"
	ups_accessrequest=ups_accessrequest&"</AccessRequest>"
	
	'// Generate Rate Request XML
	ups_raterequest=""
	ups_raterequest=ups_raterequest&"<?xml version=""1.0""?>"
	ups_raterequest=ups_raterequest&"<RatingServiceSelectionRequest xml:lang=""en-US"">"
	ups_raterequest=ups_raterequest&"<Request>"
	ups_raterequest=ups_raterequest&"<TransactionReference>"
	ups_raterequest=ups_raterequest&"<CustomerContext>Rating and Service</CustomerContext>"
	ups_raterequest=ups_raterequest&"<XpciVersion>1.0</XpciVersion>"
	ups_raterequest=ups_raterequest&"</TransactionReference>"
	ups_raterequest=ups_raterequest&"<RequestAction>rate</RequestAction>"
	ups_raterequest=ups_raterequest&"<RequestOption>shop</RequestOption>"
	ups_raterequest=ups_raterequest&"</Request>"
	ups_raterequest=ups_raterequest&"<PickupType>"
	ups_raterequest=ups_raterequest&"<Code>"&UPS_PICKUP_TYPE&"</Code>"
	ups_raterequest=ups_raterequest&"</PickupType>"
	if UPS_CLASSIFICATION_TYPE<>"" then
		ups_raterequest=ups_raterequest&"<CustomerClassification>"
		ups_raterequest=ups_raterequest&"<Code>"&UPS_CLASSIFICATION_TYPE&"</Code>"
		ups_raterequest=ups_raterequest&"</CustomerClassification>"
	end if
	ups_raterequest=ups_raterequest&"<Shipment>"
	ups_raterequest=ups_raterequest&"<Shipper>"
	if pcv_UseNegotiatedRates=1 then
		if pcv_UPSShipperNumber<>"" then
			ups_raterequest=ups_raterequest&"<ShipperNumber>"&pcv_UPSShipperNumber&"</ShipperNumber>"
		end if
	end if
	ups_raterequest=ups_raterequest&"<Address>"
	ups_raterequest=ups_raterequest&"<City>"&UPS_ShipFromCity&"</City>"
	ups_raterequest=ups_raterequest&"<StateProvinceCode>"&UPS_ShipFromState&"</StateProvinceCode>"
	ups_raterequest=ups_raterequest&"<PostalCode>"&UPS_ShipFromPostalCode&"</PostalCode>"
	ups_raterequest=ups_raterequest&"<CountryCode>"&UPS_ShipFromPostalCountry&"</CountryCode>"
	ups_raterequest=ups_raterequest&"</Address>"
	ups_raterequest=ups_raterequest&"</Shipper>"
	ups_raterequest=ups_raterequest&"<ShipTo>"
	ups_raterequest=ups_raterequest&"<Address>"
	ups_raterequest=ups_raterequest&"<City>"&Universal_destination_city&"</City>"
	ups_raterequest=ups_raterequest&"<StateProvinceCode>"&Universal_destination_provOrState&"</StateProvinceCode>"
	ups_destination_postal=replace(Universal_destination_postal, " ","")
	ups_destination_postal=replace(ups_destination_postal,"-","")
	ups_raterequest=ups_raterequest&"<PostalCode>"&ups_destination_postal&"</PostalCode>"
	ups_raterequest=ups_raterequest&"<CountryCode>"&Universal_destination_country&"</CountryCode>"
	If pResidentialShipping<>"0" then
		ups_raterequest=ups_raterequest&"<ResidentialAddress>1</ResidentialAddress>"
	else
		ups_raterequest=ups_raterequest&"<ResidentialAddress>0</ResidentialAddress>"
	end if
	ups_raterequest=ups_raterequest&"</Address>"
	ups_raterequest=ups_raterequest&"</ShipTo>"
	for q=1 to pcv_intPackageNum
		ups_raterequest=ups_raterequest&"<Package>"
		ups_raterequest=ups_raterequest&"<PackagingType>"
		ups_raterequest=ups_raterequest&"<Code>"&UPS_PACKAGE_TYPE&"</Code>"
		ups_raterequest=ups_raterequest&"<Description>Package</Description>"
		ups_raterequest=ups_raterequest&"</PackagingType>"
		ups_raterequest=ups_raterequest&"<Description>Rate Shopping</Description>"
		ups_raterequest=ups_raterequest&"<Dimensions>"
		pUPS_DIM_UNIT=ucase(UPS_DIM_UNIT)
		if q>1 then
			pcv_intOSheight=UPS_HEIGHT
			pcv_intOSwidth=UPS_WIDTH
			pcv_intOSlength=UPS_LENGTH
		end if
		if scShipFromWeightUnit="KGS" AND pUPS_DIM_UNIT="IN" then
			pUPS_DIM_UNIT="CM"
			pcv_intOSlength=pcv_intOSlength*2.54
			pcv_intOSwidth=pcv_intOSwidth*2.54
			pcv_intOSheight=pcv_intOSheight*2.54
		end if
		if scShipFromWeightUnit="LBS" AND pUPS_DIM_UNIT="CM" then
			pUPS_DIM_UNIT="IN"
			pcv_intOSlength=pcv_intOSlength/2.54
			pcv_intOSwidth=pcv_intOSwidth/2.54
			pcv_intOSheight=pcv_intOSheight/2.54
		end if
		ups_raterequest=ups_raterequest&"<UnitOfMeasurement><Code>"&pUPS_DIM_UNIT&"</Code></UnitOfMeasurement>"
		ups_raterequest=ups_raterequest&"<Length>"&pc_dimensions(session("UPSPackLength"&q))&"</Length>" 'Between 1 and 108.00
		ups_raterequest=ups_raterequest&"<Width>"&pc_dimensions(session("UPSPackWidth"&q))&"</Width>" 'Between 1 and 108.00
		ups_raterequest=ups_raterequest&"<Height>"&pc_dimensions(session("UPSPackHeight"&q))&"</Height>" 'Between 1 and 108.00
		ups_raterequest=ups_raterequest&"</Dimensions>"
		ups_raterequest=ups_raterequest&"<PackageWeight>"
		ups_raterequest=ups_raterequest&"<UnitOfMeasurement>"
		if scShipFromWeightUnit="KGS" then
			ups_raterequest=ups_raterequest&"<Code>KGS</Code>"
		else
			ups_raterequest=ups_raterequest&"<Code>LBS</Code>"
		end if
		ups_raterequest=ups_raterequest&"</UnitOfMeasurement>"
		ups_raterequest=ups_raterequest&"<Weight>"&pc_dimensions(session("UPSPackWeight"&q))&"</Weight>" '0.1 to 150.0
		ups_raterequest=ups_raterequest&"</PackageWeight>"
		ups_raterequest=ups_raterequest&"<OversizePackage>0</OversizePackage>"
		ups_raterequest=ups_raterequest&"<PackageServiceOptions>"
		ups_raterequest=ups_raterequest&"<InsuredValue>"
		ups_raterequest=ups_raterequest&"<CurrencyCode>USD</CurrencyCode>"

		pcv_TempPackPrice=session("UPSPackPrice"&q)			
		If pcv_TempPackPrice="" Then
			pcv_TempPackPrice="100.00"
		End If
		
		ups_raterequest=ups_raterequest&"<MonetaryValue>"&replace(money(pcv_TempPackPrice),",","")&"</MonetaryValue>"
		ups_raterequest=ups_raterequest&"</InsuredValue>"
		ups_raterequest=ups_raterequest&"</PackageServiceOptions>"
		ups_raterequest=ups_raterequest&"</Package>"
	next
	if pcv_UseNegotiatedRates=1 then
		ups_raterequest=ups_raterequest&"<RateInformation>"
			ups_raterequest=ups_raterequest&"<NegotiatedRatesIndicator/>"
		ups_raterequest=ups_raterequest&"</RateInformation>"
	end if
	ups_raterequest=ups_raterequest&"</Shipment>"
	ups_raterequest=ups_raterequest&"</RatingServiceSelectionRequest>"
	
'	response.clear
'	response.contenttype = "text/xml"
'	response.write ups_raterequest
'	response.end
	
	'get URL to post to
	ups_URL="https://onlinetools.ups.com/ups.app/xml/Rate"
	
	ups_postdata = ups_accessrequest & ups_raterequest

	Set srvUPSXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	srvUPSXmlHttp.open "POST", ups_URL, false
	srvUPSXmlHttp.send(ups_postdata)
	UPS_result = srvUPSXmlHttp.responseText
	
'	response.clear
'	response.contenttype = "text/xml"
'	response.write UPS_result
'	response.end

	Set UPSXMLdoc = server.CreateObject("Msxml2.DOMDocument"&scXML)
	UPSXMLDoc.async = false 
	if UPSXMLDOC.loadXML(UPS_result) then ' if loading from a string
		set objLst = UPSXMLDOC.getElementsByTagName("RatedShipment") 
		for i = 0 to (objLst.length - 1)
			varFlag=0
			for j=0 to ((objLst.item(i).childNodes.length)-1)
				If objLst.item(i).childNodes(j).nodeName="Service" then
					serviceVar=objLst.item(i).childNodes(j).text
					select case serviceVar
					case "01"
						if pcv_UPSCanadaOrigin=0 then
							availableShipStr=availableShipStr&"|?|UPS|01|"&pServiceCodeString01
						else
							availableShipStr=availableShipStr&"|?|UPS|01|"&pServiceCodeString01
						end if
						varFlag=1
						iUPSFlag=1
					case "02"
						if pcv_UPSCanadaOrigin=0 then
							availableShipStr=availableShipStr&"|?|UPS|02|"&pServiceCodeString02
						else
							availableShipStr=availableShipStr&"|?|UPS|02|"&pServiceCodeString02
						end if
						varFlag=1
						iUPSFlag=1
					case "03"
						availableShipStr=availableShipStr&"|?|UPS|03|"&pServiceCodeString03
						varFlag=1
						iUPSFlag=1
					case "07"
						availableShipStr=availableShipStr&"|?|UPS|07|"&pServiceCodeString07
						varFlag=1
						iUPSFlag=1
					case "08"
						availableShipStr=availableShipStr&"|?|UPS|08|"&pServiceCodeString08
						varFlag=1
						iUPSFlag=1
					case "11"
						availableShipStr=availableShipStr&"|?|UPS|11|"&pServiceCodeString11
						varFlag=1
						iUPSFlag=1
					case "12"
						availableShipStr=availableShipStr&"|?|UPS|12|"&pServiceCodeString12
						varFlag=1
						iUPSFlag=1
					case "13"
						if pcv_UPSCanadaOrigin=0 then
							availableShipStr=availableShipStr&"|?|UPS|13|"&pServiceCodeString13
						else
							availableShipStr=availableShipStr&"|?|UPS|13|"&pServiceCodeString13
						end if
						varFlag=1
						iUPSFlag=1
					case "14"
						if pcv_UPSCanadaOrigin=0 then
							availableShipStr=availableShipStr&"|?|UPS|14|"&pServiceCodeString14
						else
							availableShipStr=availableShipStr&"|?|UPS|14|"&pServiceCodeString14
						end if
						varFlag=1
						iUPSFlag=1
					case "54"
						if pcv_UPSCanadaOrigin=0 then
							availableShipStr=availableShipStr&"|?|UPS|54|"&pServiceCodeString54
						else
							availableShipStr=availableShipStr&"|?|UPS|54|"&pServiceCodeString54
						end if							
						varFlag=1
						iUPSFlag=1
					case "59"
						availableShipStr=availableShipStr&"|?|UPS|59|"&pServiceCodeString59
						varFlag=1
						iUPSFlag=1
					case "65"
						availableShipStr=availableShipStr&"|?|UPS|65|"&pServiceCodeString65
						varFlag=1
						iUPSFlag=1
					end select
				End if
				
				'// Get Monetary Value
				If objLst.item(i).childNodes(j).nodeName="TotalCharges" then
					for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
						if objLst.item(i).childNodes(j).childNodes(k).nodeName="MonetaryValue" then
							availableShipStr=availableShipStr&"|"&objLst.item(i).childNodes(j).childNodes(k).text
						end if
					next
				End if

				if pcv_UseNegotiatedRates=1 then
					If objLst.item(i).childNodes(j).nodeName="NegotiatedRates" then
						for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							if objLst.item(i).childNodes(j).childNodes(k).childNodes(0).childNodes(1).nodeName="MonetaryValue" then
								availableShipStr=availableShipStr&"|"&objLst.item(i).childNodes(j).childNodes(k).childNodes(0).childNodes(1).text
							else
								availableShipStr=availableShipStr&"|NONE"
							end if
						next
					End if
				end if
								
				If objLst.item(i).childNodes(j).nodeName="GuaranteedDaysToDelivery" AND varFlag=1 then
					if objLst.item(i).childNodes(j).text="1" then
						availableShipStr=availableShipStr&"|Next Day"
					else
						if objLst.item(i).childNodes(j).text<>"" then
							availableShipStr=availableShipStr&"|"&objLst.item(i).childNodes(j).text&" Days"
						else
							availableShipStr=availableShipStr&"|NA"
						end if
					end if
				End If
				If objLst.item(i).childNodes(j).nodeName="ScheduledDeliveryTime" AND varFlag=1 then
					If objLst.item(i).childNodes(j).text<>"" then
						availableShipStr=availableShipStr&" by "&objLst.item(i).childNodes(j).text
					end if
				End If
			next
			
		next 
	end if
end if 'if ups is active
%>