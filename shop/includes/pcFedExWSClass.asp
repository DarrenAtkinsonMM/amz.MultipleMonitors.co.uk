<%
FedExWS_RateVersion					= "16"
FedExWS_RegistrationVersion	= "6"
FedExWS_ShipVersion 				= "15"
FedExWS_TrackVersion 				= "9"
FedExWS_CloseVersion 				= "3"

FedExWS_ShipmentID = 9
FedExWS_ShipmentTypes = Array( _
	"FIRST_OVERNIGHT", _
	"FEDEX_FIRST_FREIGHT", _
	"PRIORITY_OVERNIGHT", _
	"STANDARD_OVERNIGHT", _
	"FEDEX_2_DAY", _
	"FEDEX_2_DAY_AM", _
	"FEDEX_EXPRESS_SAVER", _
	"FEDEX_FREIGHT_PRIORITY", _
	"FEDEX_FREIGHT_ECONOMY", _
	"FEDEX_GROUND", _
	"GROUND_HOME_DELIVERY", _
	"INTERNATIONAL_GROUND", _
	"INTERNATIONAL_FIRST", _
	"INTERNATIONAL_PRIORITY", _
	"INTERNATIONAL_ECONOMY", _
	"FEDEX_1_DAY_FREIGHT", _
	"FEDEX_2_DAY_FREIGHT", _
	"FEDEX_3_DAY_FREIGHT", _
	"INTERNATIONAL_PRIORITY_FREIGHT", _
	"INTERNATIONAL_ECONOMY_FREIGHT", _
	"SMART_POST", _
	"FEDEX_ECONOMY_CANADA" _
)

Function FedExWS_ShipmentName(ServiceCode)
	ServiceName = ""

	Select Case ServiceCode
	Case "FIRST_OVERNIGHT":								ServiceName = "FedEx First Overnight<sup>&reg;</sup>"
	Case "FEDEX_FIRST_FREIGHT":						ServiceName = "FedEx First Overnight<sup>&reg;</sup> Freight"
	Case "PRIORITY_OVERNIGHT":						ServiceName = "FedEx Priority Overnight<sup>&reg;</sup>"
	Case "STANDARD_OVERNIGHT":						ServiceName = "FedEx Standard Overnight<sup>&reg;</sup>"
	Case "FEDEX_2_DAY":										ServiceName = "FedEx 2Day<sup>&reg;</sup>"
	Case "FEDEX_2_DAY_AM":								ServiceName = "FedEx 2Day<sup>&reg;</sup> A.M."
	Case "FEDEX_EXPRESS_SAVER":						ServiceName = "FedEx Express Saver<sup>&reg;</sup>"
	Case "FEDEX_FREIGHT_PRIORITY":				ServiceName = "FedEx Freight <sup>&reg;</sup> Priority"
	Case "FEDEX_FREIGHT_ECONOMY":					ServiceName = "FedEx Freight <sup>&reg;</sup> Economy"
	Case "FEDEX_GROUND":									ServiceName = "FedEx Ground<sup>&reg;</sup>"
	Case "GROUND_HOME_DELIVERY":					ServiceName = "FedEx Home Delivery<sup>&reg;</sup>"
	Case "INTERNATIONAL_GROUND":					ServiceName = "FedEx International Ground<sup>&reg;</sup>"
	Case "INTERNATIONAL_FIRST":						ServiceName = "FedEx International First<sup>&reg;</sup>"
	Case "INTERNATIONAL_PRIORITY":				ServiceName = "FedEx International Priority<sup>&reg;</sup>"
	Case "INTERNATIONAL_ECONOMY":					ServiceName = "FedEx International Economy<sup>&reg;</sup>"
	Case "FEDEX_1_DAY_FREIGHT":						ServiceName = "FedEx 1Day<sup>&reg;</sup> Freight"
	Case "FEDEX_2_DAY_FREIGHT":						ServiceName = "FedEx 2Day<sup>&reg;</sup> Freight"
	Case "FEDEX_3_DAY_FREIGHT":						ServiceName = "FedEx 3Day<sup>&reg;</sup> Freight"
	Case "INTERNATIONAL_PRIORITY_FREIGHT":ServiceName = "FedEx International Priority<sup>&reg;</sup> Freight"
	Case "INTERNATIONAL_ECONOMY_FREIGHT":	ServiceName = "FedEx International Economy<sup>&reg;</sup> Freight"
	Case "SMART_POST":										ServiceName = "FedEx SmartPost<sup>&reg;</sup>"
	Case "FEDEX_ECONOMY_CANADA":					ServiceName = "FedEx Economy (Canada)"
	End Select

	FedExWS_ShipmentName = ServiceName
End Function

Function FedExCorrectStateProvince(stateProvinceCode)
    If stateProvinceCode = "QC" Then
        FedExCorrectStateProvince = "PQ"
    Else
        FedExCorrectStateProvince = stateProvinceCode
    End If
End Function

Function FedExRequiresStateProvince(countryCode)
    required = false

    If countryCode = "US" Or countryCode = "CA" Or countryCode = "BR" Or countryCode = "IN" Or countryCode = "MX" Or countryCode = "PR" Then
        required = true
    End If

    FedExRequiresStateProvince = required
End Function

CSPTurnOn = 1

'// For Live
FedExWSURL =  "https://gateway.fedex.com:443/web-services"
pcv_strCSPKey = "CPTi545ATGa1CD89"
pcv_strCSPPassword = "8BB07q2XIIOFyNJeJQHMLv094"
pcv_strCSPSolutionID = "120"
pcv_strClientProductID = "EIPC"
pcv_strClientProductVersion = "3424"

'// FedEx Web Services 2014 Cert
'FedExWSURL =  "https://wsbeta.fedex.com:443/web-services"
'pcv_strCSPKey = "cwSZKpWoc8gO65Yo"
'pcv_strCSPPassword = "wuePomabvFDh9fI5BfDzz1bTc"
'pcv_strCSPSolutionID = "120"
'pcv_strClientProductID = "EIPC"
'pcv_strClientProductVersion = "2017"

'/////////////////////////////////////
'// Start building the class here
'/////////////////////////////////////
Class pcFedExWSClass

	private sub Class_Initialize()
		on error resume next
		Set srvFEDEXWSXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
		Set objOutputXMLDocWS = Server.CreateObject("Microsoft.XMLDOM")
		Set objFedExStream = Server.CreateObject("ADODB.Stream")
		Set objFEDEXXmlDoc = server.createobject("Msxml2.DOMDocument"&scXML)
		objFEDEXXmlDoc.async = False
		objFEDEXXmlDoc.validateOnParse = False
		if err.number>0 then
			err.clear
		end if
	end sub



	private sub Class_Terminate()
		'// clean it all up
		Set srvFEDEXWSXmlHttp = nothing
		Set objOutputXMLDocWS = nothing
		Set objFEDEXXmlDoc = nothing
		Set objFedExStream = nothing
	end sub

  Public Function GetXMLPrefix(FedExVersion)
    'GetXMLPrefix = "v" & FedExVersion & ":"
	GetXMLPrefix = ""
  End Function

  Public Function GetXMLNamespace(FedExVersion)
    GetXMLNamespace = ":v" & FedExVersion
  End Function


	Public Sub AddNewNodeAlt(NameOfNode, FedExVersion, ValueOfNode)
		fedex_xmlPrefix = GetXMLPrefix(FedExVersion)
		
		if len(ValueOfNode)>0 then
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&NameOfNode&">"&ValueOfNode&"</"&fedex_xmlPrefix&NameOfNode&">"&vbcrlf
		end if
	End Sub



	Public Sub WriteParentAlt(NameOfParent, FedExVersion, isClosing)
		fedex_xmlPrefix = GetXMLPrefix(FedExVersion)
		
		fedex_postdataWS=fedex_postdataWS&"<"&isClosing&fedex_xmlPrefix&NameOfParent&">"&vbcrlf
	End Sub



	Public Sub WriteSingleParentAlt(NameOfParent, FedExVersion, ValueOfParent)
		fedex_xmlPrefix = GetXMLPrefix(FedExVersion)
		
		if len(ValueOfParent)>0 then
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&NameOfParent&">"&ValueOfParent&"</"&fedex_xmlPrefix&NameOfParent&">"&vbcrlf
		end if
	End Sub

	'ship v9
	Public Sub AddNewNode(NameOfNode, ValueOfNode)
		if len(ValueOfNode)>0 then
			fedex_postdataWS=fedex_postdataWS&"<"&NameOfNode&">"&ValueOfNode&"</"&NameOfNode&">"&vbcrlf
		end if
	End Sub

	Public Sub WriteParent(NameOfParent, isClosing)
		fedex_postdataWS=fedex_postdataWS&"<"&isClosing&NameOfParent&">"&vbcrlf
	End Sub

	Public Sub WriteSingleParent(NameOfParent, ValueOfParent)
		if len(ValueOfParent)>0 then
			fedex_postdataWS=fedex_postdataWS&"<"&NameOfParent&">"&ValueOfParent&"</"&NameOfParent&">"&vbcrlf
		end if
	End Sub

	Public Sub NewXMLTransaction(NameOfMethod, FedEX_AccountNumber, FedEX_MeterNumber, FedEX_CarrierCode, CustomerTransactionIdentifier)
		fedex_postdataWS=""
		fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<"&NameOfMethod&" xmlns:api=""http://www.fedex.com/fsmapi"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation="""&NameOfMethod&".xsd"">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<RequestHeader>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<CustomerTransactionIdentifier>"&CustomerTransactionIdentifier&"</CustomerTransactionIdentifier>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<AccountNumber>"&FedEX_AccountNumber&"</AccountNumber>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<MeterNumber>"&FedEX_MeterNumber&"</MeterNumber>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<CarrierCode>"&FedEX_CarrierCode&"</CarrierCode>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</RequestHeader>"&vbcrlf
	End Sub



	Public Sub NewXMLCapture(NameOfMethod, FedEX_AccountNumber, FedEX_MeterNumber)
		fedex_postdataWS=""
		fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<"&NameOfMethod&" xmlns:api=""http://www.fedex.com/fsmapi"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation="""&NameOfMethod&".xsd"">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<RequestHeader>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<AccountNumber>"&FedEX_AccountNumber&"</AccountNumber>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<MeterNumber>"&FedEX_MeterNumber&"</MeterNumber>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</RequestHeader>"&vbcrlf
	End Sub

	Public Sub NewXMLSubscription(NameOfMethod, FedEX_Key, FedEX_Password, FedExVersion, FedExName)
		if FedExWS_UseNamespace then
			fedex_xmlPrefix = GetXMLPrefix(FedExVersion)
			fedex_xmlNamespace = GetXMLNamespace(FedExVersion)
		end if
		
		fedex_postdataWS=""
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns" & fedex_xmlNamespace & "=""http://fedex.com/ws/"&FedExName&"/v"&FedExVersion&""">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Header/>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Body>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&NameOfMethod&">"&vbcrlf

		fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"WebAuthenticationDetail>"&vbcrlf
		If CSPTurnOn = 1 Then
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"CspCredential>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"Key>" & pcv_strCSPKey & "</"&fedex_xmlPrefix&"Key>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"Password>" & pcv_strCSPPassword & "</"&fedex_xmlPrefix&"Password>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"</"&fedex_xmlPrefix&"CspCredential>"&vbcrlf
		End If
		If len(FedEX_Key)>0 AND len(FedEX_Password)> 0 Then
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"UserCredential>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"Key>"&FedEX_Key&"</"&fedex_xmlPrefix&"Key>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"Password>"&FedEX_Password&"</"&fedex_xmlPrefix&"Password>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"</"&fedex_xmlPrefix&"UserCredential>"&vbcrlf
		End If

		fedex_postdataWS=fedex_postdataWS&"</"&fedex_xmlPrefix&"WebAuthenticationDetail>"&vbcrlf

	End Sub

	Public Sub NewXMLLabelWS(NameOfMethod, FedExkey, FedExPassword, FedExAccountNumber, FedExMeterNumber, FedExVersion, FedExName)
		if FedExWS_UseNamespace then
			fedex_xmlPrefix = GetXMLPrefix(FedExVersion)
			fedex_xmlNamespace = GetXMLNamespace(FedExVersion)
		end if
		
		fedex_postdataWS=""
		fedex_postdataWS=fedex_postdataWS&"<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns" & fedex_xmlNamespace & "=""http://fedex.com/ws/ship/v" & FedExVersion & """>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<soapenv:Body>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&NameOfMethod&">"&vbcrlf

		fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"WebAuthenticationDetail>"&vbcrlf
		If CSPTurnOn = 1 Then
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"CspCredential>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"Key>" & pcv_strCSPKey & "</"&fedex_xmlPrefix&"Key>"&vbcrlf
				fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"Password>" & pcv_strCSPPassword & "</"&fedex_xmlPrefix&"Password>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"</"&fedex_xmlPrefix&"CspCredential>"&vbcrlf
		End If
		fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"UserCredential>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"Key>" & FedExkey & "</"&fedex_xmlPrefix&"Key>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"Password>" & FedExPassword & "</"&fedex_xmlPrefix&"Password>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</"&fedex_xmlPrefix&"UserCredential>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</"&fedex_xmlPrefix&"WebAuthenticationDetail>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"ClientDetail>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"AccountNumber>"&FedExAccountNumber&"</"&fedex_xmlPrefix&"AccountNumber>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"MeterNumber>"&FedExMeterNumber&"</"&fedex_xmlPrefix&"MeterNumber>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"ClientProductId>" & pcv_strClientProductID & "</"&fedex_xmlPrefix&"ClientProductId>"&vbcrlf
			fedex_postdataWS=fedex_postdataWS&"<"&fedex_xmlPrefix&"ClientProductVersion>" & pcv_strClientProductVersion & "</"&fedex_xmlPrefix&"ClientProductVersion>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</"&fedex_xmlPrefix&"ClientDetail>"&vbcrlf
	End Sub

	Public Sub NewXMLLabel(TrackingNumber, EncodedLabelString, FileType, FilePreFix)
		GraphicXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName="""&FilePreFix&TrackingNumber&"."&FileType&""">"&EncodedLabelString&"</Base64Data>"
	End Sub

	Public Sub SaveBinaryLabel ()
		objFedExStream.Type = 1
		objFedExStream.Open

		objFedExStream.Write objFEDEXXmlDoc.selectSingleNode("/Base64Data").nodeTypedValue
			err.clear
		strFileName = objFEDEXXmlDoc.selectSingleNode("/Base64Data/@FileName").nodeTypedValue
		'Save the binary stream to the file and overwrite if it already exists in folder
		objFedExStream.SaveToFile server.MapPath("FedExLabels\"&strFileName),2
		objFedExStream.Close()
	End Sub

	Public Sub EndXMLTransactionAlt(NameOfMethod, FedExVersion)
		fedex_xmlPrefix = "v" & FedExVersion & ":"
		
		fedex_postdataWS=fedex_postdataWS&"</"&fedex_xmlPrefix&NameOfMethod&">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</soapenv:Body>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</soapenv:Envelope>"&vbcrlf
	End Sub

	Public Sub EndXMLTransaction(NameOfMethod)
		fedex_postdataWS=fedex_postdataWS&"</"&NameOfMethod&">"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</soapenv:Body>"&vbcrlf
		fedex_postdataWS=fedex_postdataWS&"</soapenv:Envelope>"&vbcrlf
	End Sub

	Public Sub SendXMLRequest(XMLstring)
		srvFEDEXWSXmlHttp.open "POST", FedExWSURL, false
		srvFEDEXWSXmlHttp.send(XMLstring)
		FEDEXWS_result = srvFEDEXWSXmlHttp.responseText
	End Sub

	Public Sub SendXMLShipRequest(XMLstring)
		srvFEDEXWSXmlHttp.open "POST", FedExWSURL&"/ship", false
		srvFEDEXWSXmlHttp.send(XMLstring)
		FEDEXWS_result = srvFEDEXWSXmlHttp.responseText
	End Sub

	Public Sub SendXMLCloseRequest(XMLstring)
		srvFEDEXWSXmlHttp.open "POST", FedExWSURL&"/close", false
		srvFEDEXWSXmlHttp.send(XMLstring)
		FEDEXWS_result = srvFEDEXWSXmlHttp.responseText
	End Sub

	Public Sub LoadXMLResults(FEDEXWS_result)
		objOutputXMLDocWS.loadXML FEDEXWS_result
	End Sub


	Public Sub LoadXMLLabel(FEDEXWS_result)
		objFEDEXXmlDoc.loadXML FEDEXWS_result
	End Sub


	Public Sub XMLResponseVerify(ErrPageName)
		on error resume next
		pcv_strErrorMsgWS = ReadResponseNode("//Error", "Message")
		pcv_strErrorCodeReturn = ReadResponseNode("//Error", "Code")
		if len(pcv_strErrorMsgWS)>0 then
		end if
	End Sub

	Public Sub XMLResponseVerifyCustom(ErrPageName)
		pcv_strErrorMsgWS = ReadResponseNode("//v9:ProcessShipmentReply", "v9:Notifications/v9:Message")
	End Sub

  Private Function GetNodeName(NameOfNode)
    GetNodeName = Replace(NameOfNode, "<VER>", fedex_xmlPrefix)
  End Function

	Public Function ReadResponseNode(NameOfNode, ValueOfNode)
		on error resume next
		Set Nodes = objOutputXMLDocWS.selectNodes(GetNodeName(NameOfNode))
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(GetNodeName(ValueOfNode)).Text
		Next
		ReadResponseNode = pcv_strTempValue
	End Function

	Public Function ReadResponseNodeIdx(NameOfNode, ValueOfNode, IndexOfNode)
		on error resume next
		Set Nodes = objOutputXMLDocWS.selectNodes(GetNodeName(NameOfNode))		
		idx = 0
		For Each Node In Nodes			
			if idx = IndexOfNode then
				pcv_strTempValue=Node.selectSingleNode(GetNodeName(ValueOfNode)).Text
			end if
			
			idx = idx + 1
		Next
		ReadResponseNodeIdx = pcv_strTempValue
	End Function

	Public Function ReadResponseParent(NameOfNode, ValueOfNode)
		on error resume next
		Set Nodes = objOutputXMLDocWS.selectNodes("//"&GetNodeName(NameOfNode))
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(GetNodeName(ValueOfNode)).Text
		Next
		ReadResponseParent = pcv_strTempValue
	End Function


	Public Function ReadResponsesArray(NameOfNode, ValueOfNode)
		on error resume next
		Set Nodes = objOutputXMLDocWS.selectNodes(GetNodeName(NameOfNode))
		For Each Node In Nodes
			' Read last available sub-node
			Set subNodes = Node.selectNodes(GetNodeName(ValueOfNode))
			For Each subNode In subNodes
				pcv_strTempValue = subNode.Text
			Next

			if pcv_strTempValue="" then
				pcv_strTempValue=" "
			end if
			arryFedExTmp=arryFedExTmp&pcv_strTempValue&","
		Next
		ReadResponsesArray = arryFedExTmp
	End Function


	Public Function pcf_FedExEnabled()
		on error resume next
		pcf_FedExEnabled=false
		query="SELECT ShipmentTypes.active FROM ShipmentTypes WHERE (((ShipmentTypes.idShipment)=9));"
		set rsTmp=server.CreateObject("ADODB.RecordSet")
		set rsTmp=conntemp.execute(query)
		if NOT rsTmp.eof then
			dim FedEX_active
			FedEX_active=rsTmp("active")
			if FedEX_active=true or FedEx_active="-1" then
				pcf_FedExEnabled=true
			end if
		end if
		set rsTmp=nothing
	End Function


	Public Function pcf_FedExPackages(ido)
		on error resume next
		pcf_FedExPackages=false
		query = 		"SELECT pcPackageInfo.idOrder "
		query = query & "FROM pcPackageInfo "
		query = query & "WHERE pcPackageInfo.idOrder=" & ido &" "
		set rsTmp=server.CreateObject("ADODB.RecordSet")
		set rsTmp=conntemp.execute(query)
		if NOT rsTmp.eof then
			FedEX_idOrder=rsTmp("idOrder")
			if FedEX_idOrder=cint(ido) then
				pcf_FedExPackages=true
			end if
		end if
		set rsTmp=nothing
	End Function



	Public Function pcf_FedExSPOD(ido)
		on error resume next
		pcf_FedExSPOD=false
		query = 		"SELECT pcPackageInfo.pcPackageInfo_FDXSPODFlag "
		query = query & "FROM pcPackageInfo "
		query = query & "WHERE pcPackageInfo.pcPackageInfo_ID=" & ido &" "
		set rsTmp=server.CreateObject("ADODB.RecordSet")
		set rsTmp=conntemp.execute(query)
		if NOT rsTmp.eof then
			if rsTmp("pcPackageInfo_FDXSPODFlag") = 1 then
				pcf_FedExSPOD=true
			end if
		end if
		set rsTmp=nothing
	End Function



	Public Function pcf_FedExTrimArray(tmpArray)
		on error resume next
		'// Trim the last comma if there is one
		xStringLength = len(tmpArray)
		if xStringLength>0 then
			pcf_FedExTrimArray = left(tmpArray,(xStringLength-1))
		end if
	End Function


	Public Function pcf_FedExDateFormat(FedExDate)
		on error resume next
		FedExDay=Day(FedExDate)
		FedExMonth=Month(FedExDate)
		FedExYear= Year(FedExDate)
		pcf_FedExDateFormat=FedExYear&"-"&Right(Cstr(FedExMonth + 100),2)&"-"&Right(Cstr(FedExDay + 100),2)
	End Function


	Public Sub pcs_LogTransaction(FedExData, LogFileName, LoggingEnabled)
		on error resume next
		Dim PageName, findit, fs, f
		Set fs=server.CreateObject("Scripting.FileSystemObject")
		If LoggingEnabled = true Then
			Err.number=0
			
			LogFilePath = "../includes/FedExLogs/"
			If Not fs.FolderExists(Server.MapPath(LogFilePath)) Then
				fs.CreateFolder(Server.MapPath(LogFilePath))
			End If

			findit=Server.MapPath(LogFilePath&LogFileName)
			if (fs.FileExists(findit))=True OR (fs.FileExists(findit))="True" then
				Set f=fs.GetFile(findit)
				if Err.number=0 then
					f.Delete
				end if
			end if

			if Err.number=0 then
				Set f=fs.OpenTextFile(findit, 8, True)
				f.Write FedExData
				f.Close
			end if

		End If
		Set fs=nothing
		Set f=nothing
	End Sub

	Function RandomNumber(intHighestNumber)
		Randomize
		RandomNumber = Int(Rnd * intHighestNumber) + 1
	End Function

end class
'/////////////////////////////////////
'// End building the class here
'/////////////////////////////////////

pcf_FedExWriteLegalDisclaimers = "FedEx service marks are owned by Federal Express Corporation and are used by permission."


%>