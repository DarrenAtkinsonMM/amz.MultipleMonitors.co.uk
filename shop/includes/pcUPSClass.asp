
<%
'/////////////////////////////////////
'// Start building the class here
'/////////////////////////////////////
Class pcUPSClass 

	private sub Class_Initialize() 
		'// define all parameter will use
		Set srvUPSXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
		Set objOutputXMLDoc = Server.CreateObject("Microsoft.XMLDOM")
		Set objUPSStream = Server.CreateObject("ADODB.Stream")
		Set objUPSXmlDoc = server.createobject("Msxml2.DOMDocument"&scXML)
		objUPSXmlDoc.async = False
		objUPSXmlDoc.validateOnParse = False
	end sub 
	
	
	
	private sub Class_Terminate() 
		'// clean it all up
		Set srvUPSXmlHttp = nothing
		Set objOutputXMLDoc = nothing
		Set objUPSXmlDoc = nothing
		Set objUPSStream = nothing
	end sub 
	
	
	
	Public Sub AddNewNode(NameOfNode, ValueOfNode)
		if len(ValueOfNode)>0 then
			ups_postdata=ups_postdata&"<"&NameOfNode&">"&ValueOfNode&"</"&NameOfNode&">"&vbcrlf
		end if
	End Sub
	
	
	
	Public Sub WriteParent(NameOfParent, isClosing)
		ups_postdata=ups_postdata&"<"&isClosing&""&NameOfParent&">"&vbcrlf
	End Sub
	
	
	
	Public Sub WriteSingleParent(NameOfParent, ValueOfParent)
		if len(ValueOfParent)>0 then
			ups_postdata=ups_postdata&"<"&NameOfParent&">"&ValueOfParent&"</"&NameOfParent&">"&vbcrlf
		end if
	End Sub
	
	Public Sub NewXMLTransaction(UPS_AccessLicenseNumber, UPS_UserID, UPS_Password)
		ups_accessrequest=""
		ups_accessrequest=ups_accessrequest&"<?xml version=""1.0""?>"&vbcrlf
		ups_accessrequest=ups_accessrequest&"<AccessRequest xml:lang=""en-US"">"&vbcrlf
		ups_accessrequest=ups_accessrequest&"<AccessLicenseNumber>"&UPS_AccessLicenseNumber&"</AccessLicenseNumber>"&vbcrlf
		ups_accessrequest=ups_accessrequest&"<UserId>"&UPS_UserID&"</UserId>"&vbcrlf
		ups_accessrequest=ups_accessrequest&"<Password>"&UPS_Password&"</Password>"&vbcrlf
		ups_accessrequest=ups_accessrequest&"</AccessRequest>"&vbcrlf		
	End Sub
	
	Public Sub NewXMLShipmentConfirmRequest(NameOfMethod, UPS_AV)
		ups_postdata=ups_postdata&"<?xml version=""1.0""?>"&vbcrlf
		ups_postdata=ups_postdata&"<ShipmentConfirmRequest xml:lang=""en-US"">"&vbcrlf
		ups_postdata=ups_postdata&"<Request>"&vbcrlf
		ups_postdata=ups_postdata&"<TransactionReference>"&vbcrlf
		ups_postdata=ups_postdata&"<CustomerContext>ShipConfirmUS</CustomerContext>"&vbcrlf
		ups_postdata=ups_postdata&"<XpciVersion>1.0001</XpciVersion>"&vbcrlf
		ups_postdata=ups_postdata&"</TransactionReference>"&vbcrlf
		ups_postdata=ups_postdata&"<RequestAction>"&NameOfMethod&"</RequestAction>"&vbcrlf
		ups_postdata=ups_postdata&"<RequestOption>"&UPS_AV&"</RequestOption>"&vbcrlf
		ups_postdata=ups_postdata&"</Request>"&vbcrlf
	End Sub

	Public Sub NewXMLShipmentAcceptRequest(CustomerContext, ShipmentDigest)
		ups_postdata=ups_postdata&"<?xml version=""1.0""?>"&vbcrlf
		ups_postdata=ups_postdata&"<ShipmentAcceptRequest>"&vbcrlf
		ups_postdata=ups_postdata&"<Request>"&vbcrlf
		ups_postdata=ups_postdata&"<TransactionReference>"&vbcrlf
		ups_postdata=ups_postdata&"<CustomerContext>"&CustomerContext&"</CustomerContext>"&vbcrlf
		ups_postdata=ups_postdata&"<XpciVersion>1.0001</XpciVersion>"&vbcrlf
		ups_postdata=ups_postdata&"</TransactionReference>"&vbcrlf
		ups_postdata=ups_postdata&"<RequestAction>ShipAccept</RequestAction>"&vbcrlf
		ups_postdata=ups_postdata&"<RequestOption>01</RequestOption>"&vbcrlf
		ups_postdata=ups_postdata&"</Request>"&vbcrlf
		ups_postdata=ups_postdata&"<ShipmentDigest>"&ShipmentDigest&"</ShipmentDigest>"&vbcrlf
		ups_postdata=ups_postdata&"</ShipmentAcceptRequest>"&vbcrlf
	End Sub
	
	Public Sub NewXMLShipmentVoidRequest(CustomerContext, ShipmentIDNum)
		ups_postdata=ups_postdata&"<?xml version=""1.0""?>"&vbcrlf
		ups_postdata=ups_postdata&"<VoidShipmentRequest>"&vbcrlf
		ups_postdata=ups_postdata&"<Request>"&vbcrlf
		ups_postdata=ups_postdata&"<TransactionReference>"&vbcrlf
		ups_postdata=ups_postdata&"<CustomerContext>"&CustomerContext&"</CustomerContext>"&vbcrlf
		ups_postdata=ups_postdata&"<XpciVersion>1.0001</XpciVersion>"&vbcrlf
		ups_postdata=ups_postdata&"</TransactionReference>"&vbcrlf
		ups_postdata=ups_postdata&"<RequestAction>1</RequestAction>"&vbcrlf
		ups_postdata=ups_postdata&"<RequestOption></RequestOption>"&vbcrlf
		ups_postdata=ups_postdata&"</Request>"&vbcrlf
		ups_postdata=ups_postdata&"<ShipmentIdentificationNumber>"&ShipmentIDNum&"</ShipmentIdentificationNumber>"&vbcrlf
		ups_postdata=ups_postdata&"</VoidShipmentRequest>"&vbcrlf
	End Sub
	
	Public Sub NewXMLShipmentTrackRequest(CustomerContext, TrackingIDNum)
		ups_postdata=ups_postdata&"<?xml version=""1.0""?>"
		ups_postdata=ups_postdata&"<TrackRequest xml:lang=""en-US"">"
		ups_postdata=ups_postdata&"<Request>"
		ups_postdata=ups_postdata&"<TransactionReference>"
		ups_postdata=ups_postdata&"<CustomerContext>"&CustomerContext&"</CustomerContext>"
		ups_postdata=ups_postdata&"<XpciVersion>1.0001</XpciVersion>"
		ups_postdata=ups_postdata&"</TransactionReference>"
		ups_postdata=ups_postdata&"<RequestAction>Track</RequestAction>"
		ups_postdata=ups_postdata&"<RequestOption>Activity</RequestOption>"
		ups_postdata=ups_postdata&"</Request>"
		ups_postdata=ups_postdata&"<TrackingNumber>"&TrackingIDNum&"</TrackingNumber>"
		ups_postdata=ups_postdata&"</TrackRequest>"
	End Sub
	
	Public Sub NewXMLShipmentTimeInTransitRequest(CustomerContext)
		ups_postdata=ups_postdata&"<?xml version=""1.0""?>"&vbcrlf
		ups_postdata=ups_postdata&"<TimeInTransitRequest>"&vbcrlf
		ups_postdata=ups_postdata&"<Request>"&vbcrlf
		ups_postdata=ups_postdata&"<TransactionReference>"&vbcrlf
		ups_postdata=ups_postdata&"<CustomerContext>"&CustomerContext&"</CustomerContext>"&vbcrlf
		ups_postdata=ups_postdata&"<XpciVersion>1.0002</XpciVersion>"&vbcrlf
		ups_postdata=ups_postdata&"</TransactionReference>"&vbcrlf
		ups_postdata=ups_postdata&"<RequestAction>TimeInTransit</RequestAction>"&vbcrlf
		ups_postdata=ups_postdata&"</Request>"&vbcrlf
	End Sub

	Public Sub NewXMLLabel(TrackingNumber, EncodedLabelString, FileType, FilePreFix)
		GraphicXML="<Base64Data xmlns:dt=""urn:schemas-microsoft-com:datatypes"" dt:dt=""bin.base64"" FileName="""&FilePreFix&TrackingNumber&"."&FileType&""">"&EncodedLabelString&"</Base64Data>"
	End Sub
	
	Public Sub SaveBinaryLabel ()
		objUPSStream.Type = 1
		objUPSStream.Open
		
		objUPSStream.Write objUPSXmlDoc.selectSingleNode("/Base64Data").nodeTypedValue 
			err.clear
		strFileName = objUPSXmlDoc.selectSingleNode("/Base64Data/@FileName").nodeTypedValue 
		'Save the binary stream to the file and overwrite if it already exists in folder
		objUPSStream.SaveToFile server.MapPath("UPSLabels\"&strFileName),2
		objUPSStream.Close()
	End Sub
	
	Public Sub EndXMLTransaction(NameOfMethod)
		ups_postdata=ups_postdata&"</"&NameOfMethod&">"&vbcrlf
	End Sub
	
	
	Public Sub SendXMLRequest(XMLstring, strURL)
		'// resolve, connect, send, receive - in milliseconds 
		srvUPSXmlHttp.setTimeouts 2500, 2500, 5000, 5000
		srvUPSXmlHttp.open "POST", strURL, false
		srvUPSXmlHttp.send(XMLstring)
		UPS_result = srvUPSXmlHttp.responseText	
		if err>0 then
			'// handle error
		end if
	End Sub
	
	Public Sub LoadXMLResults(UPS_result)
		objOutputXMLDoc.loadXML UPS_result
	End Sub
	
	Public Sub LoadXMLLabel(UPS_result)
		objUPSXmlDoc.loadXML UPS_result
	End Sub
		
	Public Sub XMLResponseVerify(ErrPageName)
		pcv_strErrorCode = objUPSClass.ReadResponseNode("//Error", "ErrorCode") 
		pcv_strErrorMsg = objUPSClass.ReadResponseNode("//Error", "ErrorDescription")
		pcv_strErrorSeverity = objUPSClass.ReadResponseNode("//Error", "ErrorSeverity")
		if len(pcv_strErrorMsg)>0 AND ucase(pcv_strErrorSeverity)<>"WARNING" then
			Session("ErrMsg") = "There was an error processing your request.<br>Code: "&pcv_strErrorCode&" - Error: " & pcv_strErrorMsg
			response.redirect ErrPageName & "?err=1"
		else
			pcv_strErrorMsg=""
		end if
	End Sub
	
	Public Sub XMLResponseVerifyCustom(ErrPageName)
		pcv_strErrorMsg = objUPSClass.ReadResponseNode("//Error", "ErrorDescription")
	End Sub
	
	Public Function ReadResponseNode(NameOfNode, ValueOfNode)
		on error resume next	
		Set Nodes = objOutputXMLDoc.selectNodes(NameOfNode)
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(ValueOfNode).Text  			
		Next
		ReadResponseNode = pcv_strTempValue
	End Function

	Public Function ReadTrackingNode(NameOfNode, ValueOfNode)
		intNodeCnt=0
		Set Nodes = objOutputXMLDoc.selectNodes(NameOfNode)
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(ValueOfNode).Text  
			if len(pcv_strTempValue)>1 then
				intNodeCnt=intNodeCnt+1
			end if			
		Next
		ReadTrackingNode = pcv_strTempValue
	End Function

	Public Function ReadResponseParent(NameOfNode, ValueOfNode)	
		on error resume next	
		Set Nodes = objOutputXMLDoc.selectNodes("//"&NameOfNode)	
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(ValueOfNode).Text  			
		Next
		ReadResponseParent = pcv_strTempValue
	End Function		
	
	Public Function ReadResponseasArray(NameOfNode, ValueOfNode)	
		on error resume next	
		Set Nodes = objOutputXMLDoc.selectNodes(NameOfNode)	
		For Each Node In Nodes
			pcv_strTempValue=Node.selectSingleNode(ValueOfNode).Text 
			if pcv_strTempValue="" then
				pcv_strTempValue=" "
			end if
			arryUPSTmp=arryUPSTmp&pcv_strTempValue&"," 			
		Next
		ReadResponseasArray = arryUPSTmp

	End Function
	
	Public Function pcf_UPSEnabled()	
		on error resume next
		pcf_UPSEnabled=false	
		query="SELECT ShipmentTypes.active FROM ShipmentTypes WHERE (((ShipmentTypes.idShipment)=3));"
		set rsTmp=server.CreateObject("ADODB.RecordSet")
		set rsTmp=conntemp.execute(query)
		if NOT rsTmp.eof then
			dim UPS_active
			UPS_active=rsTmp("active")
			if UPS_active=true or UPS_active="-1" then
				pcf_UPSEnabled=true
			end if
		end if 
		set rsTmp=nothing		
	End Function
	
	
	Public Function pcf_UPSPackages(ido)	
		on error resume next
		pcf_UPSPackages=false			
		query = 		"SELECT pcPackageInfo.idOrder "
		query = query & "FROM pcPackageInfo "
		query = query & "WHERE pcPackageInfo.idOrder=" & ido &" "	
		set rsTmp=server.CreateObject("ADODB.RecordSet")
		set rsTmp=conntemp.execute(query)
		if NOT rsTmp.eof then
			UPS_idOrder=rsTmp("idOrder")
			if UPS_idOrder=cint(ido) then
				pcf_UPSPackages=true
			end if
		end if 
		set rsTmp=nothing
	End Function
	
	
	Public Function pcf_UPSTrimArray(tmpArray)	
		on error resume next
		'// Trim the last comma if there is one
		xStringLength = len(tmpArray)
		if xStringLength>0 then
			pcf_UPSTrimArray = left(tmpArray,(xStringLength-1))
		end if			
	End Function
	
	
	Public Function pcf_UPSDateFormat(UPSDate)
		on error resume next
		UPSDay=Day(UPSDate)
		UPSMonth=Month(UPSDate)
		UPSYear= Year(UPSDate)
		pcf_UPSDateFormat=UPSYear&"-"&Right(Cstr(UPSMonth + 100),2)&"-"&Right(Cstr(UPSDay + 100),2)
	End Function
	
	Public Sub pcs_LogTransaction(UPSData, LogFileName, LoggingEnabled)
		Dim PageName, findit, fs, f
		Set fs=server.CreateObject("Scripting.FileSystemObject")				
		If LoggingEnabled = true Then			
			Err.number=0	
			
			findit=Server.MapPath("UPSLabels/"&LogFileName)	
			if (fs.FileExists(findit))=true then	
				Set f=fs.GetFile(findit)				
				if Err.number=0 then
					f.Delete
				end if
			end if
			
			if Err.number=0 then
				Set f=fs.OpenTextFile(findit, 2, True)
				f.Write UPSData
				f.Close
			end if
			
		End If
		Set fs=nothing
		Set f=nothing
	End Sub


end class 
'/////////////////////////////////////
'// End building the class here
'/////////////////////////////////////

pcf_UPSWriteLegalDisclaimersText = "UPS, the UPS Shield trademark, the UPS Ready mark, <br />the UPS Developer Kit mark and the Color Brown are trademarks of <br />United Parcel Service of America, Inc. All Rights Reserved."

pcf_UPSWriteLegalDisclaimers = "<table><tr><td width='58' valign='top' bgcolor='#FFFFFF'><div align='right'><img src='../UPSLicense/LOGO_S2.jpg' width='45' height='50' /></div></td><td width='457' valign='top' bgcolor='#FFFFFF'><div align='center'><br />UPS, the UPS Shield trademark, the UPS Ready mark, <br />the UPS Developer Kit mark and the Color Brown are trademarks of <br />United Parcel Service of America, Inc. All Rights Reserved.</div></td></tr></table>"
%>