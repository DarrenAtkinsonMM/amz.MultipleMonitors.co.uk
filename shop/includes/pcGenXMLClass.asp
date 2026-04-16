<%
'/////////////////////////////////////
'// Start building the class here
'/////////////////////////////////////
Class pcGenXMLClass 

	private sub Class_Initialize() 
		'// open all object that we will need
		'// define all parameter will use
		Set srvGenXMLXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
		Set objOutputXMLDoc = Server.CreateObject("Microsoft.XMLDOM")
		Set objGenXMLStream = Server.CreateObject("ADODB.Stream")
		Set objGenXMLXmlDoc = server.createobject("Msxml2.DOMDocument"&scXML)
		objGenXMLXmlDoc.async = False
		objGenXMLXmlDoc.validateOnParse = False
	end sub 
	
	
	
	private sub Class_Terminate() 
		'// clean it all up
		Set srvGenXMLXmlHttp = nothing
		Set objOutputXMLDoc = nothing
		Set objGenXMLXmlDoc = nothing
		Set objGenXMLStream = nothing
	end sub 
	

	Public Sub LoadXMLResults(GenXML_result)
		objOutputXMLDoc.loadXML GenXML_result
	End Sub
	
	Public Sub LoadXMLLabel(GenXML_result)
		objGenXMLXmlDoc.loadXML GenXML_result
	End Sub
		
	Public Sub XMLResponseVerify(ErrPageName)
	
		strErrorNumber = objGenXMLClass.ReadResponseNode("//Error", "Number") 
		strErrorSource = objGenXMLClass.ReadResponseNode("//Error", "Source")
		strErrorDescription = objGenXMLClass.ReadResponseNode("//Error", "Description")
		strErrorHelpFile = objGenXMLClass.ReadResponseNode("//Error", "HelpFile")
		strErrorHelpContext = objGenXMLClass.ReadResponseNode("//Error", "HelpContext")
		
		if len(strErrorNumber)>0 then
			response.redirect ErrPageName & "?LabelMode="&pcv_LabelMode&"&msg=There was an error processing your request.<br>" & strErrorDescription & "<br>GenXML Error Code: "&strErrorNumber
		else
			pcv_strErrorMsg=""
		end if
	End Sub
	
	Public Sub XMLResponseVerifyCustom(ErrPageName)
		pcv_strErrorMsg = objGenXMLClass.ReadResponseNode("//Error", "ErrorDescription")
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
			arryGenXMLTmp=arryGenXMLTmp&pcv_strTempValue&"," 			
		Next
		ReadResponseasArray = arryGenXMLTmp

	End Function
	

	Public Function pcf_GenXMLTrimArray(tmpArray)	
		on error resume next
		'// Trim the last comma if there is one
		xStringLength = len(tmpArray)
		if xStringLength>0 then
			pcf_GenXMLTrimArray = left(tmpArray,(xStringLength-1))
		end if			
	End Function
	
	

end class 
'/////////////////////////////////////
'// End building the class here
'/////////////////////////////////////

pcf_USPSWriteLegalDisclaimersText = "USPS, THE USPS SHIELD TRADEMARK, THE USPS READY MARK, <br />THE USPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF<br />UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED."

pcf_USPSWriteLegalDisclaimers = "<table><tr><td width='58' valign='top' bgcolor='#FFFFFF'><div align='right'><img src='../USPSLicense/LOGO_S2.jpg' width='45' height='50' /></div></td><td width='457' valign='top' bgcolor='#FFFFFF'><div align='center'><br />USPS, THE USPS SHIELD TRADEMARK, THE USPS READY MARK, <br />THE USPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF<br />UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</div></td></tr></table>"
%>