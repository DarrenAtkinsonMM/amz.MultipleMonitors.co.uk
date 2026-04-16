
<%

const API_PAYMENT_PREFIX = "PAYMENTREQUEST"		'Only used in PayPal Express Checkout

Dim API_ENDPOINT, API_HEADER, API_VERSION, objPayPalHttp, nvpstr
Dim pcPay_PayPal_TransType, PaymentAction, pcPay_PayPal_Username, pcPay_PayPal_Password, pcPay_PayPal_Sandbox, pcPay_PayPal_Method, pcPay_PayPal_Signature, pcPay_PayPal_Currency, pcPay_PayPal_CVC, logoURL, pcPay_PayPal_Subject, pcPay_PayPal_CardTypes
Dim pcPay_PayPal_PaymentPrefix, pcPay_PayPal_PaymentIndex
Dim DeclinedString
Dim pErrNumber, pErrDescription, pErrSource, pErrSeverityCode
Dim pcv_strShippingFullName, pcv_strShippingCompany, pcv_strShippingAddress, pcv_strShippingPostalCode, pcv_strShippingStateCode, pcv_strShippingProvince, pcv_strShippingPhone, pcv_strShippingCity, pcv_strShippingCountryCode, pcv_strShippingAddress2

'/////////////////////////////////////
'// Start building the class here
'/////////////////////////////////////
Class pcPayPalClass

	private sub Class_Initialize() 
		On Error Resume Next
		API_HEADER= "text/namevalue"
		API_VERSION= "119.0"
		Set objPayPalHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
		If Err.Number<>0 Then
			Err.Number=0
		End If		
	end sub 
	
	private sub Class_Terminate()		
		Set objPayPalHttp = nothing
	end sub 

	
	'----------------------------------------------------------------------------------
	' Purpose: Generates the PayPal "prefix" string used for Express Checkout payments.
	'          The pcPay_PayPal_PaymentPrefix and pcPay_PayPal_PaymentIndex variables 
	'          must be set for this method to return the proper string.
	' Inputs:  None
	' Returns: The string with the prefix in this format: "<prefix>_<payment index>_"
	'----------------------------------------------------------------------------------	
	Private Function GetPayPalPaymentString()
		str = ""
		
		If Len(pcPay_PayPal_PaymentPrefix) > 0 Then str = str & pcPay_PayPal_PaymentPrefix & "_"
		If Len(pcPay_PayPal_PaymentIndex) > 0 Then str = str & pcPay_PayPal_PaymentIndex & "_"
		
		GetPayPalPaymentString = str
	End Function
	
	'----------------------------------------------------------------------------------
	' Purpose: Make the API call to PayPal, using API signature.
	' Inputs:  Method name to be called & NVP string to be sent with the post method
	' Returns: NVP Collection object of Call Response.
	'----------------------------------------------------------------------------------	
	Public Function hash_call(methodName, nvpStr)	
		On Error Resume Next		

		pcPay_PayPal_PaymentPrefix = ""
		pcPay_PayPal_PaymentIndex = ""
		
		AddNVP "METHOD", methodName
		AddNVP "VERSION", API_VERSION
		If len(pcPay_PayPal_Username)>0 AND len(pcPay_PayPal_Password)>0 AND len(pcPay_PayPal_Signature)>0 Then
			AddNVP "USER", pcPay_PayPal_Username
			AddNVP "PWD", pcPay_PayPal_Password
			AddNVP "SIGNATURE", pcPay_PayPal_Signature
		Else
			AddNVP "SUBJECT", pcPay_PayPal_Subject
		End If
		Set Session("nvpReqArray") = deformatNVP(nvpStr)

		API_ENDPOINT = GetPayPalURL(pcPay_PayPal_Method)
		objPayPalHttp.open "POST", API_ENDPOINT, False

		'// Include this to ensure functionality in the case of an expired or invalid SSL
		objPayPalHttp.setOption(2) = (objPayPalHttp.getOption(2) - SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID)
		objPayPalHttp.setRequestHeader "Content-Type", API_HEADER
		objPayPalHttp.Send nvpStr

		If Err.Number <> 0 Then 			
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "hash_call")
			Session("nvpReqArray") =  Null
		End If

		Set Session("nvpReqArray") = deformatNVP(nvpStr)
		Set  nvpResponseCollection = deformatNVP(objPayPalHttp.responseText)
		Set  hash_call = nvpResponseCollection

		If Err.Number <> 0 Then 			
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "hash_call")
			Session("nvpReqArray") =  Null
		End If
			
	End Function

	'----------------------------------------------------------------------------------
	' Purpose: Append a new name value pair to the NVP string.
	' Inputs:  Name and Value
	' Returns: Properly Formatted String
	'----------------------------------------------------------------------------------
	Public Sub AddNVP(pName, pValue)
		On Error Resume Next
		
		nvpstr = nvpstr & "&" & GetPayPalPaymentString() & Server.URLEncode(pName)& "=" & Server.URLEncode(pValue)
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "AddNVP")
		End If	
	End Sub
	
	'----------------------------------------------------------------------------------
	' Purpose: Append a new name value pair to the NVP "Line Items" string.
	' Inputs:  Name and Value
	' Returns: Properly Formatted String
	'----------------------------------------------------------------------------------
	Public Sub AddNVPLineItem(pName, pvalue)
		On Error Resume Next
		
		nvpstr = nvpstr & "&L_" & GetPayPalPaymentString() & Server.URLEncode(pName)& "=" & Server.URLEncode(pValue)
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "AddNVPLineItem")
		End If
	End Sub	


	'----------------------------------------------------------------------------------
	' Purpose: It will convert nvp string to Collection object.
	' Inputs:  NVP string.
	' Returns: NVP Collection object deformated from NVP string.
	'----------------------------------------------------------------------------------
	Public Function deformatNVP(nvpstr)
		On Error Resume Next
		
		Dim AndSplitedArray, EqualtoSplitedArray, Index1, Index2, NextIndex
		Set NvpCollection = Server.CreateObject("Scripting.Dictionary")
		AndSplitedArray = Split(nvpstr, "&", -1, 1)
		NextIndex=0
		For Index1 = 0 To UBound(AndSplitedArray)
			EqualtoSplitedArray=Split(AndSplitedArray(Index1), "=", -1, 1)
			For Index2 = 0 To UBound(EqualtoSplitedArray)
				NextIndex=Index2+1
				NvpCollection.Add URLDecode(EqualtoSplitedArray(Index2)),URLDecode(EqualtoSplitedArray(NextIndex))
				Index2=Index2+1
			Next
		Next
		Set deformatNVP = NvpCollection
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "deformatNVP")
		End If
		
	End Function




	'----------------------------------------------------------------------------------
	' Purpose: It gives out decoded url path to dispaly.
	' Inputs:  Url string to be decoded.
	' Returns: Decoded Url string.
	'----------------------------------------------------------------------------------
	Function URLDecode(str) 
		On Error Resume Next
		
		str = Replace(str, "+", " ")		
		For i = 1 To Len(str) 
		sT = Mid(str, i, 1) 
			If sT = "%" Then 				
				'If i+2 < Len(str) Then 					
					sR = sR & Chr(CLng("&H" & Mid(str, i+1, 2))) 
					i = i+2 
				'End If 
			Else 
				sR = sR & sT 
			End If 
		Next 				   
		URLDecode = sR 
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "URLDecode")
		End If
		
	End Function




	'----------------------------------------------------------------------------------
	' Purpose: It's Workaround Method for Response.Redirect
	'          It will redirect the page to the specified url without urlencoding
	' Inputs: Url to redirect the page
	'----------------------------------------------------------------------------------
	Function ReDirectURL(url)	
		On Error Resume Next
		
		'// PayPal recommends 302 redirects, however, "AddHeader" doesnt work with certain server configurations.	
		'response.clear
		'response.status="302 Object moved"
		'response.AddHeader "location",url
		
		'// Use Redirect
		Response.Redirect(url)
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "ReDirectURL")
		End If
		
	End Function




	'----------------------------------------------------------------------------------
	' Purpose: It will Format error Messages into a HTML string
	' Inputs:  NVP string.
	' Returns: NVP Collection object deformated from NVP string.
	'----------------------------------------------------------------------------------
	Function ErrorFormatter(errDesc, errNumber, errSource, errlocation)
		
		'// API Errors need filtered here. Use Select Case to append User Friendly Information.
		Select Case errNumber
			Case "10002":errDesc=errDesc&". This error means that your API Credentials are not correct for the ""Live"", or ""Sandbox"", modes. Please double check you are in the correct mode and your credentials are valid."	
			Case"10736":errDesc="The payment system was not able to validate your address. Please review it again to ensure that everything is spelled correctly, and that the postal code is a valid one. If you are paying with PayPal, please make sure the address matches the one stored in your PayPal account.<hr/>"	
		End Select
		
		ErrorFormatter = "<div align=""left"">" & _
		"<ul>" &_
		"<li>" & "Error Source: " & errSource & "</li>"
		ErrorFormatter = ErrorFormatter & "</ul></div>"
		
		If Err.Number <> 0 Then
			Err.Clear
		End If
	End Function 

	'----------------------------------------------------------------------------------
	' Purpose: Append Our HTML error strings into one report.
	' Inputs:  pcv_PayPalErrMessage, DeclinedString
	' Returns: pcv_PayPalErrMessage + DeclinedString as one formatted string.
	'----------------------------------------------------------------------------------
	Public Sub GenerateErrorReport()
		On Error Resume Next

		pErrNumber = resArray("L_ERRORCODE0")
		pErrDescription = resArray("L_SHORTMESSAGE0")
		pErrSource = resArray("L_LONGMESSAGE0")
		pErrSeverityCode = resArray("L_SEVERITYCODE0")

		If pErrDescription <> "" Then 
			pcv_PayPalErrMessage = pcv_PayPalErrMessage & objPayPalClass.ErrorFormatter(pErrDescription, pErrNumber, pErrSource, "PayPal Service")
		End If
		
		if DeclinedString<>"" then
			pcv_PayPalErrMessage=pcv_PayPalErrMessage & "<hr/><div>API Errors</div><hr/>"		
			pcv_PayPalErrMessage=pcv_PayPalErrMessage & "<div>" & DeclinedString & "</div>"
			pcv_PayPalErrMessage=pcv_PayPalErrMessage & "<hr/>"
		end if		
	End Sub

	'----------------------------------------------------------------------------------
	' Purpose: It gives url path for the cancel & return  page.
	' Returns: Url path of current page without file name.
	'----------------------------------------------------------------------------------
	Public Function GetURL() 
		On Error Resume Next		
		
		if scSSL = "1" then
			Virtual_Path = scSslURL &"/"& scPcFolder & "/pc/"
		else
			Virtual_Path = scStoreURL &"/"& scPcFolder & "/pc/"
		end if

		GetURL = Virtual_Path
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "GetURL")
		End If
		
	End Function




	'----------------------------------------------------------------------------------
	' Purpose: It gives url path for the cancel & return  page.
	' Returns: Url path of current page without file name.
	'----------------------------------------------------------------------------------
	'Mobile-S
	Public Function GetURLMobile() 
		On Error Resume Next		
		
		if scSSL = "1" then
			Virtual_Path = scSslURL &"/"& scPcFolder & "/m/"
		else
			Virtual_Path = scStoreURL &"/"& scPcFolder & "/m/"
		end if

		Virtual_Path = replace(Virtual_Path,"://",":///")
		Virtual_Path = replace(Virtual_Path,"//","/")
		GetURLMobile = Virtual_Path

		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "GetURLMobile")
		End If
		
	End Function
	'Mobile-E




	'----------------------------------------------------------------------------------
	' Purpose: It gives url to the PayPal server.
	' Inputs:  PayPal method "sandbox" or "live"
	' Returns: Sandbox or Live Server URL
	'----------------------------------------------------------------------------------
	Public Function GetPayPalURL(pcPay_PayPal_Method)		
		if pcPay_PayPal_Method = "sandbox" then
			GetPayPalURL = "https://api-3t.sandbox.paypal.com/nvp"
		else
			GetPayPalURL = "https://api-3t.paypal.com/nvp"
		end if
	End Function

	Public Function GetECURL(pcPay_PayPal_Method)
		if scPaypalECInContext = "1" then
			subURL = "checkoutnow"
		else
			subURL = "webscr"
		end if
		
		if pcPay_PayPal_Method = "sandbox" then
			GetECURL = "https://www.sandbox.paypal.com/" & subURL
		else
			GetECURL = "https://www.paypal.com/" & subURL
		end if
	End Function


	'----------------------------------------------------------------------------------
	' Purpose: Provides a clean way to set all the PayPal variables to local.
	' Inputs:  None. Requires an open database connection
	' Returns: pcPay_PayPal_TransType, PaymentAction, pcPay_PayPal_Username, pcPay_PayPal_Password, pcPay_PayPal_Sandbox, pcPay_PayPal_Method, pcPay_PayPal_Signature, pcPay_PayPal_Subject
	'----------------------------------------------------------------------------------	
	Public Sub pcs_SetAllVariables()
		On Error Resume Next
		Dim pcMissingField
		pcMissingField = 0
		
		query = "SELECT pcPay_PayPal.pcPay_PayPal_CardTypes FROM pcPay_PayPal;"
		set rsPayPalVar=server.CreateObject("ADODB.RecordSet")
		set rsPayPalVar=conntemp.execute(query)
		if err.number<>0 then
			pcMissingField = 1
			'// Query PayPal Table
			query="SELECT pcPay_PayPal.pcPay_PayPal_TransType, pcPay_PayPal.pcPay_PayPal_Subject, pcPay_PayPal.pcPay_PayPal_Username, pcPay_PayPal.pcPay_PayPal_Password, pcPay_PayPal.pcPay_PayPal_AVS, pcPay_PayPal.pcPay_PayPal_CVC, pcPay_PayPal.pcPay_PayPal_Sandbox, pcPay_PayPal.pcPay_PayPal_Signature, pcPay_PayPal.pcPay_PayPal_Currency, pcPay_PayPal.pcPay_PayPal_CVC FROM pcPay_PayPal WHERE (((pcPay_PayPal.pcPay_PayPal_ID)=1));"
			set rsPayPalVar=server.CreateObject("ADODB.RecordSet")
			set rsPayPalVar=conntemp.execute(query)
		else		
			'// Query PayPal Table
			query="SELECT pcPay_PayPal.pcPay_PayPal_TransType, pcPay_PayPal.pcPay_PayPal_Subject, pcPay_PayPal.pcPay_PayPal_Username, pcPay_PayPal.pcPay_PayPal_Password, pcPay_PayPal.pcPay_PayPal_AVS, pcPay_PayPal.pcPay_PayPal_CVC, pcPay_PayPal.pcPay_PayPal_Sandbox, pcPay_PayPal.pcPay_PayPal_Signature, pcPay_PayPal.pcPay_PayPal_Currency, pcPay_PayPal.pcPay_PayPal_CVC, pcPay_PayPal.pcPay_PayPal_CardTypes FROM pcPay_PayPal WHERE (((pcPay_PayPal.pcPay_PayPal_ID)=1));"
			set rsPayPalVar=server.CreateObject("ADODB.RecordSet")
			set rsPayPalVar=conntemp.execute(query)
		end if
		
		'// Set Local Var
		pcPay_PayPal_TransType=rsPayPalVar("pcPay_PayPal_TransType")		
		pcPay_PayPal_Subject=trim(rsPayPalVar("pcPay_PayPal_Subject"))
		pcPay_PayPal_Username=trim(rsPayPalVar("pcPay_PayPal_Username"))
		pcPay_PayPal_Password=trim(rsPayPalVar("pcPay_PayPal_Password"))
		pcPay_PayPal_CVC = rsPayPalVar("pcPay_PayPal_CVC")
		pcPay_PayPal_Sandbox=rsPayPalVar("pcPay_PayPal_Sandbox")
		pcPay_PayPal_Signature = trim(rsPayPalVar("pcPay_PayPal_Signature"))
		pcPay_PayPal_Currency = rsPayPalVar("pcPay_PayPal_Currency")
		If pcMissingField = 1 Then
			pcPay_PayPal_CardTypes = scMobilePayPalCardTypes
		Else
			pcPay_PayPal_CardTypes = rsPayPalVar("pcPay_PayPal_CardTypes")		
		End If
		
		' Check pcPay_PayPal_Currency for NULL
		if isNULL(pcPay_PayPal_Currency)=True or pcPay_PayPal_Currency="" then
			pcPay_PayPal_Currency="USD"
		end if
		
		' Check pcPay_PayPal_CVC for NULL
		if isNULL(pcPay_PayPal_CVC)=True or pcPay_PayPal_CVC="" then
			pcPay_PayPal_CVC=1
		end if
		
		' Check pcPay_PayPal_CardTypes for NULL
		if isNULL(pcPay_PayPal_CardTypes)=True or pcPay_PayPal_CardTypes="" then
			pcPay_PayPal_CardTypes="V, M, D"
		end if
		
		' Authorize or Capture
		if pcPay_PayPal_TransType="1" then
			PaymentAction="Sale"	
		else
			PaymentAction="Authorization"
		end if
		
		' Sandbox or Live
		if pcPay_PayPal_Sandbox=1 then
			pcPay_PayPal_Method = "sandbox"
		else
			pcPay_PayPal_Method = "live"
		end if
		
		'// Close our Db connections
		set rsPayPalVar=nothing
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "pcs_SetAllVariables")
		End If
		
	End Sub
	
	
	'----------------------------------------------------------------------------------
	' Purpose: Provides a clean way to obtain the latest Address.
	'----------------------------------------------------------------------------------	
	Public Sub pcs_SetShipAddress(OrderID)
		On Error Resume Next
		
		'// Query PayPal Table
		query="SELECT ShippingFullName, shippingCompany, shippingAddress, shippingZip, shippingStateCode, shippingState, pcOrd_shippingPhone, shippingCity, shippingCountryCode, shippingAddress2 FROM orders WHERE idorder="&OrderID&";"
		set rsPayPalVar=server.CreateObject("ADODB.RecordSet")
		set rsPayPalVar=conntemp.execute(query)

		'// Set Local Var
		pcv_strShippingFullName=rsPayPalVar("ShippingFullName")	
		pcv_strShippingCompany=rsPayPalVar("shippingCompany")
		pcv_strShippingAddress=rsPayPalVar("shippingAddress")
		pcv_strShippingPostalCode=rsPayPalVar("shippingZip")
		pcv_strShippingStateCode=rsPayPalVar("shippingStateCode")
		pcv_strShippingProvince=rsPayPalVar("shippingState")
		pcv_strShippingPhone=rsPayPalVar("pcOrd_shippingPhone")
		pcv_strShippingCity=rsPayPalVar("shippingCity")
		pcv_strShippingCountryCode=rsPayPalVar("shippingCountryCode")
		pcv_strShippingAddress2=rsPayPalVar("shippingAddress2")						
		
		'// Close our Db connections
		set rsPayPalVar=nothing
		
		If Err.Number <> 0 Then 
			DeclinedString = DeclinedString & ErrorFormatter(Err.Description, Err.Number, Err.Source, "pcs_SetShipAddress")
		End If
		
	End Sub

end class 
'/////////////////////////////////////
'// End building the class here
'/////////////////////////////////////


'// Format For Field
Public Function pcf_CurrencyField(moneyAMT)	
	if scDecSign = "," then
		moneyAMT=replace(moneyAMT,".","")
		moneyAMT=replace(moneyAMT,",",".")		
	else
		moneyAMT=replace(moneyAMT,",","")
	end if
	'// Convert to proper form for PayPal:
	'// Param 1: The monetary amount to format
	'// Param 2: Number of digits after decimal = 2
	'// Param 3: Include leading 0 before decimal value = True (-1)
	'// Param 4: Place negative values in parenthesis = False (0)
	'// Param 5: Include group delimeter (commas, etc) = False (0)
	pcf_CurrencyField=FormatNumber(moneyAMT,2,-1,0,0)
End Function
%>