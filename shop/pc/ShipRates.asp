<!--#include file="../includes/dimensionsformatinc.asp"-->
<% server.ScriptTimeout = 300 %>
<% on error resume next

dim resolveTimeout, connectTimeout, sendTimeout, receiveTimeout

resolveTimeout	= 5000
connectTimeout	= 5000
sendTimeout		= 5000
receiveTimeout	= 10000
'1000ms = 1 sec

'UPS DEBUGGING VARIABLES
'******************************************************************
'// Debug UPS post and reply
'// To turn UPS debugging on, change the value of pcv_UPSDebug=1
'******************************************************************
pcv_UPSDebug=0

'******************************************************************
'// Log UPS reply
'// To turn UPS logging on, change the value of pcv_UPS_Logging=1
'******************************************************************
pcv_UPS_Logging=0

'UPS CANADA ONLY VARIABLES
'******************************************************************
'// Use Canada as the Ship From Origin
'// To set Canada as the Ship From Origin,
'// change the value of pcv_UPSCanadaOrigin=1
'// You MUST also run the Script "upddbUPSShipOrigin.asp" that
'// is located in your ProductCart Control Panel Folder
'******************************************************************
dim pcv_UPSCanadaOrigin
pcv_UPSCanadaOrigin=0


'U.S.P.S. OPTIONAL VARIABLES
'******************************************************************
'// USPS Value of Content for International Rates Only
'// If specified, it is used to compute Insurance fee
'// (if insurance is available for service and destination) and
'// indemnity coverage.
'// To turn this variable on, change the value to "1"
'//
'// For Example:
'// pcv_UseValueOfContents=1

'******************************************************************
pcv_UseValueOfContents=1
'******************************************************************
if pcv_UseValueOfContents=1 then
	pcv_ValueOfContents=pSubTotal
end if

'Set variables from Constants UPS
pcv_UseNegotiatedRates=UPS_USENEGOTIATEDRATES
pcv_UPSShipperNumber=UPS_SHIPPERNUM
pcv_InsuredValue=UPS_INSUREDVALUE
pcv_UseDynamicInsuredValue=UPS_DYNAMICINSUREDVALUE

'Set variables from Contants USPS
pcv_InsuredValue_USPS=USPS_INSUREDVALUE
pcv_UseDynamicInsuredValue_USPS=USPS_DYNAMICINSUREDVALUE

'Set variables from Constants FEDEX WS
pcv_InsuredValue_FDXWS=FEDEXWS_INSUREDVALUE '// WS
pcv_UseDynamicInsuredValue_FDXWS=FEDEXWS_DYNAMICINSUREDVALUE '// WS

if pcv_UseDynamicInsuredValue="1" then
	pcv_InsuredValue=pSubTotal
end if

if pcv_UseDynamicInsuredValue_FDXWS="1" then
	pcv_InsuredValue_FDXWS=Cdbl(calculateShipCartTotal(pcCartArray, ppcCartIndex))
end if

if pcv_UseDynamicInsuredValue_USPS="1" then
	pcv_InsuredValue_USPS=Cdbl(calculateShipCartTotal(pcCartArray, ppcCartIndex))
end if

iFedExWSFlag=0 '// WS
iCPFlag=0
iUSPSFlag=0
iCustomFlag=0
strOptionShipmentType=""
strTabShipmentType=""

pcv_intPackageNum=0
pcv_intTotPackageNum=0

dim pcv_EMWeightLimit, pcv_PMWeightLimit,pcv_EM_Null,pcv_PM_Null

pcv_EMWeightLimit=USPS_EM_FREWeightLimit
if NOT isNumeric(pcv_EMWeightLimit) or pcv_EMWeightLimit="" then
	pcv_EMWeightLimit=0
end if
pcv_PMWeightLimit=USPS_PM_FREWeightLimit
if NOT isNumeric(pcv_PMWeightLimit) or pcv_PMWeightLimit="" then
	pcv_PMWeightLimit=0
end if
pcv_EM_Null=0
pcv_PM_Null=0


'Check if any products are labeled as oversize for UPS & FedEX & USPS
Dim pcv_intOSCheck, pcv_intOSStatus, pcv_arrOSCheckArray, pcv_arrOSArray
err.clear
err.number=0

if pcv_EOSC="" then
	pcv_intOSCheck=oversizecheck(pcCartArray, ppcCartIndex)
else
	pcv_intOSCheck=eoversizecheck(request("idOrder"))
end if

pcv_intOSStatus=0

'if products are oversize, double check to be sure values exists
if pcv_intOSCheck<>"" then
	pcv_arrOSCheckArray=split(pcv_intOSCheck,",")
	for i=0 to Ubound(pcv_arrOSCheckArray)-1
		pcv_arrOSArray=split(pcv_arrOSCheckArray(i),"|||")
		if pcv_arrOSArray(0)>pcv_intOSStatus then
			pcv_intOSStatus=1
		end if
	next
end if

Dim USPS_OverSizeOrder
USPS_OverSizeOrder=pcv_intOSStatus
dim intPackageCnt, intWeightCnt
intPackageCnt=0
intWeightCnt=0
dim pcv_intOSwidth, pcv_intOSheight, pcv_intOSlength, intOSstatus
'=====================================================================
'for each oversized package, get height, width, length and weight
'keep a running package count
'---------------------------------------------------------------------

if pcv_intOSStatus<>0 then 'There are OS packages
	'keep track of BTO/OS Items
	for i=0 to Ubound(pcv_arrOSCheckArray)-1  'loop through OS packages
		intOSweight=0
		pcv_arrOSArray=split(pcv_arrOSCheckArray(i),"|||")
		if pcv_arrOSArray(0)>pcv_intOSStatus then
			pcv_arrOSArray2=pcv_arrOSArray(1)
			pcv_strOSString=split(pcv_arrOSArray2,"||")
			if ubound(pcv_strOSString)=-1 then
				pcv_intOSheight=UPS_HEIGHT
				pcv_intOSwidth=UPS_WIDTH
				pcv_intOSlength=UPS_LENGTH
				pcv_intOSStatus=0
			else
				intPackageCnt=intPackageCnt+1
				pcv_intOSwidth=pcv_strOSString(0)
				session("UPSPackWidth"&intPackageCnt)=pcv_strOSString(0)
				session("FEDEXWSPackWidth"&intPackageCnt)=pcv_strOSString(0) '// WS
				session("CPPackWidth"&intPackageCnt)=pcv_strOSString(0)
				session("USPSPackWidth"&intPackageCnt)=pcv_strOSString(0)
				pcv_intOSheight=pcv_strOSString(1)
				session("UPSPackHeight"&intPackageCnt)=pcv_strOSString(1)
				session("FEDEXWSPackHeight"&intPackageCnt)=pcv_strOSString(1) '// WS
				session("CPPackHeight"&intPackageCnt)=pcv_strOSString(1)
				session("USPSPackHeight"&intPackageCnt)=pcv_strOSString(1)
				pcv_intOSlength=pcv_strOSString(2)
				session("UPSPackLength"&intPackageCnt)=pcv_strOSString(2)
				session("FEDEXWSPackLength"&intPackageCnt)=pcv_strOSString(2) '// WS
				session("CPPackLength"&intPackageCnt)=pcv_strOSString(2)
				session("USPSPackLength"&intPackageCnt)=pcv_strOSString(2)
				pcv_intOSPrice=pcv_strOSString(6)

				'// Price of OverSized Package UPS
				if pcv_UseDynamicInsuredValue="1" then
					session("UPSPackPrice"&intPackageCnt)=pcv_intOSPrice
					'// subtract the price of this OS package from the subtotal if dynamic insured value is used in cart.
					pcv_InsuredValue=ccur(pcv_InsuredValue)-cdbl(pcv_intOSPrice)
				else
					session("UPSPackPrice"&intPackageCnt)=UPS_INSUREDVALUE
				end if

				'// Price of OverSized Package FedEX WS
				if pcv_UseDynamicInsuredValue_FDXWS="1" then
					session("FEDEXWSPackPrice"&intPackageCnt)=pcv_intOSPrice
					'// subtract the price of this OS package from the subtotal if dynamic insured value is used in cart.
					pcv_InsuredValue_FDXWS=ccur(pcv_InsuredValue_FDXWS)-cdbl(pcv_intOSPrice)
				else
					session("FEDEXWSPackPrice"&intPackageCnt)=FEDEXWS_INSUREDVALUE
				end if
				
				'// Price of OverSized Package FedEX WS
				if pcv_UseDynamicInsuredValue_USPS="1" then
					session("USPSPackPrice"&intPackageCnt)=pcv_intOSPrice
					'// subtract the price of this OS package from the subtotal if dynamic insured value is used in cart.
					pcv_InsuredValue_USPS=ccur(pcv_InsuredValue_USPS)-cdbl(pcv_intOSPrice)
				else
					session("USPSPackPrice"&intPackageCnt)=USPS_INSUREDVALUE
				end if

				intOSweight=pcv_strOSString(5)
				if pcv_EMWeightLimit<>0 AND intOSweight>Clng((pcv_EMWeightLimit*16)) then
					pcv_EM_Null=1
				end if
				if pcv_PMWeightLimit<>0 AND Clng(intOSweight)>Clng((pcv_PMWeightLimit*16)) then
					pcv_PM_Null=1
				end if
				intWeightCnt=intWeightCnt+intOSweight

				if scShipFromWeightUnit="KGS" then
					intOSintPounds=int(intOSweight/1000)
					intOSounces=intOSweight-(intOSintPounds*1000)
				else
					intOSintPounds=Int(intOSweight/16) 'intPounds used for USPS
					intOSounces=intOSweight-(intOSintPounds*16) 'intUniversalOunces used for USPS
				end if
				session("USPSPackPounds"&intPackageCnt)=intOSintPounds
				session("USPSPackOunces"&intPackageCnt)=intOSounces
				session("BasicPackPounds"&intPackageCnt)=intOSintPounds
				session("BasicPackOunces"&intPackageCnt)=intOSounces
                
				intMPackageWeight=intOSintPounds
				if intMPackageWeight<1 AND intOSounces<1 then
					intMPackageWeight=0
				end if
				if intMPackageWeight<1 AND intOSounces>0 then 'if total weight is less then a pound, make UPS/FedEX weight 1 pound
					intMPackageWeight=1
				else  'total weight is not less then a pound and ounces exist, round weight up one more pound.
					If intMPackageWeight>0 AND intOSounces>0 then
						intMPackageWeight=(intMPackageWeight+1)
					End if
				end if

				pcv_intTotPackageNum=pcv_intTotPackageNum+1
				session("UPSPackWeight"&intPackageCnt)=intMPackageWeight
				session("FEDEXWSPackWeight"&intPackageCnt)=intMPackageWeight
				session("FEDEXWSPackPounds"&intPackageCnt)=intOSintPounds
				session("FEDEXWSPackOunces"&intPackageCnt)=intOSounces
				session("CPPackPounds"&intPackageCnt)=intOSintPounds
				session("CPPackOunces"&intPackageCnt)=intOSounces
				session("OSFlaggedPackage"&intPackageCnt)="YES"
			end if
		end if
	next 'End loop through OS packages
	dim intOSpackageCnt
	intOSpackageCnt=intPackageCnt
else 'There are OS packages
	'no oversized packages
	pcv_intOSStatus=0
end if 'There are OS packages

'=====================================================================
intCustomShipWeight=intUniversalWeight
pShipWeight=intUniversalWeight-intWeightCnt

'no oversized items were in cart, packagecount at 1
if pcv_intOSStatus=0 then
	intPackageCnt=0
end if

if pShipWeight>0 then 'Weight > 0
	if scShipFromWeightUnit="KGS" then
		intPounds = Int(pShipWeight/1000)
		intUniversalOunces = Cdbl((pShipWeight-(intPounds*1000))/1000) 'intUniversalOunces used for USPS
        cblCPWeight = Cdbl(pShipWeight/1000) 
	else
		intPounds = Int(pShipWeight/16) 'intPounds used for USPS
		intUniversalOunces = pShipWeight-(intPounds*16) 'intUniversalOunces used for USPS
        cblCPWeight = Cdbl(pShipWeight/16) 
	end if

	intUniversalWeight=intPounds
	if intUniversalWeight<1 AND intUniversalOunces<1 then
		intUniversalWeight=0
	end if

	if intUniversalWeight<1 AND intUniversalOunces>0 then 'if total weight is less then a pound, make UPS/FedEX weight 1 pound
		intUniversalWeight=1
	else  'total weight is not less then a pound and ounces exist, round weight up one more pound.
		If intUniversalWeight>0 AND intUniversalOunces>0 then
			intUniversalWeight=(intUniversalWeight+1)
		End if
	end if
	
	'=====================================================================
	' check to see if there is a weight limit set for packages >0
	'---------------------------------------------------------------------
	if int(scPackageWeightLimit)<>0 then 'There is a package Weight limit set
		'see how many package this should be if over the limit
		if int(intUniversalWeight)>int(scPackageWeightLimit) then 'There are more package after OS
			'divide<br>
			intTempPackageNum=(intUniversalWeight/int(scPackageWeightLimit))
			pcv_intPackageNum=int(intUniversalWeight/int(scPackageWeightLimit))
			if intTempPackageNum>pcv_intPackageNum then
				pcv_intPackageNum=pcv_intPackageNum+1
			end if
			if pcv_UseDynamicInsuredValue="1" then
				pcv_TempInsuredValue=cdbl(pcv_InsuredValue)/Cint(pcv_intPackageNum)
			else
				pcv_TempInsuredValue=cdbl(pcv_InsuredValue)
			end if
			if pcv_UseDynamicInsuredValue_FDXWS="1" then
				pcv_TempInsuredValue_FDXWS=cdbl(pcv_InsuredValue_FDXWS)/Cint(pcv_intPackageNum)
			else
				pcv_TempInsuredValue_FDXWS=cdbl(pcv_InsuredValue_FDXWS)
			end if
			if pcv_UseDynamicInsuredValue_USPS="1" then
				pcv_TempInsuredValue_USPS=cdbl(pcv_InsuredValue_USPS)/Cint(pcv_intPackageNum)
			else
				pcv_TempInsuredValue_USPS=cdbl(pcv_InsuredValue_USPS)
			end if
			for r=1 to (pcv_intPackageNum-1)
				intPackageCnt=intPackageCnt+1

				pcv_intTotPackageNum=pcv_intTotPackageNum+1
				if ups_active=true or ups_active="-1" then
					session("UPSPackWidth"&intPackageCnt)=UPS_WIDTH
					session("UPSPackHeight"&intPackageCnt)=UPS_HEIGHT
					session("UPSPackLength"&intPackageCnt)=UPS_LENGTH
					session("UPSPackWeight"&intPackageCnt)=scPackageWeightLimit
					session("UPSPackPrice"&intPackageCnt)=pcv_TempInsuredValue
				end if
				if FedEXWS_active=true or FedExWS_active="-1" then
					session("FedEXWSPackWidth"&intPackageCnt)=FEDEXWS_WIDTH
					session("FedEXWSPackHeight"&intPackageCnt)=FEDEXWS_HEIGHT
					session("FedEXWSPackLength"&intPackageCnt)=FEDEXWS_LENGTH
					session("FedEXWSPackWeight"&intPackageCnt)=scPackageWeightLimit
					session("FEDEXWSPackPounds"&intPackageCnt)=scPackageWeightLimit
					session("FEDEXWSPackOunces"&intPackageCnt)=0
					session("FedEXWSPackPrice"&intPackageCnt)=pcv_TempInsuredValue_FDXWS
				end if
				if USPS_active=true or USPS_active="-1" then
					session("USPSPackWidth"&intPackageCnt)=USPS_WIDTH
					session("USPSPackHeight"&intPackageCnt)=USPS_HEIGHT
					session("USPSPackLength"&intPackageCnt)=USPS_LENGTH
					session("USPSPackPounds"&intPackageCnt)=scPackageWeightLimit
					session("USPSPackOunces"&intPackageCnt)=0
					session("USPSPackPrice"&intPackageCnt)=pcv_TempInsuredValue_USPS
					if pcv_EMWeightLimit<>0 AND scPackageWeightLimit>Clng(pcv_EMWeightLimit) then
						pcv_EM_Null=1
					end if
					if pcv_PMWeightLimit<>0 AND scPackageWeightLimit>Clng(pcv_PMWeightLimit) then
						pcv_PM_Null=1
					end if
				end if
				If CP_active=true or CP_active="-1" then
					session("CPPackWidth"&intPackageCnt)=CP_Width
					session("CPPackHeight"&intPackageCnt)=CP_Height
					session("CPPackLength"&intPackageCnt)=CP_Length
                    session("CPPackPounds"&intPackageCnt)=scPackageWeightLimit
                    session("CPPackOunces"&intPackageCnt)=0
				end if
				session("BasicPackPounds"&intPackageCnt)=intPounds
				session("BasicPackOunces"&intPackageCnt)=intUniversalOunces
			next

			'last package
			intLastPackageWeight=int(intUniversalWeight-((pcv_intPackageNum-1)*scPackageWeightLimit))
			intPackageCnt=intPackageCnt+1

			pcv_intTotPackageNum=pcv_intTotPackageNum+1
			if ups_active=true or ups_active="-1" then
				session("UPSPackWidth"&intPackageCnt)=UPS_WIDTH
				session("UPSPackHeight"&intPackageCnt)=UPS_HEIGHT
				session("UPSPackLength"&intPackageCnt)=UPS_LENGTH
				session("UPSPackWeight"&intPackageCnt)=intLastPackageWeight
				session("UPSPackPrice"&intPackageCnt)=pcv_TempInsuredValue
			end if
			if FedEXWS_active=true or FedExWS_active="-1" then
				session("FedEXWSPackWidth"&intPackageCnt)=FEDEXWS_WIDTH
				session("FedEXWSPackHeight"&intPackageCnt)=FEDEXWS_HEIGHT
				session("FedEXWSPackLength"&intPackageCnt)=FEDEXWS_LENGTH
				session("FedEXWSPackWeight"&intPackageCnt)=intLastPackageWeight
				session("FEDEXWSPackPounds"&intPackageCnt)=intLastPackageWeight
				session("FEDEXWSPackOunces"&intPackageCnt)=intUniversalOunces
				session("FedEXWSPackPrice"&intPackageCnt)=pcv_TempInsuredValue_FDXWS
			end if
			If CP_active=true or CP_active="-1" then
				session("CPPackWidth"&intPackageCnt)=CP_Width
				session("CPPackHeight"&intPackageCnt)=CP_Height
				session("CPPackLength"&intPackageCnt)=CP_Length
                session("CPPackPounds"&intPackageCnt)=intLastPackageWeight
                session("CPPackOunces"&intPackageCnt)=intUniversalOunces
			end if
			if USPS_active=true or USPS_active="-1" then
				session("USPSPackWidth"&intPackageCnt)=USPS_WIDTH
				session("USPSPackHeight"&intPackageCnt)=USPS_HEIGHT
				session("USPSPackLength"&intPackageCnt)=USPS_LENGTH
				session("USPSPackPounds"&intPackageCnt)=intLastPackageWeight
				session("USPSPackOunces"&intPackageCnt)=intUniversalOunces
				session("USPSPackPrice"&intPackageCnt)=pcv_TempInsuredValue_USPS
			end if
			session("BasicPackPounds"&intPackageCnt)=intPounds
			session("BasicPackOunces"&intPackageCnt)=intUniversalOunces
		else 'There are more package after OS

			intPackageCnt=intPackageCnt+1
			pcv_intTotPackageNum=pcv_intTotPackageNum+1
			if ups_active=true or ups_active="-1" then
				session("UPSPackWidth"&intPackageCnt)=UPS_WIDTH
				session("UPSPackHeight"&intPackageCnt)=UPS_HEIGHT
				session("UPSPackLength"&intPackageCnt)=UPS_LENGTH
				session("UPSPackWeight"&intPackageCnt)=intUniversalWeight
				session("UPSPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue)
			end if
			if FedEXWS_active=true or FedExWS_active="-1" then
				session("FedEXWSPackWidth"&intPackageCnt)=FEDEXWS_WIDTH
				session("FedEXWSPackHeight"&intPackageCnt)=FEDEXWS_HEIGHT
				session("FedEXWSPackLength"&intPackageCnt)=FEDEXWS_LENGTH
				session("FedEXWSPackWeight"&intPackageCnt)=intUniversalWeight
				session("FEDEXWSPackPounds"&intPackageCnt)=intPounds
				session("FEDEXWSPackOunces"&intPackageCnt)=intUniversalOunces
				session("FedEXWSPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue_FDXWS)
			end if

			If CP_active=true or CP_active="-1" then
				session("CPPackWidth"&intPackageCnt)=CP_Width
				session("CPPackHeight"&intPackageCnt)=CP_Height
				session("CPPackLength"&intPackageCnt)=CP_Length
                session("CPPackPounds"&intPackageCnt)=intPounds
                session("CPPackOunces"&intPackageCnt)=intUniversalOunces
			end if
			if USPS_active=true or USPS_active="-1" then
				session("USPSPackWidth"&intPackageCnt)=USPS_WIDTH
				session("USPSPackHeight"&intPackageCnt)=USPS_HEIGHT
				session("USPSPackLength"&intPackageCnt)=USPS_LENGTH
				session("USPSPackPounds"&intPackageCnt)=intPounds
				session("USPSPackOunces"&intPackageCnt)=intUniversalOunces
				session("USPSPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue_USPS)
				if pcv_EMWeightLimit<>0 AND pShipWeight>Clng((pcv_EMWeightLimit*16)) then
					pcv_EM_Null=1
				end if
				if pcv_PMWeightLimit<>0 AND pShipWeight>Clng((pcv_PMWeightLimit*16)) then
					pcv_PM_Null=1
				end if
			end if
			session("BasicPackPounds"&intPackageCnt)=intPounds
			session("BasicPackOunces"&intPackageCnt)=intUniversalOunces
		end if 'There are more package after OS
		
	else 'There is a package Weight limit set
		'no weight limit set
		intPackageCnt=intPackageCnt+1
		pcv_intTotPackageNum=pcv_intTotPackageNum+1
		if ups_active=true or ups_active="-1" then
			session("UPSPackWidth"&intPackageCnt)=UPS_WIDTH
			session("UPSPackHeight"&intPackageCnt)=UPS_HEIGHT
			session("UPSPackLength"&intPackageCnt)=UPS_LENGTH
			session("UPSPackWeight"&intPackageCnt)=intUniversalWeight
			session("UPSPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue)
		end if
		if FedEXWS_active=true or FedExWS_active="-1" then
			session("FedEXWSPackWidth"&intPackageCnt)=FEDEXWS_WIDTH
			session("FedEXWSPackHeight"&intPackageCnt)=FEDEXWS_HEIGHT
			session("FedEXWSPackLength"&intPackageCnt)=FEDEXWS_LENGTH
			session("FedEXWSPackWeight"&intPackageCnt)=intUniversalWeight
			session("FEDEXWSPackPounds"&intPackageCnt)=intPounds
			session("FEDEXWSPackOunces"&intPackageCnt)=intUniversalOunces
			session("FedEXWSPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue_FDXWS)
		end if

		If CP_active=true or CP_active="-1" then
			session("CPPackWidth"&intPackageCnt)=CP_Width
			session("CPPackHeight"&intPackageCnt)=CP_Height
			session("CPPackLength"&intPackageCnt)=CP_Length
            session("CPPackPounds"&intPackageCnt)=intPounds
            session("CPPackOunces"&intPackageCnt)=intUniversalOunces
		end if
		if USPS_active=true or USPS_active="-1" then
			session("USPSPackWidth"&intPackageCnt)=USPS_WIDTH
			session("USPSPackHeight"&intPackageCnt)=USPS_HEIGHT
			session("USPSPackLength"&intPackageCnt)=USPS_LENGTH
			session("USPSPackPounds"&intPackageCnt)=intPounds
			session("USPSPackOunces"&intPackageCnt)=intUniversalOunces
			session("USPSPackPrice"&intPackageCnt)=cdbl(pcv_InsuredValue_USPS)
			if pcv_EMWeightLimit<>0 AND pShipWeight>Clng((pcv_EMWeightLimit*16)) then
				pcv_EM_Null=1
			end if
			if pcv_PMWeightLimit<>0 AND pShipWeight>Clng((pcv_PMWeightLimit*16)) then
				pcv_PM_Null=1
			end if
		end if

		session("BasicPackPounds"&intPackageCnt)=intPounds
		session("BasicPackOunces"&intPackageCnt)=intUniversalOunces
	end if 'There is a package Weight limit set

end if 'Weight > 0
'=====================================================================

pcv_intPackageNum=intPackageCnt

'string
availableShipStr=""
dim iUPSActive, iFedExWSActive, iUSPSActive, iCPActive
iUPSActive=0
iFedExWSActive=0
iUSPSActive=0
iCPActive=0
UPS_ShipFromCity = scShipFromCity
UPS_ShipFromState = scShipFromState
UPS_ShipFromPostalCode = scShipFromPostalCode
UPS_ShipFromPostalCountry = scShipFromPostalCountry
%>
<%
'// FedEX WS
%>
<!--#include file="FedExWebServices.asp"-->
<% If FEDEXWS_SATURDAYDELIVERY<>"0" Then %>
	<!--#include file="FedExWebServicesSaturday.asp"-->
<% End If %>

<%
	
if iFedExWSFlag=1 then
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Pre-Filter FEDEX WS for Availability
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	pcv_strOptionFilterPass=pcf_PreFilter("FEDEXWS", availableShipStr)

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Pre-Filter FEDEX WS for Availability
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if pcv_strOptionFilterPass=-1 then
		strDefaultProvider="FEDEXWS"
		iShipmentTypeCnt=iShipmentTypeCnt+1
		strOptionShipmentType=strOptionShipmentType&"<option value=FedExWS>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_m")&"</option>"
		strTabShipmentType=strTabShipmentType&"[TAB]FedExWS,"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_m")&"[/TAB]"
	end if
end if

if pcv_UPSCanadaOrigin=0 then
	'//Get US Origin UPS ShipType Names
	pServiceCodeString01="UPS Next Day Air&reg;"
	pServiceCodeString02="UPS 2nd Day Air&reg;"
	pServiceCodeString03="UPS Ground"
	pServiceCodeString07="UPS Worldwide Express<sup>SM</sup>"
	pServiceCodeString08="UPS Worldwide Expedited<sup>SM</sup>"
	pServiceCodeString11="UPS Standard To Canada"
	pServiceCodeString12="UPS 3 Day Select<sup>SM</sup>"
	pServiceCodeString13="UPS Next Day Air Saver&reg;"
	pServiceCodeString14="UPS Next Day Air&reg; Early A.M.&reg;"
	pServiceCodeString54="UPS Worldwide Express Plus<sup>SM</sup>"
	pServiceCodeString59="UPS 2nd Day Air A.M.&reg;"
	pServiceCodeString65="UPS Express Saver<sup>SM</sup>"
else
	pServiceCodeString01="UPS Express<sup>SM</sup>"
	pServiceCodeString02="UPS Expedited<sup>SM</sup>"
	pServiceCodeString03=""
	pServiceCodeString07="UPS Worldwide Express<sup>SM</sup>"
	pServiceCodeString08="UPS Worldwide Expedited<sup>SM</sup>"
	pServiceCodeString11="UPS Standard To Canada"
	pServiceCodeString12="UPS 3 Day Select<sup>SM</sup>"
	pServiceCodeString13="UPS Express Saver&reg;"
	pServiceCodeString14="UPS Express Saver&reg; Early A.M.&reg;"
	pServiceCodeString54="UPS Worldwide Express Plus<sup>SM</sup>"
	pServiceCodeString59=""
	pServiceCodeString65=""
end if
err.clear %>

<!--#include file="UPSShipRates.asp"-->
<%
if pcv_UPSDebug=1 then
	'// Show Post
	response.write ups_postdata&"<HR><BR>"
	'// Show Reply
	response.write "UPS Reply: "&UPS_result&"<BR>"
	response.End()
end if

if pcv_UPS_Logging=1 then
	'/////////////////////////////////////////////////////
	'// Create Log of response and save in includes
	'/////////////////////////////////////////////////////
	if PPD="1" then
		pcStrFileName=Server.Mappath ("/"&scPcFolder&"/includes/UPSLOG.LOG")
	else
		pcStrFileName=Server.Mappath ("../includes/UPSLOG.LOG")
	end if

	dim strFileName
	dim strItem
	dim fs
	dim OutputFile
	dim t

	'Specify directory and file to store silent post information
	strFileName = pcStrFileName
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set OutputFile = fs.OpenTextFile (strFileName, 8, True)

	OutputFile.WriteLine now()
	OutputFile.WriteLine "UPS XML REQUEST: "
	OutputFile.WriteLine ups_postdata
	OutputFile.WriteBlankLines(2)
	OutputFile.WriteLine "ANY ERRORS: "
	OutputFile.WriteLine err.description
	OutputFile.WriteBlankLines(2)
	OutputFile.WriteLine "UPS XML RESPONSE: "
	OutputFile.WriteLine UPS_result
	OutputFile.WriteBlankLines(2)

	OutputFile.Close
	'/////////////////////////////////////////////////////
	'// End - Create Log of response and save in includes
	'/////////////////////////////////////////////////////
end if

if iUPSFlag=1 then
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Pre-Filter UPS for Availability
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	pcv_strOptionFilterPass=pcf_PreFilter("UPS", availableShipStr)
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Pre-Filter UPS for Availability
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if pcv_strOptionFilterPass=-1 then
		strDefaultProvider="UPS"
		iShipmentTypeCnt=iShipmentTypeCnt+1
		strOptionShipmentType=strOptionShipmentType&"<option value=UPS>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_k")&"</option>"
		strTabShipmentType=strTabShipmentType&"[TAB]UPS,"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_k")&"[/TAB]"
	end if
end if

If usps_active=true or usps_active="-1" then
	iUSPSActive=1
	Dim USPS_PackageSize
	'check that all packages can go through USPS
	dim intUSPSnoShpmnt
	intUSPSnoShpmnt=0
	session("BMP")="Y"
	USPS_DWeightOver =""
	USPS_PackageSizeC=""
	for q=1 to pcv_intPackageNum
		'//If any one side is greater then 12" package is labeled as "LARGE"
		If USPS_PackageSizeC="" AND ((Cint(session("USPSPackLength"&q))>12) OR (Cint(session("USPSPackWidth"&q))>12) OR (Cint(session("USPSPackHeight"&q))>12)) Then
			USPS_PackageSizeC="LARGE"
		Else
			USPS_PackageSize=(Cint(session("USPSPackLength"&q)) + ((Cint(session("USPSPackWidth"&q))*2)+(Cint(session("USPSPackHeight"&q))*2)))
			If USPS_PackageSize<85 then
				USPS_PackageSizeC="REGULAR"
			End if
			If USPS_PackageSize>85 AND USPS_PackageSize<108 AND intPounds<15 then
				USPS_PackageSizeC="LARGE"
				if (Cint(Cint(session("USPSPackLength"&q))) * (Cint(Cint(session("USPSPackWidth"&q))) * Cint(Cint(session("USPSPackHeight"&q)))))>1728 then
					USPS_DWeightOver = "YES"
				end if
			End if
			If USPS_PackageSizeC="" AND USPS_PackageSize<131 then
				USPS_PackageSizeC="OVERSIZE"
			End if
			if USPS_PackageSizeC="" OR USPS_PackageSizeC="REGULAR" then
				if (Cint(Cint(session("USPSPackLength"&q))) * (Cint(Cint(session("USPSPackWidth"&q))) * Cint(Cint(session("USPSPackHeight"&q)))))>1728 then
					USPS_PackageSizeC="LARGE"
					USPS_DWeightOver = "YES"
				end if
			end if
		End If
	next
	If USPS_PackageSizeC="" then
		intUSPSnoShpmnt=1
	end if
	IF session("USPSPackPounds"&q)>15 THEN
		session("BMP")="N"
	end if

	If intUSPSnoShpmnt=0 then
		'//USPS RATES - Domestic
		If Universal_destination_country="US" then
			'parse +4 from the zip code
			if len(Universal_destination_postal)>5 then
				Universal_destination_postal=left(Universal_destination_postal,5)
			end if
			usps_postdata=""
			usps_postdata=usps_postdata&usps_server&"?API=RateV4&XML="
			
			usps_postxml=""
			usps_postxml=usps_postxml&"<RateV4Request USERID="&chr(34)&usps_userid&chr(34)&">"
			
			for q=1 to pcv_intPackageNum
				iNum=q-1
				USPS_PackageSizeC=""

				'//If any one side is greater then 12" package is labeled as "LARGE"
				If (Cint(session("USPSPackLength"&q))>12) OR (Cint(session("USPSPackWidth"&q))>12) OR (Cint(session("USPSPackHeight"&q))>12) Then
					USPS_PackageSizeC="LARGE"
					USPS_DWeightOver = "YES"
				Else
					USPS_PackageSize=(Cint(session("USPSPackLength"&q)) + ((Cint(session("USPSPackWidth"&q))*2)+(Cint(session("USPSPackHeight"&q))*2)))
					If USPS_PackageSize<85 then
						USPS_PackageSizeC="REGULAR"
					End if
					If USPS_PackageSize>85 AND USPS_PackageSize<108 then
						USPS_PackageSizeC="LARGE"
						if (Cint(Cint(session("USPSPackLength"&q))) * (Cint(Cint(session("USPSPackWidth"&q))) * Cint(Cint(session("USPSPackHeight"&q)))))>1728 then
							USPS_DWeightOver = "YES"
						end if
					End if
					If USPS_PackageSizeC="" AND USPS_PackageSize<131 then
						USPS_PackageSizeC="OVERSIZE"
					End if
					if USPS_PackageSizeC="" OR USPS_PackageSizeC="REGULAR" then
						if (Cint(Cint(session("USPSPackLength"&q))) * (Cint(Cint(session("USPSPackWidth"&q))) * Cint(Cint(session("USPSPackHeight"&q)))))>1728 then
							USPS_PackageSizeC="LARGE"
							USPS_DWeightOver = "YES"
						end if
					end if
				End If

				usps_postxml=usps_postxml&"<Package ID="&chr(34)&iNum&chr(34)&">"
				usps_postxml=usps_postxml&"<Service>All</Service>"
				usps_postxml=usps_postxml&"<ZipOrigination>"&scShipFromPostalCode&"</ZipOrigination>"
				usps_postxml=usps_postxml&"<ZipDestination>"&Universal_destination_postal&"</ZipDestination>"
				usps_postxml=usps_postxml&"<Pounds>"&session("USPSPackPounds"&q)&"</Pounds>"
				usps_postxml=usps_postxml&"<Ounces>"&round(session("USPSPackOunces"&q))&"</Ounces>"
				if USPS_DWeightOver = "YES" then
					usps_postxml=usps_postxml&"<Container>RECTANGULAR</Container>"
					usps_postxml=usps_postxml&"<Size>"&USPS_PackageSizeC&"</Size>"
					usps_postxml=usps_postxml&"<Width>"&Cint(session("USPSPackWidth"&q))&"</Width>"
					usps_postxml=usps_postxml&"<Length>"&Cint(session("USPSPackLength"&q))&"</Length>"
					usps_postxml=usps_postxml&"<Height>"&Cint(session("USPSPackHeight"&q))&"</Height>"
					usps_postxml=usps_postxml&"<Girth>"&USPS_PackageSize&"</Girth>"
				else

					usps_postxml=usps_postxml&"<Container>VARIABLE</Container>"
					usps_postxml=usps_postxml&"<Size>"&USPS_PackageSizeC&"</Size>"
				end if
					
				IF USPS_PackageSizeC="LARGE" THEN
					'Check if Machinable or not
					if Cint(session("USPSPackLength"&q))<3 OR Cint(session("USPSPackLength"&q))>34 OR Cint(session("USPSPackWidth"&q))<3 OR Cint(session("USPSPackWidth"&q))>17 OR Cint(session("USPSPackHeight"&q))>17 OR USPS_DWeightOver = "YES" then
						usps_postxml=usps_postxml&"<Machinable>false</Machinable>"
					else
						usps_postxml=usps_postxml&"<Machinable>true</Machinable>"
					end if
				else
					usps_postxml=usps_postxml&"<Machinable>true</Machinable>"
				END IF
				
				usps_postxml=usps_postxml&"</Package>"
			next

			usps_postxml=usps_postxml&"</RateV4Request>"
			
			usps_postdata=usps_postdata&usps_postxml
			
'			Response.Clear()
'			Response.ContentType = "text/xml"
'			Response.write usps_postxml
'			Response.End()
			
			Set srvUSPS2XmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
			srvUSPS2XmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
			srvUSPS2XmlHttp.open "GET", usps_postdata, false
			srvUSPS2XmlHttp.send
			USPS2_result = srvUSPS2XmlHttp.responseText
			
'			Response.Clear()
'			Response.ContentType = "text/xml"
'			Response.write USPS2_result
'			Response.End()

			' Parse the XML document.
			Set objOutputXMLDoc = Server.CreateObject("Microsoft.XMLDOM")
			objOutputXMLDoc.loadXML srvUSPS2XmlHttp.responseText

			Set Nodes = objOutputXMLDoc.selectNodes("//Postage")

			iUSPSEMFlag=0
			iUSPSPMFlag=0
			iUSPSFCFlag=0
			iUSPSPPFlag=0
			iUSPSBPMFlag=0
			iUSPSMMFlag=0
			iUSPSLMFlag=0
			iUSPSEMRate=0
			iUSPSPMRate=0
			iUSPSFCRate=0
			iUSPSPPRate=0
			iUSPSBPMRate=0
			iUSPSMMRate=0
			iUSPSLMRate=0
			iUSPSEMCnt=0
			iUSPSPMCnt=0
			iUSPSFCCnt=0
			iUSPSPPCnt=0
			iUSPSBPMCnt=0
			iUSPSMMCnt=0
			iUSPSLMCnt=0
			iUSPSEMFlagAdded=0
			iUSPSPMFlagAdded=0
			iUSPSFCFlagAdded=0
			iUSPSPPFlagAdded=0
			iUSPSBPMFlagAdded=0
			iUSPSMMFlagAdded=0
			iUSPSLMFlagAdded=0

			USPSErrorDetect1=0

			set objLst=objOutputXMLDoc.getElementsByTagName("Package")
			for i = 0 to (objLst.length - 1)
				for j=0 to ((objLst.item(i).childNodes.length)-1)
					If objLst.item(i).childNodes(j).nodeName="Error" then
						USPSErrorDetect1=1
						for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="Number" then
								strUSPSError = objLst.item(i).childNodes(j).childNodes(m).text
							end if
						next
					End if
				next
			next

			for i = 0 to (objLst.length - 1)
				for j=0 to ((objLst.item(i).childNodes.length)-1)
					If objLst.item(i).childNodes(j).nodeName="Postage" then
						intCLASSID=objLst.item(i).childNodes(j).getAttribute("CLASSID")
						for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="MailService" then
								strMailService = objLst.item(i).childNodes(j).childNodes(m).text
							end if
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="Rate" then
								strRate = objLst.item(i).childNodes(j).childNodes(m).text
							end if

						next

						if USPSErrorDetect1=0 then

							select case intCLASSID

							case "1"
								if ucase(USPS_PM_PACKAGE)="NONE" OR (pcv_PM_Null=1 AND USPS_PM_FREOption="NONE") then
									iUSPSPMFlag=1
									iUSPSPMCnt=iUSPSPMCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSPMRate=iUSPSPMRate+strRate
								end if
							case "3"
								if ucase(USPS_EM_PACKAGE)="NONE" OR (pcv_EM_Null=1 AND USPS_EM_FREOption="1") then
									iUSPSEMFlag=1
									iUSPSEMCnt=iUSPSEMCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSEMRate=iUSPSEMRate+strRate
								end if
							case "0"
								If instr(lcase(strMailService), "package") Then
									if iUSPSFCFlag=0 then
										iUSPSFCFlag=1
										iUSPSFCCnt=iUSPSFCCnt+1
										iUSPSFlag=1
										if isNumeric(strRate) then
											strRate=cdbl(strRate)
										end if
										iUSPSFCRate=iUSPSFCRate+strRate
									end if
								End If
							case "4"
									iUSPSPPFlag=1
									iUSPSPPCnt=iUSPSPPCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSPPRate=iUSPSPPRate+strRate
							case "Bound Printed Matter"
								if session("BMP")="Y" then
									iUSPSBPMFlag=1
									iUSPSBPMCnt=iUSPSBPMCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSBPMRate=iUSPSBPMRate+strRate
								end if
							case "6"
									iUSPSMMFlag=1
									iUSPSMMCnt=iUSPSMMCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSMMRate=iUSPSMMRate+strRate
							case "7"
									iUSPSLMFlag=1
									iUSPSLMCnt=iUSPSLMCnt+1
									iUSPSFlag=1
									if isNumeric(strRate) then
										strRate=cdbl(strRate)
									end if
									iUSPSLMRate=iUSPSLMRate+strRate
							end select
							'Priority Mail
							if iUSPSPMCnt<pcv_intPackageNum then
								iUSPSPMFlag=0
							end if
							if (iUSPSPMFlag=1) AND (iUSPSPMFlagAdded = 0) AND (USPS_OverSizeOrder=0) then
								availableShipStr=availableShipStr&"|?|USPS|9901|"&"Priority Mail <sup>&reg;</sup>|"&iUSPSPMRate&"|NA"
								iUSPSPMFlagAdded = 1
							end if
							'Express Mail
							if iUSPSEMCnt<pcv_intPackageNum then
								iUSPSEMFlag=0
							end if
							if (iUSPSEMFlag=1) AND (iUSPSEMFlagAdded = 0) AND (USPS_OverSizeOrder=0) then
								availableShipStr=availableShipStr&"|?|USPS|9902|"&"Express Mail <sup>&reg;</sup>|"&iUSPSEMRate&"|NA"
								iUSPSEMFlagAdded=1
							end if
							'First Class Mail
							if iUSPSFCCnt<pcv_intPackageNum then
								iUSPSFCFlag=0
							end if
							if (iUSPSFCFlag=1) AND (iUSPSFCFlagAdded = 0) AND (USPS_OverSizeOrder=0) then
								availableShipStr=availableShipStr&"|?|USPS|9904|"&"First-Class Mail <sup>&reg;</sup>|"&iUSPSFCRate&"|NA"
								iUSPSFCFlagAdded =1
							end if
							'Retail Ground (v2 Standard Post)
							if iUSPSPPCnt<pcv_intPackageNum then
								iUSPSPPFlag=0
							end if
							if (iUSPSPPFlag=1)  AND (iUSPSPPFlagAdded = 0) AND (USPS_OverSizeOrder=0) then
								availableShipStr=availableShipStr&"|?|USPS|9903|"&"Retail Ground  <sup>&trade;</sup>|"&iUSPSPPRate&"|NA"
								iUSPSPPFlagAdded = 1
							end if
							'Bound Printed Matter
							if iUSPSBPMCnt<pcv_intPackageNum then
								iUSPSBPMFlag=0
							end if
							if (iUSPSBPMFlag=1) AND (iUSPSBPMFlagAdded = 0) AND (USPS_OverSizeOrder=0) then
								availableShipStr=availableShipStr&"|?|USPS|9915|"&"Bound Printed Matter <sup>&reg;</sup>|"&iUSPSBPMRate&"|NA"
								iUSPSBPMFlagAdded = 1
							end if
							'Media Mail
							if iUSPSMMCnt<pcv_intPackageNum then
								iUSPSMMFlag=0
							end if
							if (iUSPSMMFlag=1) AND (iUSPSMMFlagAdded = 0) AND (USPS_OverSizeOrder=0) then
								availableShipStr=availableShipStr&"|?|USPS|9916|"&"Media Mail <sup>&reg;</sup>|"&iUSPSMMRate&"|NA"
								iUSPSMMFlagAdded = 1
							end if
							'Library Mail
							if iUSPSLMCnt<pcv_intPackageNum then
								iUSPSLMFlag=0
							end if
							if (iUSPSLMFlag=1) AND (iUSPSLMFlagAdded = 0) AND (USPS_OverSizeOrder=0) then
								availableShipStr=availableShipStr&"|?|USPS|9917|"&"Library Mail <sup>&reg;</sup>|"&iUSPSLMRate&"|NA"
								iUSPSLMFlagAdded = 1
							end if

						end if
					End If
				Next
			Next


		end if

		'//USPS RATES - Express
		if ucase(USPS_EM_PACKAGE)<>"NONE" then

			'If envelope isn't valid due to weight, check if Your Package is allowed
			if (pcv_EM_Null=1 AND USPS_EM_FREOption="1") OR pcv_EM_Null=0 then

				If Universal_destination_country="US" then
					usps_postdata=""
					usps_postdata=usps_postdata&usps_server&"?API=RateV4&XML="
					
					usps_postxml=""
					usps_postxml=usps_postxml&"<RateV4Request USERID="&chr(34)&usps_userid&chr(34)&">"

					for q=1 to pcv_intPackageNum
						pcv_EM_Package=""
						iNum=q-1
						USPS_PackageSizeC=""
						'//If any one side is greater then 12" package is labeled as "LARGE"
						If (Cint(session("USPSPackLength"&q))>12) OR (Cint(session("USPSPackWidth"&q))>12) OR (Cint(session("USPSPackHeight"&q))>12) Then
							USPS_PackageSizeC="LARGE"
						Else
							USPS_PackageSize=(Cint(session("USPSPackLength"&q)) + ((Cint(session("USPSPackWidth"&q))*2)+(Cint(session("USPSPackHeight"&q))*2)))
							If USPS_PackageSize<85 then
								USPS_PackageSizeC="REGULAR"
							End if
							If USPS_PackageSize>85 AND USPS_PackageSize<108 AND intPounds<15 then
								USPS_PackageSizeC="LARGE"
							End if
							If USPS_PackageSizeC="" AND USPS_PackageSize<131 then
								USPS_PackageSizeC="OVERSIZE"
							End if
							pcv_EM_Package=USPS_EM_PACKAGE
							if USPS_PackageSizeC="LARGE" OR USPS_PackageSizeC="OVERSIZE" then
								pcv_EM_Package=""
							end if
						End If
						usps_postxml=usps_postxml&"<Package ID="&chr(34)&iNum&chr(34)&">"
						usps_postxml=usps_postxml&"<Service>Express</Service>"
						usps_postxml=usps_postxml&"<ZipOrigination>"&scShipFromPostalCode&"</ZipOrigination>"
						usps_postxml=usps_postxml&"<ZipDestination>"&Universal_destination_postal&"</ZipDestination>"
						usps_postxml=usps_postxml&"<Pounds>"&session("USPSPackPounds"&q)&"</Pounds>"
						usps_postxml=usps_postxml&"<Ounces>"&round(session("USPSPackOunces"&q))&"</Ounces>"
						'// If FRE is the default, check for weight limit and alternate container
						if pcv_EM_Null=1 AND USPS_EM_FREOption="1" then
							pcv_EM_Package="NONE"
						end if
						usps_postxml=usps_postxml&"<Container>"&pcv_EM_Package&"</Container>"
						usps_postxml=usps_postxml&"<Size>"&USPS_PackageSizeC&"</Size>"
						usps_postxml=usps_postxml&"</Package>"

					next

					usps_postxml=usps_postxml&"</RateV4Request>"
					
'					Response.Clear()
'					Response.ContentType = "text/xml"
'					Response.Write usps_postxml
'					Response.End()
					
					usps_postdata=usps_postdata&usps_postxml

					Set srvUSPS2XmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
					srvUSPS2XmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
					srvUSPS2XmlHttp.open "GET", usps_postdata, false
					srvUSPS2XmlHttp.send
					USPS2_result = srvUSPS2XmlHttp.responseText

'					Response.Clear()
'					Response.ContentType = "text/xml"
'					Response.Write USPS2_result
'					Response.End()

					' Parse the XML document.
					Set objOutputXMLDoc = Server.CreateObject("Microsoft.XMLDOM")
					objOutputXMLDoc.loadXML USPS2_result

					intUSPSPostage=0
					session("EMService")=""

					set objLst=objOutputXMLDoc.getElementsByTagName("Package")
					for i = 0 to (objLst.length - 1)
						USPS_TempSize=""
						for j=0 to ((objLst.item(i).childNodes.length)-1)
							If objLst.item(i).childNodes(j).nodeName="Size" then
								USPS_TempSize=objLst.item(i).childNodes(j).Text
							End if
							If objLst.item(i).childNodes(j).nodeName="Postage" then
								for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
									If objLst.item(i).childNodes(j).childNodes(m).nodeName="MailService" then
										strMailService = objLst.item(i).childNodes(j).childNodes(m).text
									end if
									If objLst.item(i).childNodes(j).childNodes(m).nodeName="Rate" then
										strRate = objLst.item(i).childNodes(j).childNodes(m).text
									end if
								next
							End if
						next
						pcv_EM_MailService=""
						if instr(strMailService, "(") then
							arrMailService=split(strMailService,"(")
							strMailService=arrMailService(0)
						end if

						if (instr(strMailService, "Priority Mail Express")) AND (USPS_OverSizeOrder=0) then
							strMailService = replace(strMailService,"&amp;lt;","<")
							strMailService = replace(strMailService,"&amp;gt;",">")
							strMailService = replace(strMailService,"&amp;","&")
							strMailService = replace(strMailService,"&lt;","<")
							strMailService = replace(strMailService,"&gt;",">")

							pcv_EM_MailService="USPS "&strMailService
							session("USPSshipStr")="|?|USPS|9902|"&"X|X|X|"
							if isNumeric(strRate) then
								strRate=cdbl(strRate)
							end if
							intUSPSPostage=intUSPSPostage+strRate
							iUSPSFlag=1
						end if

						if USPS_TempSize="LARGE" then
							session("EMService")="LARGE"
						end if
					next

					if objLst.length > 0 then
						if session("EMService")="LARGE" then
							pcv_EM_MailService="USPS Express Mail <sup>&reg;</sup>"
						end if
						availableShipStr=availableShipStr&replace(session("USPSshipStr"),"X|X|X|", pcv_EM_MailService)&"|"&intUSPSPostage&"|NA"
						session("USPSshipStr")=""
					end if

				end if
			end if
		end if
		
		
		'Priority Mail
		If Universal_destination_country="US" AND iUSPSPMFlag=0 then
			pcv_PMOption=USPS_PM_FREOption

			if isNull(pcv_PMOption) or pcv_PMOption="" then
				pcv_PMOption="0"
			end if

			if (pcv_PM_Null=1 AND pcv_PMOption<>"0" AND pcv_PMOption<>"NONE") OR pcv_PM_Null=0 then

				usps_postdata=""
				usps_postdata=usps_postdata&usps_server&"?API=RateV4&XML="
				
				usps_postxml=""
				usps_postxml=usps_postxml&"<RateV4Request USERID="&chr(34)&usps_userid&chr(34)&">"
				usps_postxml=usps_postxml&"<Revision>2</Revision>"
				for q=1 to pcv_intPackageNum
					iNum=q-1
					USPS_PackageSizeC=""
					'//If any one side is greater then 12" package is labeled as "LARGE"
					If (Cint(session("USPSPackLength"&q))>12) OR (Cint(session("USPSPackWidth"&q))>12) OR (Cint(session("USPSPackHeight"&q))>12) Then
						USPS_PackageSizeC="LARGE"
					Else
						USPS_PackageSize=(Cint(session("USPSPackLength"&q)) + ((Cint(session("USPSPackWidth"&q))*2)+(Cint(session("USPSPackHeight"&q))*2)))
						If USPS_PackageSize<85 then
							USPS_PackageSizeC="REGULAR"
						End if
						If USPS_PackageSize>85 AND USPS_PackageSize<108 AND intPounds<15 then
							USPS_PackageSizeC="LARGE"
						End if
						If USPS_PackageSizeC="" AND USPS_PackageSize<131 then
							USPS_PackageSizeC="OVERSIZE"
						End if
					End If
					'//Eliminate small box if any side is over 3 inchdes
					dim intshowSmallBox, intShowMedBox
					intshowSmallBox = 1
					intShowMedBox = 1

					LengthIsCal = 0
					A = Cint(session("USPSPackLength"&q))
					B = Cint(session("USPSPackWidth"&q))
					C = Cint(session("USPSPackHeight"&q))
					IntLongestLength = Cint(0)
					IntMidLength = Cint(0)
					IntShortestLength = Cint(0)

					If A=>B AND A=>C Then
						'A is the longest
						IntLongestLength = A
						LengthIsCal = 1
						If B=>C Then
							'B is the mid
							IntMidLength = B
							'C is the shortest
							IntShortestLength = C
						Else
							'C is the mid
							IntMidLength = C
							'B is the shortest
							IntShortestLength = B
						End If
					End If

					If (B=>A AND B=>C) AND (LengthIsCal = 0) Then
						'B is the longest
						IntLongestLength = B
						LengthIsCal = 1
						If A=>C Then
							'A is the mid
							IntMidLength = A
							'C is the shortest
							IntShortestLength = C
						Else
							'C is the mid
							IntMidLength = C
							'A is the shortest
							IntShortestLength = A
						End If
					End If

					If (C=>A AND C=>B) AND (LengthIsCal = 0) Then
						'C is the longest
						IntLongestLength = C
						LengthIsCal = 1
						If B=>A Then
							'B is the mid
							IntMidLength = B
							'A is the shortest
							IntShortestLength = A
						Else
							'A is the mid
							IntMidLength = A
							'B is the shortest
							IntShortestLength = B
						End If
					End If

					If IntShortestLength=>5.50 Then
						tUSPS_PM_PACKAGE = "NONE"
						pcv_PM_Null=0
					Else
						tUSPS_PM_PACKAGE=USPS_PM_PACKAGE
					End If

					If tUSPS_PM_PACKAGE <> "NONE" AND IntShortestLength=>1.5 Then
						intShowSmallBox = 0
						pcv_PMOption="Flat Rate Box1"
					End If
					If pcv_PMOption="Flat Rate Box1" AND IntShortestLength=>3.5 Then
						intShowMedBox = 0
						pcv_PMOption="Flat Rate Box2"
					End If

					usps_postxml=usps_postxml&"<Package ID="&chr(34)&iNum&chr(34)&">"
					usps_postxml=usps_postxml&"<Service>PRIORITY</Service>"
					usps_postxml=usps_postxml&"<ZipOrigination>"&scShipFromPostalCode&"</ZipOrigination>"
					usps_postxml=usps_postxml&"<ZipDestination>"&Universal_destination_postal&"</ZipDestination>"
					usps_postxml=usps_postxml&"<Pounds>"&session("USPSPackPounds"&q)&"</Pounds>"
					usps_postxml=usps_postxml&"<Ounces>"&round(session("USPSPackOunces"&q))&"</Ounces>"

					'// If FRE is the default, check for weight limit and alternate container
					if pcv_PM_Null=1 AND pcv_PMOption<>"0" then
						if pcv_PMOption="Flat Rate Box" AND intshowSmallBox = 1 then
							pcv_PMOption = "Sm Flat Rate Box"
						end if
						if pcv_PMOption="Flat Rate Box1" AND intShowMedBox = 1 then
							pcv_PMOption = "Md Flat Rate Box"
						end if
						if pcv_PMOption="Flat Rate Box2" then
							pcv_PMOption = "Lg Flat Rate Box"
						end if
						tUSPS_PM_PACKAGE=pcv_PMOption
					end if

					'private const USPS_PM_FREOption="0"
					if USPS_PM_PACKAGE<>"Flat Rate Envelope" then
						if tUSPS_PM_PACKAGE <> "NONE" then
							'pcv_PMOption=USPS_PM_PACKAGE
							if pcv_PMOption="Flat Rate Box" AND intshowSmallBox = 1 then
								pcv_PMOption = "Sm Flat Rate Box"
							end if
							if pcv_PMOption="Flat Rate Box1" AND intShowMedBox = 1 then
								pcv_PMOption = "Md Flat Rate Box"
							end if
							if pcv_PMOption="Flat Rate Box2" then
								pcv_PMOption = "Lg Flat Rate Box"
							end if
							tUSPS_PM_PACKAGE=pcv_PMOption
						else
							if ucase(tUSPS_PM_PACKAGE)="NONE" then
								'check for Priority Mail totals from previous
								if iUSPSPMRate=0 then
									tUSPS_PM_PACKAGE="VARIABLE"
								end if
							end if
						end if
					end if

					if USPS_PackageSizeC="REGULAR" AND ucase(tUSPS_PM_PACKAGE)="NONE" then
						tUSPS_PM_PACKAGE="VARIABLE"
					end if
					
					IF USPS_PackageSizeC="LARGE" then
						tUSPS_PM_PACKAGE="RECTANGULAR"
					end if
					usps_postxml=usps_postxml&"<Container>"&tUSPS_PM_PACKAGE&"</Container>"
					usps_postxml=usps_postxml&"<Size>"&USPS_PackageSizeC&"</Size>"
					IF USPS_PackageSizeC="LARGE" THEN
						usps_postxml=usps_postxml&"<Width>"&session("USPSPackWidth"&q)&"</Width>"
						usps_postxml=usps_postxml&"<Length>"&session("USPSPackLength"&q)&"</Length>"
						usps_postxml=usps_postxml&"<Height>"&session("USPSPackHeight"&q)&"</Height>"
						usps_postxml=usps_postxml&"<Girth>"&USPS_PackageSize&"</Girth>"
					END IF
					If session("USPSPackPrice"&q)&"" = "" Then
						session("USPSPackPrice"&q) = "100"
					End If
					usps_postxml=usps_postxml&"<Value>"&session("USPSPackPrice"&q)&"</Value>"
				
						
					usps_postxml=usps_postxml&"</Package>"
				next

				usps_postxml=usps_postxml&"</RateV4Request>"
				
				usps_postdata=usps_postdata&usps_postxml
			
'				Response.Clear()
'				Response.ContentType = "text/xml"
'				Response.write usps_postxml
'				Response.End()

				intUSPSPostage=0
				session("PMService")=""
				strMailService=""
				Set srvUSPS2XmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
				srvUSPS2XmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
				srvUSPS2XmlHttp.open "GET", usps_postdata, false
				srvUSPS2XmlHttp.send
				USPS2_result = srvUSPS2XmlHttp.responseText
				
'				Response.Clear()
'				Response.ContentType = "text/xml"
'				Response.write USPS2_result
'				Response.End()

				' Parse the XML document.
				Set objOutputXMLDoc = Server.CreateObject("Microsoft.XMLDOM")
				objOutputXMLDoc.loadXML srvUSPS2XmlHttp.responseText

				set objLst=objOutputXMLDoc.getElementsByTagName("Package")
				for i = 0 to (objLst.length - 1)
					USPS_TempSize=""
					for j=0 to ((objLst.item(i).childNodes.length)-1)
						If objLst.item(i).childNodes(j).nodeName="Size" then
							USPS_TempSize=objLst.item(i).childNodes(j).Text
						End if
						If objLst.item(i).childNodes(j).nodeName="Postage" then
							for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
								If objLst.item(i).childNodes(j).childNodes(m).nodeName="MailService" then
									strMailService = objLst.item(i).childNodes(j).childNodes(m).text
								end if
								If objLst.item(i).childNodes(j).childNodes(m).nodeName="Rate" then
									strRate = objLst.item(i).childNodes(j).childNodes(m).text
								end if
							next
						End if
					next

					pcv_PM_MailService=""
					if instr(strMailService, "(") then
						arrMailService=split(strMailService,"(")
						tstrMailService=arrMailService(0)
					end if
					intHasPriority = Cint(0)
					if (instr(strMailService, "Priority Mail Flat-Rate Envelope") OR strMailService="Priority Mail Flat Rate Envelope" OR instr(strMailService, "Flat Rate Envelope")) AND (USPS_OverSizeOrder=0) then
						strMailService = replace(strMailService,"&amp;lt;","<")
						strMailService = replace(strMailService,"&amp;gt;",">")
						strMailService = replace(strMailService,"&amp;","&")
						strMailService = replace(strMailService,"&lt;","<")
						strMailService = replace(strMailService,"&gt;",">")
						pcv_PM_MailService="USPS "&strMailService
						session("USPSshipStr")="|?|USPS|9901|"&"X|X|X|"
						if isNumeric(strRate) then
							strRate=cdbl(strRate)
						end if
						intUSPSPostage=intUSPSPostage+strRate
						iUSPSFlag=1
						intHasPriority = 1
					end if

					if (intHasPriority = 0) AND (USPS_OverSizeOrder=0) AND (instr(strMailService, "Priority Mail Flat-Rate Box") OR instr(strMailService, "Priority Mail Regular Flat-Rate Boxes") OR instr(strMailService, "Priority Mail Regular/Medium Flat-Rate Boxes") OR strMailService="Priority Mail Small Flat Rate Box"  OR strMailService="Priority Mail Medium Flat Rate Box"  OR strMailService="Priority Mail Large Flat Rate Box" OR instr(strMailService,"Flat Rate Box")) then
						strMailService = replace(strMailService,"&amp;lt;","<")
						strMailService = replace(strMailService,"&amp;gt;",">")
						strMailService = replace(strMailService,"&amp;","&")
						strMailService = replace(strMailService,"&lt;","<")
						strMailService = replace(strMailService,"&gt;",">")
						pcv_PM_MailService="USPS "&strMailService
						session("USPSshipStr")="|?|USPS|9901|"&"X|X|X|"
						if isNumeric(strRate) then
							strRate=cdbl(strRate)
						end if
						intUSPSPostage=intUSPSPostage+strRate
						iUSPSFlag=1
						intHasPriority = 1
					end if

					if (intHasPriority = 0) AND (USPS_TempSize="LARGE") AND (USPS_OverSizeOrder=0) then
						session("PMService")="LARGE"
						session("USPSshipStr")="|?|USPS|9901|"&"X|X|X|"
						if isNumeric(strRate) then
							strRate=cdbl(strRate)
						end if
						intUSPSPostage=intUSPSPostage+strRate
						iUSPSFlag=1
						intHasPriority = 1
					end if

					if (intHasPriority = 0) AND (instr(strMailService, "Priority Mail")) AND (USPS_OverSizeOrder=0) Then
						'Priority Mail&amp;lt;sup&amp;gt;&amp;amp;reg;&amp;lt;/sup&amp;gt;
						strMailService = replace(strMailService,"&amp;lt;","<")
						strMailService = replace(strMailService,"&amp;gt;",">")
						strMailService = replace(strMailService,"&amp;","&")
						strMailService = replace(strMailService,"&lt;","<")
						strMailService = replace(strMailService,"&gt;",">")
						pcv_PM_MailService="USPS "&strMailService
						session("USPSshipStr")="|?|USPS|9901|"&"X|X|X|"
						if isNumeric(strRate) then
							strRate=cdbl(strRate)
						end if
						intUSPSPostage=intUSPSPostage+strRate
						iUSPSFlag=1
						intHasPriority = 1
					end if
					intHasPriority = 0
				next
				
				if objLst.length > 0 then
					if iUSPSPMRate<>0 then
						intUSPSPostage=intUSPSPostage+iUSPSPMRate
					end if
					if session("PMService")="LARGE" then
						pcv_PM_MailService="USPS Priority Mail <sup>&reg;</sup>"
					end if
					availableShipStr=availableShipStr&replace(session("USPSshipStr"),"X|X|X|", pcv_PM_MailService)&"|"&intUSPSPostage&"|NA"
					session("USPSshipStr")=""
				end if
			end if
		end if
	end if 'size and weight are ok


	err.number=0


	'//USPS RATES - International %>
	<!--#include file="../includes/USPSCountry.asp"-->
	<% USPS_destination_country=USPSCountry(Universal_destination_country)
	
	intShowUSPSInternational = 0
	for q=1 to pcv_intPackageNum
		'//If any one side is greater 108 girth we skip USPS
		USPS_PackageSize=(Cint(session("USPSPackLength"&q)) + ((Cint(session("USPSPackWidth"&q))*2)+(Cint(session("USPSPackHeight"&q))*2)))
		If USPS_PackageSize>108 then
			intShowUSPSInternational = 1
		End if
	next
	
	
	if intShowUSPSInternational = 0 then
		
        
        '////////////////////////////////////////////////////////////////////////////
        '// START: USPS v1
        '////////////////////////////////////////////////////////////////////////////
		session("USPS_ShowGlobalRates")=""
		session("USPS_ShowExpressRates")=""
		session("USPS_ShowPriorityRates")=""
		session("USPS_ShowFirstClassRates")=""
	
		usps_postdata=""
		usps_postdata=usps_postdata&usps_server&"?API=IntlRate&XML="
	
		usps_postdata=usps_postdata&"<IntlRateRequest%20USERID="&chr(34)&usps_userid&chr(34)&">"
		for q=1 to pcv_intPackageNum
			'////////////////////////////////////
			'// Check Package Sizes for services
			'////////////////////////////////////
			'/ Get Dimensional Weight for Global Express
			pcv_USPS_Length=Cint(session("USPSPackLength"&q))
			pcv_USPS_Width=Cint(session("USPSPackWidth"&q))
			pcv_USPS_Height=Cint(session("USPSPackHeight"&q))
			pcv_USPS_DimWeight=((pcv_USPS_Length+pcv_USPS_Width+pcv_USPS_Height)/166)
	
			pcv_Decval = Mid(pcv_USPS_DimWeight, InStr(1, pcv_USPS_DimWeight, ".") + 1)
			pcv_DimWeightRound = CDbl(pcv_USPS_DimWeight)
			If pcv_Decval >= 0 Then
				 pcv_DimWeightRound = CInt(pcv_USPS_DimWeight)
				 pcv_DimWeightRound = pcv_DimWeightRound + 1
			End If
	
			if pcv_dimWeightRound>session("USPSPackPounds"&q) then
				'// Uncomment the following two line to use Dimensional Weight for USPS Global Express International Packages
				'session("USPSPackPounds"&q)=pcv_DimWeightRound
				'session("USPSPackOunces"&q)=0
			end if
	
			iNum=q-1

			usps_postdata=usps_postdata&"<Package%20ID="&chr(34)&iNum&chr(34)&">"
			usps_postdata=usps_postdata&"<Pounds>"&session("USPSPackPounds"&q)&"</Pounds>"
			usps_postdata=usps_postdata&"<Ounces>"&round(session("USPSPackOunces"&q))&"</Ounces>"
			usps_postdata=usps_postdata&"<MailType>Package</MailType>"
			'usps_postdata=usps_postdata&"<MailType>envelope</MailType>"
			if pcv_UseValueOfContents=1 then
				usps_postdata=usps_postdata&"<ValueOfContents>"&pcv_ValueOfContents&"</ValueOfContents>"
			end if           
			usps_postdata=usps_postdata&"<Country>"&USPS_destination_country&"</Country>" 

            If len(scShipFromPostalCode)>0 Then
            '    usps_postdata=usps_postdata&"<OriginZip>"&scShipFromPostalCode&"</OriginZip>"
            End If
            'usps_postdata=usps_postdata&"<DestinationPostalCode>"&Universal_destination_postal&"</DestinationPostalCode>"
			usps_postdata=usps_postdata&"</Package>"
	
			'if weight is over 70 lbs for any package, we do not show rates for Global Express
			if session("USPSPackPounds"&q)>70 then
				session("USPS_ShowGlobalRates")="NO"
			end if
			if pcv_USPS_Length>46 OR pcv_USPS_Width>46 OR pcv_USPS_Height>46 then
				session("USPS_ShowGlobalRates")="NO"
			end if
			'If Express demension of one side exceeds 36, don't show rates
			if session("USPSPackPounds"&q)>66 then
				session("USPS_ShowExpressRates")="NO"
			end if
			if pcv_USPS_Length>60 OR pcv_USPS_Width>60 OR pcv_USPS_Height>60 then
				session("USPS_ShowExpressRates")="NO"
			end if
			'If Priority demension of one side exceeds 36, don't show rates
			if session("USPSPackPounds"&q)>66 then
				session("USPS_ShowPriorityRates")="NO"
			end if
			if pcv_USPS_Length>60 OR pcv_USPS_Width>60 OR pcv_USPS_Height>60 then
				session("USPS_ShowPriorityRates")="NO"
			end if
			'if First Class is over 4 pounds
			if session("USPSPackPounds"&q)>4 then
				session("USPS_ShowFirstClassRates")="NO"
			end if
			if pcv_USPS_Length>24 OR pcv_USPS_Width>24 OR pcv_USPS_Height>24 then
				session("USPS_ShowFirstClassRates")="NO"
			end if
		next
		usps_postdata=usps_postdata&"</IntlRateRequest>"
        
        'Response.Clear()
        'Response.ContentType = "text/xml"
        'Response.Write srvUSPSINTXmlHttp.responseText
        'Response.End()
	
		Set srvUSPSINTXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
		srvUSPSINTXmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
		srvUSPSINTXmlHttp.open "GET", usps_postdata, false
		srvUSPSINTXmlHttp.send
        
        'Response.Clear()
        'Response.ContentType = "text/xml"
        'Response.Write srvUSPSINTXmlHttp.responseText
        'Response.End()
	
		USPSINT_result = srvUSPSINTXmlHttp.responseText
		Set USPSINTXMLdoc = server.CreateObject("Msxml2.DOMDocument"&scXML)
		USPSINTXMLDoc.async = false
	
		if USPSINTXMLDoc.loadXML(USPSINT_result) then ' if loading from a string
			USPSA=0
			USPSA_1=Ccur(0)
			USPSA_2=""
			USPSB=0
			USPSB_1=Ccur(0)
			USPSB_2=""
			USPSC=0
			USPSC_1=Ccur(0)
			USPSC_2=""
			USPSD=0
			USPSD_1=Ccur(0)
			USPSD_2=""
			USPSE=0
			USPSE_1=Ccur(0)
			USPSE_2=""
			USPSF=0
			USPSF_1=Ccur(0)
			USPSF_2=""
			USPSG=0
			USPSG_1=Ccur(0)
			USPSG_2=""
			USPSH=0
			USPSH_1=Ccur(0)
			USPSH_2=""
			USPSI=0
			USPSI_1=Ccur(0)
			USPSI_2=""
			USPSJ=0
			USPSJ_1=Ccur(0)
			USPSJ_2=""
	
			set objLst=USPSINTXMLDoc.getElementsByTagName("Package")
	
			for i = 0 to (objLst.length - 1)
				for j=0 to ((objLst.item(i).childNodes.length)-1)
					If objLst.item(i).childNodes(j).nodeName="Service" then
						intCLASSID=objLst.item(i).childNodes(j).getAttribute("ID")
						usps_int_1="0"
						usps_int_2="0"
						usps_int_3="0"
						for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="SvcCommitments" then
								usps_int_2 = objLst.item(i).childNodes(j).childNodes(m).text
							end if
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="Postage" then
								usps_int_1 = objLst.item(i).childNodes(j).childNodes(m).text
							end if
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="Insurance" then
								usps_int_3 = objLst.item(i).childNodes(j).childNodes(m).text
							end if
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="SvcDescription" then
								serviceVar = objLst.item(i).childNodes(j).childNodes(m).text
							end if
						Next

						select case intCLASSID

						case "4", "Global Express Guaranteed", "Global Express Guaranteed (GXG)"
							if session("USPS_ShowGlobalRates")="" then
								USPSA=1
								USPSA_1=USPSA_1+ccur(usps_int_1)
								USPSA_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSA_1=USPSA_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "6", "Global Express Guaranteed Non-Document Rectangular"
							if session("USPS_ShowGlobalRates")="" then
								USPSB=1
								USPSB_1=USPSB_1+ccur(usps_int_1)
								USPSB_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSB_1=USPSB_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "7", "Global Express Guaranteed Non-Document Non-Rectangular"
							if session("USPS_ShowGlobalRates")="" then
								USPSC=1
								USPSC_1=USPSC_1+ccur(usps_int_1)
								USPSC_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSC_1=USPSC_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "1", "Express Mail International (EMS)", "Express Mail International"
							'// Moved to v2
						case "10", "Express Mail International (EMS) Flat Rate Envelope", "Express Mail International Flat Rate Envelope"
							if session("USPS_ShowExpressRates")="" then
								USPSE=1
								USPSE_1=USPSE_1+ccur(usps_int_1)
								USPSE_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSE_1=USPSE_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "2", "Priority Mail International"
							'// Moved to v2
						case "8", "Priority Mail International Flat Rate Envelope"
							if session("USPS_ShowPriorityRates")="" then
								USPSG=1
								USPSG_1=USPSG_1+ccur(usps_int_1)
								USPSG_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSG_1=USPSG_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "9", "Priority Mail International Medium Flat Rate Box"
							if session("USPS_ShowPriorityRates")="" then
								USPSH=1
								USPSH_1=USPSH_1+ccur(usps_int_1)
								USPSH_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSH_1=USPSH_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "15", "First-Class Mail International", "First Class Mail International Package", "First-Class Mail International Package"
							'// Moved to v2
						end select
					End If
				Next
			Next
	
			'//COMPILE availableShipStr
	
            If objLst.length > 0 AND USPS_OverSizeOrder=0 then
                if USPSA=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9914|"&"Global Express Guaranteed<sup>&reg;</sup>|"&USPSA_1&"|"&Trim(replace(USPSA_2,CHR(10),""))
                end if
                if USPSB=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9905|"&"Global Express Guaranteed<sup>&reg;</sup> Non-Document Rectangular|"&USPSB_1&"|"&Trim(replace(USPSB_2,CHR(10),""))
                end if
                if USPSC=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9910|"&"Global Express Guaranteed<sup>&reg;</sup> Non-Document Non-Rectangular|"&USPSC_1&"|"&Trim(replace(USPSC_2,CHR(10),""))
                end if
                if USPSD=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9906|"&"Express Mail<sup>&reg;</sup> International (EMS)|"&USPSD_1&"|"&Trim(replace(USPSD_2,CHR(10),""))
                end if
                if USPSE=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9911|"&"Express Mail<sup>&reg;</sup> International (EMS) Flat Rate Envelope|"&USPSE_1&"|"&Trim(replace(USPSE_2,CHR(10),""))
                end if
                if USPSF=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9907|"&"Priority Mail<sup>&reg;</sup> International|"&USPSF_1&"|"&Trim(replace(USPSF_2,CHR(10),""))
                end if
                if USPSG=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9908|"&"Priority Mail<sup>&reg;</sup> International Flat Rate Envelope|"&USPSG_1&"|"&Trim(replace(USPSG_2,CHR(10),""))
                end if
                if USPSH=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9909|"&"Priority Mail<sup>&reg;</sup> International Flat Rate Box|"&USPSH_1&"|"&Trim(replace(USPSH_2,CHR(10),""))
                end if
                if USPSI=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9912|"&"First-Class Mail<sup>&reg;</sup> International|"&USPSI_1&"|"&Trim(replace(USPSI_2,CHR(10),""))
                end if
            End if
	
		end if
        '////////////////////////////////////////////////////////////////////////////
        '// END: USPS v1
        '////////////////////////////////////////////////////////////////////////////
        
        '////////////////////////////////////////////////////////////////////////////
        '// START: USPS v2
        '////////////////////////////////////////////////////////////////////////////
		session("USPS_ShowGlobalRates")=""
		session("USPS_ShowExpressRates")=""
		session("USPS_ShowPriorityRates")=""
		session("USPS_ShowFirstClassRates")=""
	
		usps_postdata=""
		usps_postdata=usps_postdata&usps_server&"?API=IntlRateV2&XML="
	
		usps_postdata=usps_postdata&"<IntlRateV2Request%20USERID="&chr(34)&usps_userid&chr(34)&">"
		for q=1 to pcv_intPackageNum
			'////////////////////////////////////
			'// Check Package Sizes for services
			'////////////////////////////////////
			'/ Get Dimensional Weight for Global Express
			pcv_USPS_Length=Cint(session("USPSPackLength"&q))
			pcv_USPS_Width=Cint(session("USPSPackWidth"&q))
			pcv_USPS_Height=Cint(session("USPSPackHeight"&q))
			pcv_USPS_DimWeight=((pcv_USPS_Length+pcv_USPS_Width+pcv_USPS_Height)/166)
	
			pcv_Decval = Mid(pcv_USPS_DimWeight, InStr(1, pcv_USPS_DimWeight, ".") + 1)
			pcv_DimWeightRound = CDbl(pcv_USPS_DimWeight)
			If pcv_Decval >= 0 Then
				 pcv_DimWeightRound = CInt(pcv_USPS_DimWeight)
				 pcv_DimWeightRound = pcv_DimWeightRound + 1
			End If
	
			if pcv_dimWeightRound>session("USPSPackPounds"&q) then
				'// Uncomment the following two line to use Dimensional Weight for USPS Global Express International Packages
				'session("USPSPackPounds"&q)=pcv_DimWeightRound
				'session("USPSPackOunces"&q)=0
			end if
	
			iNum=q-1
            
            
            USPS_PackageSizeC=""
            '//If any one side is greater then 12" package is labeled as "LARGE"
            If (Cint(session("USPSPackLength"&q))>12) OR (Cint(session("USPSPackWidth"&q))>12) OR (Cint(session("USPSPackHeight"&q))>12) Then
                USPS_PackageSizeC="LARGE"
            Else
                USPS_PackageSize=(Cint(session("USPSPackLength"&q)) + ((Cint(session("USPSPackWidth"&q))*2)+(Cint(session("USPSPackHeight"&q))*2)))
                If USPS_PackageSize<85 then
                    USPS_PackageSizeC="REGULAR"
                End if
                If USPS_PackageSize>85 AND USPS_PackageSize<108 AND intPounds<15 then
                    USPS_PackageSizeC="LARGE"
                End if
                If USPS_PackageSizeC="" AND USPS_PackageSize<131 then
                    USPS_PackageSizeC="OVERSIZE"
                End if
            End If
            '//Eliminate small box if any side is over 3 inchdes

            intshowSmallBox = 1
            intShowMedBox = 1

            LengthIsCal = 0
            A = Cint(session("USPSPackLength"&q))
            B = Cint(session("USPSPackWidth"&q))
            C = Cint(session("USPSPackHeight"&q))
            IntLongestLength = Cint(0)
            IntMidLength = Cint(0)
            IntShortestLength = Cint(0)

            If A=>B AND A=>C Then
                'A is the longest
                IntLongestLength = A
                LengthIsCal = 1
                If B=>C Then
                    'B is the mid
                    IntMidLength = B
                    'C is the shortest
                    IntShortestLength = C
                Else
                    'C is the mid
                    IntMidLength = C
                    'B is the shortest
                    IntShortestLength = B
                End If
            End If

            If (B=>A AND B=>C) AND (LengthIsCal = 0) Then
                'B is the longest
                IntLongestLength = B
                LengthIsCal = 1
                If A=>C Then
                    'A is the mid
                    IntMidLength = A
                    'C is the shortest
                    IntShortestLength = C
                Else
                    'C is the mid
                    IntMidLength = C
                    'A is the shortest
                    IntShortestLength = A
                End If
            End If

            If (C=>A AND C=>B) AND (LengthIsCal = 0) Then
                'C is the longest
                IntLongestLength = C
                LengthIsCal = 1
                If B=>A Then
                    'B is the mid
                    IntMidLength = B
                    'A is the shortest
                    IntShortestLength = A
                Else
                    'A is the mid
                    IntMidLength = A
                    'B is the shortest
                    IntShortestLength = B
                End If
            End If

            If IntShortestLength=>5.50 Then
                tUSPS_PM_PACKAGE = "NONE"
                pcv_PM_Null=0
            Else
                tUSPS_PM_PACKAGE=USPS_PM_PACKAGE
            End If

            If tUSPS_PM_PACKAGE <> "NONE" AND IntShortestLength=>1.5 Then
                intShowSmallBox = 0
                pcv_PMOption="Flat Rate Box1"
            End If
            If pcv_PMOption="Flat Rate Box1" AND IntShortestLength=>3.5 Then
                intShowMedBox = 0
                pcv_PMOption="Flat Rate Box2"
            End If
            
            usps_postdata=usps_postdata&"<Revision>2</Revision>"
			usps_postdata=usps_postdata&"<Package%20ID="&chr(34)&iNum&chr(34)&">"
			usps_postdata=usps_postdata&"<Pounds>"&session("USPSPackPounds"&q)&"</Pounds>"
			usps_postdata=usps_postdata&"<Ounces>"&round(session("USPSPackOunces"&q))&"</Ounces>"
			usps_postdata=usps_postdata&"<MailType>All</MailType>"
			'if pcv_UseValueOfContents=1 then
				usps_postdata=usps_postdata&"<ValueOfContents>"&pcv_ValueOfContents&"</ValueOfContents>"
			'end if            
			usps_postdata=usps_postdata&"<Country>"&USPS_destination_country&"</Country>" 
            
            
            '// If FRE is the default, check for weight limit and alternate container
            if pcv_PM_Null=1 AND pcv_PMOption<>"0" then
                if pcv_PMOption="Flat Rate Box" AND intshowSmallBox = 1 then
                    pcv_PMOption = "Sm Flat Rate Box"
                end if
                if pcv_PMOption="Flat Rate Box1" AND intShowMedBox = 1 then
                    pcv_PMOption = "Md Flat Rate Box"
                end if
                if pcv_PMOption="Flat Rate Box2" then
                    pcv_PMOption = "Lg Flat Rate Box"
                end if
                tUSPS_PM_PACKAGE=pcv_PMOption
            end if

            'private const USPS_PM_FREOption="0"
            if USPS_PM_PACKAGE<>"Flat Rate Envelope" then
                if tUSPS_PM_PACKAGE <> "NONE" then
                    'pcv_PMOption=USPS_PM_PACKAGE
                    if pcv_PMOption="Flat Rate Box" AND intshowSmallBox = 1 then
                        pcv_PMOption = "Sm Flat Rate Box"
                    end if
                    if pcv_PMOption="Flat Rate Box1" AND intShowMedBox = 1 then
                        pcv_PMOption = "Md Flat Rate Box"
                    end if
                    if pcv_PMOption="Flat Rate Box2" then
                        pcv_PMOption = "Lg Flat Rate Box"
                    end if
                    tUSPS_PM_PACKAGE=pcv_PMOption
                else
                    if ucase(tUSPS_PM_PACKAGE)="NONE" then
                        'check for Priority Mail totals from previous
                        if iUSPSPMRate=0 then
                            tUSPS_PM_PACKAGE="NONRECTANGULAR"
                        end if
                    end if
                end if
            end if
            
            if USPS_PackageSizeC="REGULAR" AND ucase(tUSPS_PM_PACKAGE)="NONE" then
                tUSPS_PM_PACKAGE="NONRECTANGULAR"
            end if     
                   
            IF USPS_PackageSizeC="LARGE" then
                tUSPS_PM_PACKAGE="RECTANGULAR"
            end if
            
            
            usps_postdata=usps_postdata&"<Container>"&tUSPS_PM_PACKAGE&"</Container>"
            usps_postdata=usps_postdata&"<Size>"&USPS_PackageSizeC&"</Size>"

            'IF USPS_PackageSizeC="LARGE" THEN
                usps_postdata=usps_postdata&"<Width>"&session("USPSPackWidth"&q)&"</Width>"
                usps_postdata=usps_postdata&"<Length>"&session("USPSPackLength"&q)&"</Length>"
                usps_postdata=usps_postdata&"<Height>"&session("USPSPackHeight"&q)&"</Height>"
                usps_postdata=usps_postdata&"<Girth>"&USPS_PackageSize&"</Girth>"
            'END IF
            
            If len(scShipFromPostalCode)>0 Then
                usps_postdata=usps_postdata&"<OriginZip>"&scShipFromPostalCode&"</OriginZip>"
            End If
            'usps_postdata=usps_postdata&"<DestinationPostalCode>"&Universal_destination_postal&"</DestinationPostalCode>"
			usps_postdata=usps_postdata&"</Package>"
	
			'if weight is over 70 lbs for any package, we do not show rates for Global Express
			if session("USPSPackPounds"&q)>70 then
				session("USPS_ShowGlobalRates")="NO"
			end if
			if pcv_USPS_Length>46 OR pcv_USPS_Width>46 OR pcv_USPS_Height>46 then
				session("USPS_ShowGlobalRates")="NO"
			end if
			'If Express demension of one side exceeds 36, don't show rates
			if session("USPSPackPounds"&q)>66 then
				session("USPS_ShowExpressRates")="NO"
			end if
			if pcv_USPS_Length>60 OR pcv_USPS_Width>60 OR pcv_USPS_Height>60 then
				session("USPS_ShowExpressRates")="NO"
			end if
			'If Priority demension of one side exceeds 36, don't show rates
			if session("USPSPackPounds"&q)>66 then
				session("USPS_ShowPriorityRates")="NO"
			end if
			if pcv_USPS_Length>60 OR pcv_USPS_Width>60 OR pcv_USPS_Height>60 then
				session("USPS_ShowPriorityRates")="NO"
			end if
			'if First Class is over 4 pounds
			if session("USPSPackPounds"&q)>4 then
				session("USPS_ShowFirstClassRates")="NO"
			end if
			if pcv_USPS_Length>24 OR pcv_USPS_Width>24 OR pcv_USPS_Height>24 then
				session("USPS_ShowFirstClassRates")="NO"
			end if
		next
		usps_postdata=usps_postdata&"</IntlRateV2Request>"
        
        'Response.Clear()
        'Response.ContentType = "text/xml"
        'Response.Write usps_postdata
        'Response.End() 
        
		Set srvUSPSINTXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
		srvUSPSINTXmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
		srvUSPSINTXmlHttp.open "GET", usps_postdata, false
		srvUSPSINTXmlHttp.send
        
        'Response.Clear()
        'Response.ContentType = "text/xml"
        'Response.Write srvUSPSINTXmlHttp.responseText
        'Response.End()
	
		USPSINT_result = srvUSPSINTXmlHttp.responseText
		Set USPSINTXMLdoc = server.CreateObject("Msxml2.DOMDocument"&scXML)
		USPSINTXMLDoc.async = false
	
		if USPSINTXMLDoc.loadXML(USPSINT_result) then ' if loading from a string
			USPSA=0
			USPSA_1=Ccur(0)
			USPSA_2=""
			USPSB=0
			USPSB_1=Ccur(0)
			USPSB_2=""
			USPSC=0
			USPSC_1=Ccur(0)
			USPSC_2=""
			USPSD=0
			USPSD_1=Ccur(0)
			USPSD_2=""
			USPSE=0
			USPSE_1=Ccur(0)
			USPSE_2=""
			USPSF=0
			USPSF_1=Ccur(0)
			USPSF_2=""
			USPSG=0
			USPSG_1=Ccur(0)
			USPSG_2=""
			USPSH=0
			USPSH_1=Ccur(0)
			USPSH_2=""
			USPSI=0
			USPSI_1=Ccur(0)
			USPSI_2=""
			USPSJ=0
			USPSJ_1=Ccur(0)
			USPSJ_2=""
	
			set objLst=USPSINTXMLDoc.getElementsByTagName("Package")
	
			for i = 0 to (objLst.length - 1)
				for j=0 to ((objLst.item(i).childNodes.length)-1)
					If objLst.item(i).childNodes(j).nodeName="Service" then
						intCLASSID=objLst.item(i).childNodes(j).getAttribute("ID")
						usps_int_1="0"
						usps_int_2="0"
						usps_int_3="0"
						for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="SvcCommitments" then
								usps_int_2 = objLst.item(i).childNodes(j).childNodes(m).text
							end if
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="Postage" then
								usps_int_1 = objLst.item(i).childNodes(j).childNodes(m).text
							end if
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="Insurance" then
								usps_int_3 = objLst.item(i).childNodes(j).childNodes(m).text
							end if
							If objLst.item(i).childNodes(j).childNodes(m).nodeName="SvcDescription" then
								serviceVar = objLst.item(i).childNodes(j).childNodes(m).text
							end if
						Next

						select case intCLASSID

						case "4", "Global Express Guaranteed", "Global Express Guaranteed (GXG)"
							'// Only v1
						case "6", "Global Express Guaranteed Non-Document Rectangular"
							'// Only v1
						case "7", "Global Express Guaranteed Non-Document Non-Rectangular"
							'// Only v1
						case "1", "Express Mail International (EMS)", "Express Mail International"
							if session("USPS_ShowExpressRates")="" then
								USPSD=1
								USPSD_1=USPSD_1+ccur(usps_int_1)
								USPSD_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSD_1=USPSD_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						case "10", "Express Mail International (EMS) Flat Rate Envelope", "Express Mail International Flat Rate Envelope"
							'// Only v1
						case "2", "Priority Mail International"
							if session("USPS_ShowPriorityRates")="" then
								If ccur(usps_int_1)>0 Then
                                    USPSF=1
                                    USPSF_1=USPSF_1+ccur(usps_int_1)
                                    USPSF_2=Trim(replace(usps_int_2,CHR(10),""))
                                    if isNumeric(usps_int_3) then
                                        usps_int_3=cdbl(usps_int_3)
                                        USPSF_1=USPSF_1+usps_int_3
                                    end if
                                    iUSPSFlag=1
                                End If
							end if
						case "8", "Priority Mail International Flat Rate Envelope"
							'// Only v1
						case "9", "Priority Mail International Medium Flat Rate Box"
							'// Only v1
						case "15", "First-Class Mail International", "First Class Mail International Package", "First-Class Mail International Package"
							if session("USPS_ShowFirstClassRates")="" then
								USPSI=1
								USPSI_1=USPSI_1+ccur(usps_int_1)
								USPSI_2=Trim(replace(usps_int_2,CHR(10),""))
								if isNumeric(usps_int_3) then
									usps_int_3=cdbl(usps_int_3)
									USPSI_1=USPSI_1+usps_int_3
								end if
								iUSPSFlag=1
							end if
						end select
					End If
				Next
			Next
	
			'//COMPILE availableShipStr
	
            If objLst.length > 0 AND USPS_OverSizeOrder=0 then
                if USPSA=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9914|"&"Global Express Guaranteed<sup>&reg;</sup>|"&USPSA_1&"|"&Trim(replace(USPSA_2,CHR(10),""))
                end if
                if USPSB=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9905|"&"Global Express Guaranteed<sup>&reg;</sup> Non-Document Rectangular|"&USPSB_1&"|"&Trim(replace(USPSB_2,CHR(10),""))
                end if
                if USPSC=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9910|"&"Global Express Guaranteed<sup>&reg;</sup> Non-Document Non-Rectangular|"&USPSC_1&"|"&Trim(replace(USPSC_2,CHR(10),""))
                end if
                if USPSD=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9906|"&"Express Mail<sup>&reg;</sup> International (EMS)|"&USPSD_1&"|"&Trim(replace(USPSD_2,CHR(10),""))
                end if
                if USPSE=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9911|"&"Express Mail<sup>&reg;</sup> International (EMS) Flat Rate Envelope|"&USPSE_1&"|"&Trim(replace(USPSE_2,CHR(10),""))
                end if
                if USPSF=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9907|"&"Priority Mail<sup>&reg;</sup> International|"&USPSF_1&"|"&Trim(replace(USPSF_2,CHR(10),""))
                end if
                if USPSG=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9908|"&"Priority Mail<sup>&reg;</sup> International Flat Rate Envelope|"&USPSG_1&"|"&Trim(replace(USPSG_2,CHR(10),""))
                end if
                if USPSH=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9909|"&"Priority Mail<sup>&reg;</sup> International Flat Rate Box|"&USPSH_1&"|"&Trim(replace(USPSH_2,CHR(10),""))
                end if
                if USPSI=1 then
                    availableShipStr=availableShipStr&"|?|USPS|9912|"&"First-Class Mail<sup>&reg;</sup> International|"&USPSI_1&"|"&Trim(replace(USPSI_2,CHR(10),""))
                end if
            End if
	
		end if
        '////////////////////////////////////////////////////////////////////////////
        '// END: USPS v2
        '////////////////////////////////////////////////////////////////////////////


	end if
end if 'if usps is active

if iUSPSFlag=1 then
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Pre-Filter USPS for Availability
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	pcv_strOptionFilterPass=pcf_PreFilter("USPS", availableShipStr)
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Pre-Filter USPS for Availability
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if (pcv_strOptionFilterPass=-1) AND (USPS_OverSizeOrder=0) then
		strDefaultProvider="USPS"
		iShipmentTypeCnt=iShipmentTypeCnt+1
		strOptionShipmentType=strOptionShipmentType&"<option value=USPS>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_l")&"</option>"
		strTabShipmentType=strTabShipmentType&"[TAB]USPS,"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_l")&"[/TAB]"
	end if
end if

If CP_active=true or CP_active="-1" then
	iCPActive=1
	'//Canada Post
	'compile xml
	CP_postdata=""
	CP_postdata=CP_postdata&"<?xml version=""1.0"" encoding=""UTF-8""?>"
	CP_postdata=CP_postdata&"<mailing-scenario xmlns=""http://www.canadapost.ca/ws/ship/rate-v3"">"
	CP_postdata=CP_postdata&"<customer-number>"&CP_custNo&"</customer-number>"
	CP_postdata=CP_postdata&"<parcel-characteristics>"
	
	for q=1 to pcv_intPackageNum
        tmpPounds = session("CPPackPounds"&q)
        tmpOunces = session("CPPackOunces"&q)
        if NOT isNumeric(tmpPounds) then
            tmpPounds = 0
        end if
        if NOT isNumeric(tmpOunces) then
            tmpOunces = 0
        end if
        If scShipFromWeightUnit="LBS" Then
            tmpOuncesDec = CDbl(tmpOunces)/16
            pcv_intCPWeight = Round(((tmpPounds + tmpOuncesDec) * 0.4536), 2)
        Else
            tmpOuncesDec = tmpOunces
            pcv_intCPWeight = tmpPounds + tmpOuncesDec
        End If
        
		CP_postdata=CP_postdata&"<weight>"& pcv_intCPWeight &"</weight>"
		CP_postdata=CP_postdata&"<dimensions>"
		CP_postdata=CP_postdata&"<length>"&session("CPPackLength"&q)&"</length>"
		CP_postdata=CP_postdata&"<width>"&session("CPPackWidth"&q)&"</width>"
		CP_postdata=CP_postdata&"<height>"&session("CPPackHeight"&q)&"</height>"
		CP_postdata=CP_postdata&"</dimensions>"
	next
	CP_postdata=CP_postdata&"</parcel-characteristics>"
	CP_postdata=CP_postdata&"<origin-postal-code>"&replace(scShipFromPostalCode," ","")&"</origin-postal-code>"
	CP_postdata=CP_postdata&"<destination>"
	
	if Universal_destination_country="CA" then
		CP_postdata=CP_postdata&"<domestic>"
		CP_postdata=CP_postdata&"<postal-code>"&replace(Universal_destination_postal," ","")&"</postal-code>"
		CP_postdata=CP_postdata&"</domestic>"
	elseif Universal_destination_country="US" then
		CP_postdata=CP_postdata&"<united-states>"
		CP_postdata=CP_postdata&"<zip-code>"&replace(Universal_destination_postal," ","")&"</zip-code>"
		CP_postdata=CP_postdata&"</united-states>"
	else
		CP_postdata=CP_postdata&"<international>"
		CP_postdata=CP_postdata&"<country-code>"&Universal_destination_country&"</country-code>"
		CP_postdata=CP_postdata&"</international>"
	end if
	CP_postdata=CP_postdata&"</destination>"
	CP_postdata=CP_postdata&"</mailing-scenario>"
    
    'Response.Clear()
    'Response.ContentType = "text/xml"
    'Response.Write CP_postdata
    'Response.End() 

	Set srvCPXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	srvCPXmlHttp.open "POST", CP_server, false
	srvCPXmlHttp.SetRequestHeader "Content-Type", "application/vnd.cpc.ship.rate-v3+xml"
	srvCPXmlHttp.SetRequestHeader "Accept", "application/vnd.cpc.ship.rate-v3+xml"
	srvCPXmlHttp.SetRequestHeader "Authorization", "Basic " & Base64_Encode(CP_username & ":" & CP_password)
	srvCPXmlHttp.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
	srvCPXmlHttp.send(CP_postdata)

	CP_result = srvCPXmlHttp.responseText
    
    'Response.Clear()
    'Response.ContentType = "text/xml"
    'Response.Write CP_result
    'Response.End() 
	
	Set CPXMLdoc = server.CreateObject("Msxml2.DOMDocument"&scXML)
	CPXMLDoc.async = false
	if CPXMLDoc.loadXML(CP_result) then ' if loading from a string
		set objLst = CPXMLDoc.getElementsByTagName("price-quote")
		for i = 0 to (objLst.length - 1)
			varFlag=0
			for j=0 to ((objLst.item(i).childNodes.length)-1)
				If objLst.item(i).childNodes(j).nodeName="service-code" then
					CP_ID=objLst.item(i).childNodes(j).text
					select case CP_ID

					case "DOM.RP"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Regular Parcel"
						varFlag=1
						iCPFlag=1
					case "DOM.EP"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Expedited Parcel"
						varFlag=1
						iCPFlag=1
					case "DOM.XP"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Xpresspost"
						varFlag=1
						iCPFlag=1
					case "DOM.XP.CERT"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Xpresspost Certified"
						varFlag=1
						iCPFlag=1
					case "DOM.PC"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Priority"
						varFlag=1
						iCPFlag=1
					case "DOM.DT"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Delivered Tonight"
						varFlag=1
						iCPFlag=1
					case "DOM.LIB"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Library Materials"
						varFlag=1
						iCPFlag=1
					case "USA.EP"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Expedited Parcel USA"
						varFlag=1
						iCPFlag=1
					case "USA.PW.ENV"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Priority Worldwide Envelope USA"
						varFlag=1
						iCPFlag=1
					case "USA.PW.PAK"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Priority Worldwide Pak USA"
						varFlag=1
						iCPFlag=1
					case "USA.PW.PARCEL"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Priority Worldwide Parcel USA"
						varFlag=1
						iCPFlag=1
					case "USA.SP.AIR"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Small Packet USA Air"
						varFlag=1
						iCPFlag=1
					case "USA.TP"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Tracked Packet – USA"
						varFlag=1
						iCPFlag=1
					case "USA.TP.LVM"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Tracked Packet – USA (Large Volume Mailers)"
						varFlag=1
						iCPFlag=1
					case "USA.XP"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Xpresspost USA"
						varFlag=1
						iCPFlag=1
					case "INT.XP"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Xpresspost International"
						varFlag=1
						iCPFlag=1
					case "INT.IP.AIR"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - International Parcel Air"
						varFlag=1
						iCPFlag=1
					case "INT.IP.SURF"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - International Parcel Surface"
						varFlag=1
						iCPFlag=1
					case "INT.PW.ENV"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Priority Worldwide Envelope International"
						varFlag=1
						iCPFlag=1
					case "INT.PW.PAK"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Priority Worldwide Pak International"
						varFlag=1
						iCPFlag=1
					case "INT.PW.PARCEL"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Priority Worldwide Parcel International"
						varFlag=1
						iCPFlag=1
					case "INT.SP.AIR"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Small Packet International Air"
						varFlag=1
						iCPFlag=1
					case "INT.SP.SURF"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Small Packet International Surface"
						varFlag=1
						iCPFlag=1
					case "INT.TP"
						availableShipStr=availableShipStr&"|?|CP|"&CP_ID&"|"&"Canada Post - Tracked Packet – International"
						varFlag=1
						iCPFlag=1

					end select

				End if
				
				If objLst.item(i).childNodes(j).nodeName="price-details" AND varFlag=1 then
					availableShipStr=availableShipStr&"|"&objLst.item(i).childNodes(j).childNodes(2).text
				End if

				If objLst.item(i).childNodes(j).nodeName="service-standard" AND varFlag=1 then
					
					for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
						if objLst.item(i).childNodes(j).childNodes(k).nodeName="expected-delivery-date" then
							deliveryDate=objLst.item(i).childNodes(j).childNodes(k).text
						end if
					next
					
					if deliveryDate<>"" then
						deliveryDateArry=split(deliveryDate,"-")
						deliveryDateMonth=deliveryDateArry(1)
						deliveryDateYear=deliveryDateArry(0)
						deliveryDateDay=deliveryDateArry(2)
						deliveryDateFrmt=(deliveryDateMonth&"/"&deliveryDateDay&"/"&deliveryDateYear)
						
						shippingDateMonth=month(Date)
						shippingDateYear=year(Date)
						shippingDateDay=day(Date)
						shippingDateFrmt=(shippingDateMonth&"/"&shippingDateDay&"/"&shippingDateYear)
						
						DeliveryDays=DateDiff("d",shippingDateFrmt,deliveryDateFrmt)
						availableShipStr=availableShipStr&"|"&DeliveryDays&" Days"
					else
						availableShipStr=availableShipStr&"|NA"
					end if
					
				End if
			next
		next
	end if
end if 'if canada post is active

if iCPFlag=1 then
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Pre-Filter CP for Availability
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	pcv_strOptionFilterPass=pcf_PreFilter("CP", availableShipStr)
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Pre-Filter CP for Availability
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if pcv_strOptionFilterPass=-1 then
		strDefaultProvider="CP"
		iShipmentTypeCnt=iShipmentTypeCnt+1
		strOptionShipmentType=strOptionShipmentType&"<option value=CP>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_n")&"</option>"
		strTabShipmentType=strTabShipmentType&"[TAB]CP,"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_n")&"[/TAB]"
	end if
end if

'//Custom Rates
ERR.NUMBER=0

mySQL="SELECT idFlatShiptype,WQP,FlatShipTypeDesc,FlatShipTypeDelivery,startIncrement FROM FlatShipTypes"
set rstemp=conntemp.execute(mySQL)

Do until rstemp.eof
	ifound=0
	idTemp=rstemp("idFlatShiptype")
	VarWQP=trim(rstemp("WQP"))
	
	CustomShipDelivery=rstemp("FlatShipTypeDelivery")
	if CustomShipDelivery="" then
		CustomShipDelivery="NA"
	end if

	If len(VarWQP)>0 Then

		select case VarWQP
			case "W"
			  mySQL2="SELECT shippingPrice FROM FlatShipTypeRules WHERE idFlatShipType="& idTemp &" AND quantityTo>=" &intCustomShipWeight& " AND quantityFrom<=" &intCustomShipWeight
			case "Q"
			  mySQL2="SELECT shippingPrice FROM FlatShipTypeRules WHERE idFlatShipType="& idTemp &" AND quantityTo>=" &pCartShipQuantity & " AND quantityFrom<=" & pCartShipQuantity
			case "P"
			  mySQL2="SELECT shippingPrice FROM FlatShipTypeRules WHERE idFlatShipType="& idTemp &" AND quantityTo>=" & pShipSubTotal & " AND quantityFrom<=" & pShipSubTotal
			case "O"
			  mySQL2="SELECT shippingPrice FROM FlatShipTypeRules WHERE idFlatShipType="& idTemp &" AND quantityTo>=" & pShipSubTotal & " AND quantityFrom<=" & pShipSubTotal
			case "I"
			  if pCartShipQuantity=1 then
				  pCartShipQuantity2=2
			  else
				  pCartShipQuantity2=pCartShipQuantity
			  end if
			  mySQL2="SELECT shippingPrice, quantityTo FROM FlatShipTypeRules WHERE idFlatShipType="& idTemp &" AND quantityFrom<=" & pCartShipQuantity2
		end select
		
		set rsShipObj=conntemp.execute(mySQL2)
		
		if NOT rsShipObj.eof then
			ifound=1
			tempShipPrice=rsShipObj("shippingPrice")
			availableShipStr=availableShipStr&"|?|CUSTOM|C"&idTemp&"|"&rstemp("FlatShipTypeDesc")
			iCustomFlag=1
			
			'calculate shipping price for I and O
			if VarWQP="O" then
				'shipping price is the percentage
				tempPercentage=tempShipPrice
				tempShipPrice=((tempPercentage/100)*pShipSubTotal)
			end if
			
			if VarWQP="I" then
				dim iRegPrice, iAddRegPrice
				mySQL3="SELECT startIncrement FROM FlatShipTypes WHERE startIncrement>0 AND idFlatShipType="&idTemp
				set rsIncretObj=conntemp.execute(mySQL3)
				TempShipPrice=rsIncretObj("startIncrement")
				IShipCnt=pCartShipQuantity
				iTempCompleted=0
				query="SELECT quantityFrom, quantityTo, shippingPrice FROM FlatShipTypeRules WHERE (((FlatShipTypeRules.idFlatshipType)="&idTemp&")) ORDER BY FlatShipTypeRules.num;"
				set rsIncretObj=conntemp.execute(query)
				Do until rsIncretObj.eof or iTempCompleted=1
					iQuantityFrom=rsIncretObj("quantityFrom")
					iQuantityTo=rsIncretObj("quantityTo")
					AddPrice=rsIncretObj("shippingPrice")
			
					if ccur(IShipCnt) - ccur(iQuantityTo) => 0 then
						TierCnt = (ccur(iQuantityTo) - ccur(iQuantityFrom))+1
						TempShipPrice = TempShipPrice + (ccur(TierCnt) * ccur(AddPrice))
					else
						if IShipCnt=>ccur(iQuantityFrom) then
							TierCnt = (ccur(IShipCnt) - ccur(iQuantityFrom))+1
							TempShipPrice = TempShipPrice + (TierCnt * ccur(AddPrice))
							iTempCompleted=1
						else
							iTempCompleted=1
						end if
					end if
					rsIncretObj.moveNext
				loop
			end if
			
			availableShipStr=availableShipStr&"|"&tempShipPrice
		else
			query1="SELECT idshipservice FROM shipService WHERE serviceCode like 'C" & idTemp & "' AND serviceFree<>0 AND serviceFreeOverAmt<" & pShipSubTotal & ";"
			set rsShipObj=conntemp.execute(query1)
			if NOT rsShipObj.eof then
				ifound=1
				tempShipPrice=0
				availableShipStr=availableShipStr&"|?|CUSTOM|C"&idTemp&"|"&rstemp("FlatShipTypeDesc")&"|0"
				iCustomFlag=1
			end if
		end if
		
		availableShipStr=availableShipStr&"|"&CustomShipDelivery
		
		if pcv_UseNegotiatedRates<>"1" then
			availableShipStr=availableShipStr&"|NONE"
		end if

	End If '// If len(VarWQP)>0 Then
	rstemp.moveNext
loop

if iCustomFlag=1 then
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Start: Pre-Filter Custom Options for Availability
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	pcv_strOptionFilterPass=pcf_PreFilter("CUSTOM", availableShipStr)
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' End: Pre-Filter Custom Options for Availability
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if pcv_strOptionFilterPass=-1 then
		strDefaultProvider="CUSTOM"
		iShipmentTypeCnt=iShipmentTypeCnt+1
		strOptionShipmentType=strOptionShipmentType&"<option value=CUSTOM>"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_j")&"</option>"
		strTabShipmentType=strTabShipmentType&"[TAB]CUSTOM,"&ship_dictLanguage.Item(Session("language")&"_chooseShpmnt_j")&"[/TAB]"
	end if
end if

if instr(strOptionShipmentType, scDefaultProvider) AND scDefaultProvider<>"" AND iShipmentTypeCnt>1 then
	strDefaultProvider=scDefaultProvider
	TempDefaultProvider = scDefaultProvider
else
	If instr(strOptionShipmentType, "UPS") Then
		strDefaultProvider="UPS"
		TempDefaultProvider = strDefaultProvider
	End If
end if

if pcv_intTotPackageNum="1" then
	Dim tmpList
	tmpList="*****"
	Dim tmpCount,tmpCount1
	tmpCount1=0
	tmpCount=0
	pcCartArray=Session("pcCartSession")
	pcCartIndex=Session("pcCartIndex")

	for f=1 to pcCartIndex
		tmp_idproduct=pcCartArray(f,0)
		query="SELECT products.pcDropShipper_ID,pcDropShippersSuppliers.pcDS_IsDropShipper FROM products,pcDropShippersSuppliers WHERE products.idproduct=" & tmp_idproduct & " AND products.pcProd_IsDropShipped=1 AND pcDropShippersSuppliers.idproduct=products.idproduct;"
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

		if not rs.eof then
			if Instr(tmpList,"*****" & rs("pcDropShipper_ID") & "**" & rs("pcDS_IsDropShipper") & "*****")=0 then
				tmpList=tmpList & rs("pcDropShipper_ID") & "**" & rs("pcDS_IsDropShipper") & "*****"
				tmpCount=tmpCount+1
			end if
		else
			tmpCount1=tmpCount1+1
		end if
		set rs=nothing
	next
	if tmpCount1>0 then
		pcv_intTotPackageNum=pcv_intTotPackageNum+tmpCount
	else
		pcv_intTotPackageNum=tmpCount
	end if
end if

'kill sessions
for q=1 to pcv_intPackageNum
	'session("UPSPackWidth"&q)=""
	session("FEDEXWSPackWidth"&q)="" '// WS
	session("CPPackWidth"&q)=""
	session("UPSPackHeight"&q)=""
	session("FEDEXWSPackHeight"&q)="" '// WS
	session("CPPackHeight"&q)=""
	session("UPSPackLength"&q)=""
	session("FEDEXWSPackLength"&q)="" '// WS
	session("CPPackLength"&q)=""
	'session("UPSPackWeight"&q)=""
	session("UPSPackPrice"&q)=""
	session("FEDEXWSPackWeight"&q)="" '// WS
	session("FEDEXWSPackPounds"&q)="" '// WS
	session("FEDEXWSPackOunces"&q)="" '// WS
	session("CPPackWeight"&q)=""
	session("USPSPackWidth"&q)=""
	session("USPSPackHeight"&q)=""
	session("USPSPackLength"&q)=""
	session("USPSPackPounds"&q)=""
	session("USPSPackOunces"&q)=""
next

Public Function pcf_PreFilter(ShippingProvidor, availableShipStr)
		on error resume next

	pcv_strCustomOptionFilterPass=0
	Session("FilterArray"&ShippingProvidor)=split(availableShipStr,"|?|")
	for i=lbound(Session("FilterArray"&ShippingProvidor)) to (Ubound(Session("FilterArray"&ShippingProvidor)))
		PreFilterDetailsArray=split(Session("FilterArray"&ShippingProvidor)(i),"|")

		if ubound(PreFilterDetailsArray)>0 then
			'// LOOP WITH EACH CUSTOM OPTION
			if (ucase(PreFilterDetailsArray(0))=ShippingProvidor) then
				if PreFilterDetailsArray(1)<>"" then

					'// Pre-Filter Customer Limitations
					query="SELECT serviceLimitation FROM shipService WHERE shipService.serviceActive=-1 AND serviceCode='"& PreFilterDetailsArray(1) &"';"
					set rsPreFilter=Server.CreateObject("ADODB.RecordSet")
					set rsPreFilter=connTemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rsPreFilter=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					do until rsPreFilter.eof
						serviceLimitation=rsPreFilter("serviceLimitation")
						customerLimitation=0
						if serviceLimitation<>0 then
							if serviceLimitation=1 then
								if Universal_destination_country=scShipFromPostalCountry then
									customerLimitation=1
								end if
							end if
							if serviceLimitation=2 then
								if Universal_destination_country<>scShipFromPostalCountry then
									customerLimitation=1
								end if
							end if
							if serviceLimitation=3 then
								if ucase(trim(Universal_destination_country))<>"US" then
									customerLimitation=1
								else
									if ucase(trim(Universal_destination_provOrState))="AK" OR ucase(trim(Universal_destination_provOrState))="HI" OR ucase(trim(Universal_destination_provOrState))="AS" OR ucase(trim(Universal_destination_provOrState))="BVI" OR ucase(trim(Universal_destination_provOrState))="GU" OR ucase(trim(Universal_destination_provOrState))="MPI" OR ucase(trim(Universal_destination_provOrState))="MP" OR ucase(trim(Universal_destination_provOrState))="PR" OR ucase(trim(Universal_destination_provOrState))="VI" then
										customerLimitation=1
									end if
								end if
							end if
							if serviceLimitation=4 then
								if ucase(trim(Universal_destination_country))<>"US" then
									customerLimitation=1
								else
									if ucase(trim(Universal_destination_provOrState))<>"AK" AND ucase(trim(Universal_destination_provOrState))<>"HI" then
										customerLimitation=1
									end if
								end if
							end if
						end if
						if customerLimitation=0 then
							pcv_strCustomOptionFilterPass=-1
						end if
					rsPreFilter.movenext
					loop
					set rsPreFilter=nothing

				end if '// if PreFilterDetailsArray(1)<>"" then
			end if '// if (ucase(PreFilterDetailsArray(0))="CUSTOM") then
		end if '// if ubound(PreFilterDetailsArray)>0 then
	next
	Session("FilterArray"&ShippingProvidor)=""
	pcf_PreFilter=pcv_strCustomOptionFilterPass
End Function
%>