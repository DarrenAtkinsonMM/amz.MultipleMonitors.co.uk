<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="opc_contentType.asp" -->
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

Call SetContentType()
Call pcs_CheckLoggedIn()

dim taxLoc, taxPrdAmount


taxLoc=0
taxPrdAmount=0

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Retreive the saved shipping information from the customer sessions table
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
query="SELECT  pcCustomerSessions.idDbSession, pcCustomerSessions.randomKey, pcCustomerSessions.idCustomer, pcCustomerSessions.pcCustSession_CustomerEmail, pcCustomerSessions.pcCustSession_ShippingResidential, pcCustomerSessions.pcCustSession_ShippingAddress, pcCustomerSessions.pcCustSession_ShippingAddress2, pcCustomerSessions.pcCustSession_ShippingCity, pcCustomerSessions.pcCustSession_ShippingStateCode, pcCustomerSessions.pcCustSession_ShippingProvince, pcCustomerSessions.pcCustSession_ShippingPostalCode, pcCustomerSessions.pcCustSession_ShippingCountryCode FROM pcCustomerSessions WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&") AND ((pcCustomerSessions.idCustomer)="&session("idCustomer")&")) ORDER BY pcCustomerSessions.idDbSession DESC;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if NOT rs.eof then
	pcCustomerEmail=rs("pcCustSession_CustomerEmail")
	pResidentialShipping=rs("pcCustSession_ShippingResidential")
	pcShippingAddress=rs("pcCustSession_ShippingAddress")
	pcShippingAddress2=rs("pcCustSession_ShippingAddress2")
	pcShippingCity=rs("pcCustSession_ShippingCity")
	pcShippingStateCode=rs("pcCustSession_ShippingStateCode")
	pcShippingProvince=rs("pcCustSession_ShippingProvince")
	pcShippingPostalCode=rs("pcCustSession_ShippingPostalCode")
	pcShippingCountryCode=rs("pcCustSession_ShippingCountryCode")
end if

set rs=nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Retreive the saved shipping information from the customer sessions table
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Retreive the saved shipping information from the customer sessions table
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
query="SELECT  customers.idCustomer, customers.address, customers.address2, customers.stateCode, customers.state, customers.city, customers.zip, customers.countryCode FROM customers WHERE customers.idCustomer="& session("idCustomer") &";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if NOT rs.eof then
	pcBillingAddress=rs("address")
	pcBillingAddress2=rs("address2")
	pcBillingStateCode=rs("stateCode")
	pcBillingCity=rs("city")
	pcBillingProvince=rs("state")
	pcBillingPostalCode=rs("zip")
	pcBillingCountryCode=rs("countryCode")
end if

set rs=nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Retreive the saved shipping information from the customer sessions table
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If pcShippingAddress<>"" AND ptaxshippingaddress="1" then
	pcBillingStateCode=pcShippingStateCode
	pcBillingCountryCode=pcShippingCountryCode
	pcBillingPostalCode=pcShippingPostalCode
end if

'Trim PostalCode
if len(pcBillingPostalCode)>5 then
	pcBillingPostalCode=left(pcBillingPostalCode,5)
end if

'Check the PostalCode Length for United States
If pcBillingCountryCode="US" Then
	if len(pcBillingPostalCode)<5 then
		response.clear
		Call SetContentType()
		response.Write("ZIPLENGTH")
		response.End()
	end if
End If

' if customer use anotherState, insert a dummy state code to simplify SQL sentence
if pcBillingStateCode="" then
   pcBillingStateCode="**"
end if

If ptaxfile=1 then
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Customer is using TAX FILE
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if request.Form("TaxSubmit")<>"" then
		TaxChoice=getUserInput(request.Form("TaxChoice"),0)
		taxLoc=getUserInput(request.Form("TOTAL_SALES_TAX"&TaxChoice),0)
		TAX_SHIPPING_ALONE=getUserInput(request.Form("TAX_SHIPPING_ALONE"&TaxChoice),0)
		TAX_SHIPPING_AND_HANDLING_TOGETHER=getUserInput(request.Form("TAX_SHIPPING_AND_HANDLING_TOGETHER"&TaxChoice),0)
	else
		'dynamically retrieve the current directory
		Dim sScriptDir
		sScriptDir = Request.ServerVariables("SCRIPT_NAME")
		sScriptDir = StrReverse(sScriptDir)
		sScriptDir = Mid(sScriptDir, InStr(1, sScriptDir, "/"))
		sScriptDir = StrReverse(sScriptDir)

		'get the file name
		dim Filename
		Filename=ptaxfilename
		Const ForReading = 1, ForWriting = 2, ForAppending = 8 

		dim FSO
		set FSO = Server.CreateObject("scripting.FileSystemObject") 
		
		'map the logincal path to the physical system path
		Dim Filepath
		Filepath=Server.MapPath(sScriptDir) & "\tax\" & Filename
		FileExist=1
		if NOT FSO.FileExists(Filepath) then
			FileExist=0
		end if
		
		If pcBillingCountryCode="US" OR pcBillingCountryCode="CA" then
		
			'see if state is a taxable one and then flag
			taxStateArray=split(ptaxRateState,", ")
			taxRateArray=split(ptaxRateDefault,", ")
			taxSNHArray=split(ptaxSNH,", ")
			intTaxableState=0
			for i=0 to ubound(taxStateArray)-1
				if taxStateArray(i)=pcBillingStateCode then
					'flag
					intTaxableState=1
					if ubound(taxRateArray)>-1 then
						intTaxRateDefault=taxRateArray(i)
					else
						intTaxRateDefault=0
					end if
					
					if ubound(taxSNHArray)>-1 then
						strTaxSNH=taxSNHArray(i)
					else
						strTaxSNH="NN"
					end if
					
					select case strTaxSNH
						case "YY"
							TAX_SHIPPING_ALONE="Y"
							TAX_SHIPPING_AND_HANDLING_TOGETHER="Y"
						case "YN"
							TAX_SHIPPING_ALONE="Y"
							TAX_SHIPPING_AND_HANDLING_TOGETHER="N"
						case "NN"
							TAX_SHIPPING_ALONE="N"
							TAX_SHIPPING_AND_HANDLING_TOGETHER="N"
					end select
				end if
			next
			
			dim f
			IF FileExist=1 THEN
			set f=FSO.GetFile(Filepath)
			
			Dim TextStream
			set TextStream=f.OpenAsTextStream(ForReading, -2) 
			
			zipCnt=0
			optionStr=""
			tmpSalesTax=""
			showDropDown=0
			do While NOT TextStream.AtEndOfStream
			  'line is found, now write it to new string
				Line=TextStream.readline

				'ignore first line
				if instr(ucase(Line), "ZIP") then
					iArray=split(Line,",")
					pcv_PostalCodeColumnFlag=False

					'loop to find correct array for each
					for q=0 to ubound(iArray)
						if iArray(q)="ZIP_CODE" then '// identify the zip code column
							ZIP_CODE_NUM=q
							pcv_PostalCodeColumnFlag=True
							'response.write q&"<BR>"
						end if
						if iArray(q)="COUNTY_NAME" then
							COUNTY_NAME_NUM=q
							'response.write q&"<BR>"
						end if
						if iArray(q)="CITY_NAME" then
							CITY_NAME_NUM=q
							'response.write q&"<BR>"
						end if
						if iArray(q)="TOTAL_SALES_TAX" then
							TOTAL_SALES_TAX_NUM=q
							'response.write q&"<BR>"
						end if
						if iArray(q)="TAX_SHIPPING_ALONE" then
							TAX_SHIPPING_ALONE_NUM=q
							'response.write q&"<BR>"
						end if
						if iArray(q)="TAX_SHIPPING_AND_HANDLING_TOGETHER" then
							TAX_SHIPPING_AND_HANDLING_TOGETHER_NUM=q
							'response.write q&"<BR>"
						end if
					next
					'response.end
				else
					'SEE IF MORE THEN ONE ZIP CODE EXIST
					if instr(Line, pcBillingPostalCode) then						
						zArray=split(Line,",")	
				
						pcv_ValidPostalCode = False						
						if pcv_PostalCodeColumnFlag=True then '// If we can identify the zip code column, then check it's valid
							ZIP_CODE_NAME = zArray(ZIP_CODE_NUM)
							if ZIP_CODE_NAME = pcBillingPostalCode then
								pcv_ValidPostalCode=True
							end if
						end if

						if pcv_PostalCodeColumnFlag=False OR pcv_ValidPostalCode=True then
                        
                            COUNTY_NAME=zArray(COUNTY_NAME_NUM)
                            CITY_NAME=zArray(CITY_NAME_NUM)
                            TOTAL_SALES_TAX=zArray(TOTAL_SALES_TAX_NUM)
                            TAX_SHIPPING_ALONE=zArray(TAX_SHIPPING_ALONE_NUM)
                            TAX_SHIPPING_AND_HANDLING_TOGETHER=zArray(TAX_SHIPPING_AND_HANDLING_TOGETHER_NUM)

							if tmpSalesTax="" then
								tmpSalesTax=TOTAL_SALES_TAX
							else
								if showDropDown=0 AND tmpSalesTax<>TOTAL_SALES_TAX then
									showDropDown=1
								end if
							end if					
							zipCnt=zipCnt+1

							optionStr=optionStr&"<option value="""&zipCnt&""">"&CITY_NAME&" - "&COUNTY_NAME&"</option>"
							optionTotalTax=optionTotalTax&"<input type='hidden' name='TOTAL_SALES_TAX"&zipCnt&"' Value="""&TOTAL_SALES_TAX&""">" 
							optionTaxShipAloneStr=optionTaxShipAloneStr&"<input type='hidden' name='TAX_SHIPPING_ALONE"&zipCnt&"' Value="""&TAX_SHIPPING_ALONE&""">" 
							optionTaxShipHandStr=optionTaxShipHandStr&"<input type='hidden' name='TAX_SHIPPING_AND_HANDLING_TOGETHER"&zipCnt&"' Value="""&TAX_SHIPPING_AND_HANDLING_TOGETHER&""">" 
						end if
					end if
				end if
			loop
			
			TextStream.Close
			END IF
			if showDropDown=0 AND zipCnt>1 then
				zipCnt=1
			end if
			If zipCnt=0 AND intTaxableState=1 then
				taxLoc=taxLoc+(intTaxRateDefault/100) 
			End If

			if zipCnt>1 then %>
				<form name="TaxForm" id="TaxForm">
					<div class="pcShowContent">
						<div class="pcFormItem">
							<div class="pcFormItemFull"><% Response.Write dictLanguage.Item(Session("language")&"_calculateTax_1")%></div>
						</div>
						<div class="pcSpacer"></div>
						<div class="pcFormItem">
							<select name="TaxChoice">
								<%=optionStr%>
							</select>
							<%=optionTotalTax%>
							<%=optionTaxShipAloneStr%>
							<%=optionTaxShipHandStr%>
							<input type="hidden" name="IntZipCnt" value="<%=zipCnt%>">
						</div>

						<div id="TaxLoader"></div>
					
						<div class="pcSpacer"></div>

						<div class="pcFormButtons">
							<button class="pcButton pcButtonContinue" name="TaxSubmit" id="TaxSubmit">
								<img src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>" />
								<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
							</button>
						</div>
					</div>
				</form>
				<script type=text/javascript>
				//*Submit Tax Form
						$pc('#TaxSubmit').click(function(){
						{
							$pc("#TaxLoader").html('<img src="<%=pcf_getImagePath("images","ajax-loader1.gif")%>" style="width: 20px; height: 20px; text-align: center"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_tax_1"))%>');
							$pc("#TaxLoader").show();	
							$pc.ajax({
								type: "POST",
								url: "opc_tax.asp",
								data: $pc('#TaxForm').formSerialize() + "&TaxSubmit=yes",
								timeout: 450000,
								success: function(data, textStatus){
								if (data=="SECURITY")
								{
									// Session Expired
									window.location="msg.asp?message=1";
								}
								else
								{
									if (data=="OK")
									{

										<%
										'// TAX APPLIED
										'   Generate new Order Preview
										'	Open Payment Panel
										%>

										$pc("#TaxLoader").hide();
										$pc('#TaxContentArea').hide();	
										recalculate("","#TaxLoadContentMsg",1,'Y');
										$pc("#PaymentContentArea").show();
										
										
									}
									else
									{
										$pc("#TaxLoader").html('<img src="<%=pcf_getImagePath("images","pc_icon_error_small.png")%>" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_tax_3"))%>');
										$pc("#TaxLoader").show();
									}
									}
								}
				 			});
							return(false);
						}
						return(false);
						});
					</script>
			<% else
				taxLoc=taxLoc+(TOTAL_SALES_TAX)
				zipCnt=1
			end if
		Else
			taxLoc=taxLoc+(TOTAL_SALES_TAX)
			zipCnt=1
		End if
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// END TAX FILE
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
else
	if ptaxVAT="1" then
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Customer is using VAT SETTINGS
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		taxLoc=0
		TAX_SHIPPING_ALONE="NA"
		TAX_SHIPPING_AND_HANDLING_TOGETHER="NA"
		zipCnt=1
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END VAT SETTINGS
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	else
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Customer is using MANUAL TAX PER PLACE
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		If pcBillingPostalCode&""="" Then
			query="SELECT taxLoc, taxDesc FROM taxLoc WHERE ((stateCode='" &pcBillingStateCode& "' AND stateCodeEq=-1) OR (stateCode IS NULL) OR (stateCode<>'" &pcBillingStateCode& "' AND stateCodeEq=0)) AND ((CountryCode='"&pcBillingCountryCode&"' AND CountryCodeEq=-1) OR (CountryCode IS NULL) OR (CountryCode<>'" &pcBillingCountryCode& "' AND CountryCodeEq=0));"
		Else
			query="SELECT taxLoc, taxDesc FROM taxLoc WHERE ((stateCode='" &pcBillingStateCode& "' AND stateCodeEq=-1) OR (stateCode IS NULL) OR (stateCode<>'" &pcBillingStateCode& "' AND stateCodeEq=0)) AND ((CountryCode='"&pcBillingCountryCode&"' AND CountryCodeEq=-1) OR (CountryCode IS NULL) OR (CountryCode<>'" &pcBillingCountryCode& "' AND CountryCodeEq=0)) AND ((zip='" &pcBillingPostalCode& "' AND zipEq=-1) OR (zip IS NULL) OR (zip<>'" &pcBillingPostalCode& "' AND zipEq=0));"
		End If
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		if  rs.eof then 
		 ' there are no taxes defined for that zone
		else
		
			taxCnt=0
			
			do until rs.eof
				pcv_tmpTaxLoc=rs("taxLoc")
				pcv_tmpTaxDesc= rs("taxDesc")
				if ptaxseparate="1" then
					taxCnt=taxCnt+1
					session("taxDesc"&taxCnt)=pcv_tmpTaxDesc
					session("tax"&taxCnt)=pcv_tmpTaxLoc
				end if     
				taxLoc=taxLoc+pcv_tmpTaxLoc

				rs.movenext
			loop
		end if
		
		set rs=nothing
		
		if ptaxseparate="1" then
			session("taxCnt")=taxCnt
		end if 
		TAX_SHIPPING_ALONE="NA"
		TAX_SHIPPING_AND_HANDLING_TOGETHER="NA"
		zipCnt=1
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// END MANUAL TAX PER PLACE
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	end if
End if

if zipCnt=1 OR request.Form("TaxSubmit")<>"" then
	'check for taxes per product now.
	if session("customerType")=1 AND ptaxwholesale=0 then
		taxPrdAmount=Cdbl(0)
	else
		Dim pcCartArray
		'*****************************************************************************************************
		'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
		'*****************************************************************************************************
		%><!--#include file="pcVerifySession.asp"--><%
		pcs_VerifySession
		'*****************************************************************************************************
		'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
		'*****************************************************************************************************
	
		for f=1 to pcCartIndex
			if pcCartArray(f,10)=0 then

				query="SELECT taxPerProduct FROM taxPrd WHERE ((stateCode='" &pcBillingStateCode& "' AND stateCodeEq=-1) OR (stateCode IS NULL) OR (stateCode<>'" &pcBillingStateCode& "' AND stateCodeEq=0)) AND ((CountryCode='"&pcBillingCountryCode&"' AND CountryCodeEq=-1) OR (CountryCode IS NULL) OR (CountryCode<>'" &pcBillingCountryCode& "' AND CountryCodeEq=0)) AND ((zip='" &pcBillingPostalCode& "' AND zipEq=-1) OR (zip IS NULL) OR (zip<>'" &pcBillingPostalCode& "' AND zipEq=0)) AND idProduct=" &pcCartArray(f,0)
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				'Check if this might be an apparel sub-product (Not an Apparel Add-on file - this code works fine in the standard build)
				If rs.eof AND statusAPP="1" then
					query = "SELECT pcprod_ParentPrd FROM products WHERE idProduct = " &pcCartArray(f,0)
					set rs1=server.CreateObject("ADODB.RecordSet")
					set rs1=conntemp.execute(query)
					tmp_TPIdProduct=0
					if not rs1.eof then
						tmp_TPIdProduct = rs1("pcprod_ParentPrd")
					end if
					set rs1=nothing
					If tmp_TPIdProduct<>0 Then
						query="SELECT taxPerProduct FROM taxPrd WHERE ((stateCode='" &pcBillingStateCode& "' AND stateCodeEq=-1) OR (stateCode IS NULL) OR (stateCode<>'" &pcBillingStateCode& "' AND stateCodeEq=0)) AND ((CountryCode='"&pcBillingCountryCode&"' AND CountryCodeEq=-1) OR (CountryCode IS NULL) OR (CountryCode<>'" &pcBillingCountryCode& "' AND CountryCodeEq=0)) AND ((zip='" &pcBillingPostalCode& "' AND zipEq=-1) OR (zip IS NULL) OR (zip<>'" &pcBillingPostalCode& "' AND zipEq=0)) AND idProduct=" &tmp_TPIdProduct
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=conntemp.execute(query)
					End If
				End If
				
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
		
				taxPrdArray=0
				do until rs.eof 
					taxPrdAmount=taxPrdAmount+(rs("taxPerProduct") * ( pcCartArray(f,2) * (pcCartArray(f,5)+pcCartArray(f,3)) )) 
					taxPrdArray=1  
					rs.movenext
				loop
				
				set rs=nothing
				
				pcCartArray(f,24)=taxPrdArray
			end if ' pcCartArray =0
		next
	end if	
	
	if pcBillingStateCode="**" then
		 pcBillingStateCode=""
	end if
	
	'Update customer session data
	query="UPDATE pcCustomerSessions SET pcCustSession_TaxShippingAlone='"&TAX_SHIPPING_ALONE&"', pcCustSession_TaxShippingAndHandlingTogether='"&TAX_SHIPPING_AND_HANDLING_TOGETHER&"',pcCustSession_TaxLocation='"&taxLoc&"',pcCustSession_TaxProductAmount='"&taxPrdAmount&"',pcCustSession_TaxCountyCode='' WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&") AND ((pcCustomerSessions.idCustomer)="&session("idCustomer")&"));"

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	set rs=nothing
	call closedb()
	response.clear
	Call SetContentType()
	response.write "OK"

else
	'Update customer session data
	query="UPDATE pcCustomerSessions SET pcCustSession_TaxShippingAlone='"&TAX_SHIPPING_ALONE&"',		pcCustSession_TaxShippingAndHandlingTogether='"&TAX_SHIPPING_AND_HANDLING_TOGETHER&"',pcCustSession_TaxLocation='"&taxLoc&"',pcCustSession_TaxProductAmount='"&taxPrdAmount&"',pcCustSession_TaxCountyCode='' WHERE (((pcCustomerSessions.idDbSession)="&session("pcSFIdDbSession")&") AND ((pcCustomerSessions.randomKey)="&session("pcSFRandomKey")&") AND ((pcCustomerSessions.idCustomer)="&session("idCustomer")&"));"

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	set rs=nothing
	call closedb()
end if
%>
<% 
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing %>