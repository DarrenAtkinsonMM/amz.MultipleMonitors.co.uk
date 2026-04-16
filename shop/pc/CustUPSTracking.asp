<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Response.buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/UPSconstants.asp" -->
<!--#include file="CustLIv.asp"-->
<!--#include file="header_wrapper.asp"-->

<%
	If Request.QueryString("msg") = "1" Then
		msg = "You must agree to the UPS Tracking Terms and Conditions before continuing."
	End If	
%>
<style>
.resultstable{
    color:<%=FColor%>;
    font-family:<%=FFType%>;
    font-size:12px;
	}
TD{	font-family:<%=FFType%>;
	font-size: 12px;
	color:<%=FColor%>;
	}
</style>

<% 
Function ShowUPSTerms()
%>
	<div class="pcSpacer"></div>
	<div class="pcFormItem">
		<div class="pcFormItemFull"><strong>UPS Tracking Terms &amp; Conditions:</strong></div>
		<div class="pcFormItemFull">
			<div class="pcSmallText">
				NOTICE: The UPS package tracking systems accessed via this
				service (the "Tracking Systems") and tracking information obtained
				through this service (the "Information") are the private property of
				UPS. UPS authorizes you to use the Tracking Systems solely to track
				shipments tendered by or for you to UPS for delivery and for no
				other purpose. Without limitation, you are not authorized to make
				the Information available on any web site or otherwise reproduce,
				distribute, copy, store, use or sell the Information for commercial
				gain without the express written consent of UPS. This is a personal
				service, thus your right to use the Tracking Systems or Information
				is non-assignable. Any access or use that is inconsistent with these
				terms is unauthorized and strictly prohibited.
			</div>
		</div>
	</div>
<%
End Function
%>

<div id="pcMain">
	<div class="pcMainContent">
		<h3>
			<img src="<%=pcf_getImagePath("../UPSLicense","LOGO_S2.jpg")%>" width="45" height="50">
			UPS Tracking
		</h3>
		<%

		'//UPS Variables
		mySQL="SELECT active, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=3"
		set rstemp=conntemp.execute(mySQL)
		ups_license_key=trim(rstemp("AccessLicense"))
		ups_userid=trim(rstemp("userID"))
		ups_password=trim(rstemp("password"))
		ups_active=rstemp("active")

		if request.form("SubmitTracking")<>"" then
			dim pitracknumber
			pitracknumber=request("itracknumber")
			pitracknumber=replace(pitracknumber," ","")
			session("itracknumber")=uCase(pitracknumber)
			if request.form("iagree")="" then
				call closedb()
				response.redirect "CustUPSTracking.asp?msg=1"
			end if
			'//UPS Rates
			ups_trackdata=""
			ups_trackdata="<?xml version=""1.0""?>"
			ups_trackdata=ups_trackdata&"<AccessRequest xml:lang=""en-US"">"
			ups_trackdata=ups_trackdata&"<AccessLicenseNumber>"&ups_license_key&"</AccessLicenseNumber>"
			ups_trackdata=ups_trackdata&"<UserId>"&ups_userid&"</UserId>"
			ups_trackdata=ups_trackdata&"<Password>"&ups_password&"</Password>"
			ups_trackdata=ups_trackdata&"</AccessRequest>"
			ups_trackdata=ups_trackdata&"<?xml version=""1.0""?>"
			ups_trackdata=ups_trackdata&"<TrackRequest xml:lang=""en-US"">"
			ups_trackdata=ups_trackdata&"<Request>"
			ups_trackdata=ups_trackdata&"<TransactionReference>"
			ups_trackdata=ups_trackdata&"<CustomerContext>Example 1</CustomerContext>"
			ups_trackdata=ups_trackdata&"<XpciVersion>1.0001</XpciVersion>"
			ups_trackdata=ups_trackdata&"</TransactionReference>"
			ups_trackdata=ups_trackdata&"<RequestAction>Track</RequestAction>"
			ups_trackdata=ups_trackdata&"<RequestOption>activity</RequestOption>"
			ups_trackdata=ups_trackdata&"</Request>"
			ups_trackdata=ups_trackdata&"<TrackingNumber>"&session("itracknumber")&"</TrackingNumber>"
			ups_trackdata=ups_trackdata&"</TrackRequest>"
			'get URL to post to
			ups_URL="https://onlinetools.ups.com/ups.app/xml/Track"
	
			toResolve = 3000
			toConnect = 3000
			toSend = 3000
			toReceive = 3000
	
			Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp")
			'srvXmlHttp.setTimeouts toResolve, toConnect, toSend, toReceive ' not needed but a handy feature
			srvXmlHttp.open "POST", ups_URL, false
			srvXmlHttp.send(ups_trackdata)
			result = srvXmlHttp.responseText
			'response.write result&"<BR>"
			'response.end
			Set XMLdoc = server.CreateObject("Msxml2.DOMDocument")
			XMLDoc.async = false 
			if xmldoc.loadXML(result) then ' if loading from a string
				set objLst = xmldoc.getElementsByTagName("ResponseStatusCode") 
				for i = 0 to (objLst.length - 1)
					varStatus=objLst.item(i).text
					if varStatus="0" then
						set objLst = xmldoc.getElementsByTagName("Error") 
						for j = 0 to (objLst.length - 1)
							for k=0 to ((objLst.item(j).childNodes.length)-1)
								If objLst.item(j).childNodes(k).nodeName="ErrorDescription" then
									varErrorDescription=objLst.item(j).childNodes(k).text
								end if
							next
						next
					end if
				next
			
				if varStatus="0" then %>
					<div class="pcErrorMessage">
						Error retrieving tracking information: <%= varErrorDescription%>
					</div>
						
					<% ShowUPSTerms %>
				<% else %>
					<h4>Tracking Summary</h4>
					<% set objLst = xmldoc.getElementsByTagName("ShipTo")
					for i = 0 to (objLst.length - 1)
					for j=0 to ((objLst.item(i).childNodes.length)-1)
						If objLst.item(i).childNodes(j).nodeName="Address" then
							for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
								if objLst.item(i).childNodes(j).childNodes(k).nodeName="AddressLine1" then
									upstAddressLine1= objLst.item(i).childNodes(j).childNodes(k).text
								end if
								if objLst.item(i).childNodes(j).childNodes(k).nodeName="City" then
									upstCity= objLst.item(i).childNodes(j).childNodes(k).text
								end if
								if objLst.item(i).childNodes(j).childNodes(k).nodeName="StateProvinceCode" then
									upstStateProvinceCode= objLst.item(i).childNodes(j).childNodes(k).text
								end if
								if objLst.item(i).childNodes(j).childNodes(k).nodeName="PostalCode" then
									upstPostalCode= objLst.item(i).childNodes(j).childNodes(k).text
								end if
								if objLst.item(i).childNodes(j).childNodes(k).nodeName="CountryCode" then
									upstCountryCode= objLst.item(i).childNodes(j).childNodes(k).text
								end if
							next
						End if
					next
					next
					%>
					<% set objLst = xmldoc.getElementsByTagName("Service")
					for i = 0 to (objLst.length - 1)
					for j=0 to ((objLst.item(i).childNodes.length)-1)
						If objLst.item(i).childNodes(j).nodeName="Code" then
							upstServiceCode= objLst.item(i).childNodes(j).text
						End if
						If objLst.item(i).childNodes(j).nodeName="Description" then
							upstServiceDescription= objLst.item(i).childNodes(j).text
						End if
					next
					next
					select case upstServiceCode
					case "01"
						upstService="UPS Next Day Air <sup>&reg;</sup>"
					case "02"
						upstService="UPS 2nd Day Air <sup>&reg;</sup>"
					case "03"
						upstService="UPS Ground"
					case "07"
						upstService="UPS Worldwide Express <sup>SM</sup>"
					case "08"
						upstService="UPS Worldwide Expedited <sup>SM</sup>"
					case "11"
						upstService="UPS Standard To Canada"
					case "12"
						upstService="UPS 3 Day Select <sup>&reg;</sup>"
					case "13"
						upstService="UPS Next Day Air Saver <sup>&reg;</sup>"
					case "14"
						upstService="UPS Next Day Air<sup>&reg;</sup> Early A.M. <sup>&reg;</sup>"
					case "54"
						upstService="UPS Worldwide Express Plus <sup>SM</sup>"
					case "59"
						upstService="UPS 2nd Day Air A.M. <sup>&reg;</sup>"
					case "65"
						upstService="UPS Express Saver <sup>SM</sup>"
					case else
						upstService=upstServiceDescription
					end select
					%>
					<% set objLst = xmldoc.getElementsByTagName("Package")
					for i = 0 to (objLst.length - 1)
					for j=0 to ((objLst.item(i).childNodes.length)-1)
						If objLst.item(i).childNodes(j).nodeName="TrackingNumber" then
							upstNumber= objLst.item(i).childNodes(j).text
						End if
					next
					next
					%>
					<div class="pcFormItem">
						<div class="pcFormItemFull"><strong>Tracking Number: </strong>&nbsp;<%=upstNumber%></div>
						<div class="pcFormItemFull"><strong>Shipped To: </strong>&nbsp;<%=upstAddressLine1%>&nbsp;<%=upstCity%>,&nbsp;<%=upstStateProvinceCode%>&nbsp;<%=upstPostalCode%>&nbsp;<%=upstCountryCode%></div>
						<div class="pcFormItemFull"><strong>Service: </strong>&nbsp;<%=upstService%></div>
					</div>

					<div class="pcSpacer"></div>
					<%
						col_DateClass			= "pcCol-3"
						col_TimeClass			= "pcCol-3"
						col_LocationClass = "pcCol-3"
						col_ActivityClass = "pcCol-3"
					%>
					<div id="pcTableUPSTracking" class="pcTable">
						<div class="pcTableHeader">
							<div class="<%= col_DateClass %>">Date</div>
							<div class="<%= col_TimeClass %>">Time</div>
							<div class="<%= col_LocationClass %>">Location</div>
							<div class="<%= col_ActivityClass %>">Activity</div>
						</div>

					<%
						set objLst = xmldoc.getElementsByTagName("Activity") 
						for i = 0 to (objLst.length - 1)
							for j=0 to ((objLst.item(i).childNodes.length)-1)
								If objLst.item(i).childNodes(j).nodeName="ActivityLocation" then
									for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
										if objLst.item(i).childNodes(j).childNodes(k).nodeName="Address" then
											UPSLOCATION= objLst.item(i).childNodes(j).childNodes(k).text
										end if
									next
								End if
								If objLst.item(i).childNodes(j).nodeName="Status" then
									for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
										if objLst.item(i).childNodes(j).childNodes(k).nodeName="StatusType" then
											for l=0 to ((objLst.item(i).childNodes(j).childNodes(k).childNodes.length)-1)
												if objLst.item(i).childNodes(j).childNodes(k).childNodes(l).nodeName="Description" then
													UPSACTIVITY= objLst.item(i).childNodes(j).childNodes(k).childNodes(l).text
												end if
											next
										end if
									next
								End if
								If objLst.item(i).childNodes(j).nodeName="Date" then
									UPSDATE= objLst.item(i).childNodes(j).text
									UPSYEAR=left(UPSDATE,4)
									UPSMONTH=Mid(UPSDATE, 5, 2)
									UPSDAY=right(UPSDATE, 2) 
								End If
								If objLst.item(i).childNodes(j).nodeName="Time" then
									UPSTIME= objLst.item(i).childNodes(j).text
									UPSHOUR=left(UPSTIME,2)
									UPSMINUTE=Mid(UPSTIME, 3, 2)
									if CINT(UPSHOUR)>11 then
									 UPSAMPM="PM"
									else
									 UPSAMPM="AM"
									End If
									SELECT CASE UPSHOUR
										CASE "00"
											UPSHOUR="12"
										CASE "01"
											UPSHOUR="1"
										CASE "02"
											UPSHOUR="2"
										CASE "03"
											UPSHOUR="3"
										CASE "04"
											UPSHOUR="4"
										CASE "05"
											UPSHOUR="5"
										CASE "06"
											UPSHOUR="6"
										CASE "07"
											UPSHOUR="7"
										CASE "08"
											UPSHOUR="8"
										CASE "09"
											UPSHOUR="9"
										CASE "10"
											UPSHOUR="10"
										CASE "11"
											UPSHOUR="11"
										CASE "12"
											UPSHOUR="12"
										CASE "13"
											UPSHOUR="1"
										CASE "14"
											UPSHOUR="2"
										CASE "15"
											UPSHOUR="3"
										CASE "16"
											UPSHOUR="4"
										CASE "17"
											UPSHOUR="5"
										CASE "18"
											UPSHOUR="6"
										CASE "19"
											UPSHOUR="7"
										CASE "20"
											UPSHOUR="8"
										CASE "21"
											UPSHOUR="9"
										CASE "22"
											UPSHOUR="10"
										CASE "23"
											UPSHOUR="11"
										END SELECT
								End If
							next %>
								<div class="pcTableRow">
									<div class="<%= col_DateClass %>"><%=UPSMONTH&"/"&UPSDAY&"/"&UPSYEAR%></div>
									<div class="<%= col_TimeClass %>"><%=UPSHOUR&":"&UPSMINUTE&" "&UPSAMPM%></div>
									<div class="<%= col_LocationClass %>"><%=UPSLOCATION%></div>
									<div class="<%= col_ActivityClass %>"><%=UPSACTIVITY%></div>
								</div>
							<% 	next %>
						</div>

						<% ShowUPSTerms %>
				<% end if %>
			<% end if
		else 'form is submitted
			itracknumber=request.querystring("itracknumber") 
			session("itracknumber")=itracknumber %>
			<form name="trackform" method="post" action="CustUPSTracking.asp">
				
				<% if msg<>"" then %>
					<div class="pcErrorMessage"><%= msg %></div>
				<% end if %>

				<div class="pcFormItem">
					<div class="pcFormLabel">Tracking Number:</div>
					<div class="pcFormField"><input name="itracknumber" type="text" id="itracknumber" size="30" value="<%=session("itracknumber")%>"></div>
				</div>

				<% ShowUPSTerms %>

				<div class="pcSpacer"></div>

				<div class="pcFormItem">
					<div class="pcFormItemFull">
						<input type=checkbox name="iagree" value="1">&nbsp;I agree to the Terms &amp; Conditions Set Forth Above
					</div>
				</div>

				<div class="pcFormButtons">					
					<a class="pcButton pcButtonBack" href="javascript:history.back();">
						<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
					</a>

					<button class="pcButtonSubmit" type="submit" name="SubmitTracking" value="Submit">Submit</button>
					<button class="pcButtonReset" name="Reset" value="Reset">Reset</button>
				</div>
			</form>
		<% end if %>
				
			<div class="pcSpacer"></div>
		
			<% If Request.Form.Count > 0 Then %>
				<div class="pcFormButtons">					
					<a class="pcButton pcButtonBack" href="javascript:history.back();">
						<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
					</a>
				</div>
			<% End if %>
				
			<div class="pcSpacer"></div>

			<div class="pcFormItem">
				<div class="pcFormItemFull">UPS, the UPS Shield trademark, the UPS Ready mark, <br />the UPS Developer Kit mark and the Color Brown are trademarks of <br />United Parcel Service of America, Inc. All Rights Reserved.</div>
			</div>
		</div>
	</div>
<!--#include file="footer_wrapper.asp"-->
