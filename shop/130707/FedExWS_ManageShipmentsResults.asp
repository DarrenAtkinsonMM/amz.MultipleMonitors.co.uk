<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Center for FedEx" %>
<% Section="mngAcc" %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/FedExWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="AdminHeader.asp"-->
<% 
Const iPageSize=5

Dim iPageCurrent, varFlagIncomplete, strORD, pcv_intOrderID


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// SET PAGE NAMES
pcPageName = "FedExWS_ManageShipmentsResults.asp"
ErrPageName = "FedExWS_ManageShipmentsTrack.asp"

'// OPEN DATABASE


'// SET THE FEDEX OBJECT
set objFedExClass = New pcFedExWSClass

'// GET PAGE NUMBER
if request.querystring("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=Request.QueryString("iPageCurrent")
end if

'// SORT ORDER
strORD=request("order")
if strORD="" then
	strORD="pcPackageInfo_ShippedDate DESC, idOrder"
End If
strSort=request("sort")
if strSort="" Then
	strSort="DESC"
End If 

'// GET ORDER ID
pcv_strOrderID=Request("id")
pcv_strSessionOrderID=Session("pcAdminOrderID")
if pcv_strSessionOrderID="" OR len(pcv_strOrderID)>0 then
	pcv_intOrderID=pcv_strOrderID
	Session("pcAdminOrderID")=pcv_intOrderID
else
	pcv_intOrderID=pcv_strSessionOrderID
end if
'response.write Session("pcAdminOrderID")
'response.end

' GET THE PACKAGES
' >>> Tables: pcPackageInfo
query = 		"SELECT pcPackageInfo.* "
query = query & "FROM pcPackageInfo "
query = query & "WHERE pcPackageInfo.idOrder=" & pcv_intOrderID &" ORDER BY pcPackageInfo.pcPackageInfo_ID"	

' >>> Conditions
If Request.QueryString("TypeSearch")="idOrder" Then
	tempqryStr=Request.QueryString("advquery")
	if tempqryStr="" then
		tempqryStr=" ORDER BY "& strORD &" "& strSort
	else
		tempqryStr=(int(tempqryStr) - scpre)
		query=query & " WHERE idOrder LIKE '%" & _
		tempqryStr & "%' ORDER BY "& strORD &" "& strSort
	end if
End If	
'If Request.QueryString("TypeSearch")="orderstatus" Then
'	query=query & " WHERE orders.idCustomer=customers.idCustomer AND orderstatus LIKE '%" & _
'	Request.QueryString("advquery") & "%' ORDER BY "& strORD &" "& strSort
'End If

set rs=Server.CreateObject("ADODB.Recordset") 

rs.CursorLocation=adUseClient
rs.CacheSize=iPageSize
rs.PageSize=iPageSize
rs.Open query, conntemp

if err.number <> 0 then
	call rs.Close
	set rs=nothing
	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
end If

If Not RS.EOF Then

	rs.MoveFirst
	
	'// GET MAX PAGES
	Dim iPageCount
	iPageCount=rs.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
	If iPageCurrent < 1 Then iPageCurrent=1
	
	'// SET ABSOLUTE PAGE
	rs.AbsolutePage=iPageCurrent

End If

'// DISPLAY ERROR MSG
msg=request.querystring("msg")

if msg<>"" then 
	%>
	<div class="pcCPmessage">
		<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
	</div>
	<% 
end if 


'// DISPLAY HEADER
if rs.eof then 
	presults="0"
else 
%>
	<table class="pcCPcontent">
		<tr>
			<td colspan="3">Manage FedEx&reg; Shipments for Order #: <strong><%=(scpre+int(pcv_intOrderID))%></strong><a name="top"></a></td>
		</tr>
		<tr> 
			<td width="25%" align="left" valign="bottom"> 
			<% 
				'// Showing total number of pages found and the current page number
				Response.Write "Displaying Page <b>" & iPageCurrent & "</b> of <b>" & iPageCount & "</b><br>"
				Response.Write "Total Shipments Found: <b>" & rs.RecordCount & "</b>" 

				%></td>
			<td width="50%" align="center">
				<img src="images/Clct_Prf_2c_Pos_Plt_150.png">
			</td>
		    <td width="25%" align="right" valign="bottom"><input type="button" class="btn btn-default"  name="Button2" value="Closeout &amp; Print Manifest" onClick="document.location.href='FedExWS_ManageShipmentsClose.asp?PackageInfo_ID=<%=pcv_intOrderID%>'"></td>
		</tr>
	</table>
<% end if %>

<table class="pcCPcontent">

	<form name="checkboxform" action="FedExWS_ManageShipmentsTrack.asp?id=<%=pcv_intOrderID%>&action=batch" method="post" class="pcForms">
		<tr> 
			<th nowrap>Shipped</th>
			<th nowrap>Tracking #</th>
			<th nowrap>Contents Description</th>
			<th nowrap>Package Details</th>
			<th nowrap>Labels &amp; Documents</th>
			<th nowrap>Returns</th>
			<th nowrap>Select</th>
		</tr>
		<% 
		mcount=0
		If rs.EOF Then %>
			<tr>
			<td colspan="11">
				<div class="pcCPmessage"><img src="images/pcadmin_note.gif" width="20" height="20"> No Results Found</div>			</td>
			</tr>
		<% Else
			' Showing relevant records
			Dim strCol
			strCol="#E1E1E1" 
			Dim rcount, i, x
			
			For i=1 To rs.PageSize
				
				rcount=i
				If currentPage > 1 Then
					For x=1 To (currentPage - 1)
						rcount=10 + rcount
					Next
				End If
                          
				If Not rs.EOF Then 
					If strCol <> "#FFFFFF" Then
						strCol="#FFFFFF"
					Else 
						strCol="#E1E1E1"
					End If
					
					pcv_intPackageInfo_ID=rs("pcPackageInfo_ID")
					pcv_intOrder=rs("idOrder")
					pidPackageNumber=rs("pcPackageInfo_PackageNumber")
					pcv_strTrackingNumber=rs("pcPackageInfo_TrackingNumber")
					pidPackageWeight=rs("pcPackageInfo_PackageWeight")
					pidPackageShippedDate=rs("pcPackageInfo_ShippedDate")					
					pcv_strShipMethod = rs("pcPackageInfo_ShipMethod")
					pcv_strFDXRate=rs("pcPackageInfo_FDXRate")
					pcv_strFDXCarrierCode = rs("pcPackageInfo_FDXCarrierCode")

					select case pcv_strFDXCarrierCode
						case "FDXE"
							pcv_strCarrierCode = "FedEx Express"
						case "FDXG"
							pcv_strCarrierCode = "FedEx Ground"
						case "FXCC"
							pcv_strCarrierCode = "FedEx Cargo"
						case "FXSP", "FXFR"
							pcv_strCarrierCode = "FedEx Freight"
					end select
					
          '// Get Ship Method String
          pcv_ShipMethod = ""
          query = "SELECT serviceDescription FROM shipService WHERE serviceCode = '" & Replace(pcv_strShipMethod, "FedEx: ", "") & "';"
          Set rsservice = connTemp.execute(query)
          If Not rsservice.eof Then
            pcv_ShipMethod = rsservice("serviceDescription")
          End If
          Set rsservice = Nothing
					
					mcount=mcount+1 
					%>
												
					<tr style="background-color: <%= strCol %>;" valign="top">
						<td><%=FormatDateTime(pidPackageShippedDate)%></td>
						<td>
							<% If pcv_strCarrierCode <> "FedEx Freight" Then %>
								<a href="FedExWS_ManageShipmentsTrack.asp?id=<%=pcv_intOrder%>&PackageInfo_ID=<%=pcv_intPackageInfo_ID%>"><%=pcv_strTrackingNumber%></a>
							<% Else %>
								<%=pcv_strTrackingNumber%>
							<% End If %>
            </td>
						<td>
						<%
						' GET THE PACKAGE CONTENTS
						' >>> Tables: products, ProductsOrdered
						query = 		"SELECT ProductsOrdered.pcPackageInfo_ID , products.description, products.idProduct  "
						query = query & "FROM ProductsOrdered "
						query = query & "INNER JOIN products "
						query = query & "ON ProductsOrdered.idProduct = products.idProduct "
						query = query & "WHERE ProductsOrdered.pcPackageInfo_ID=" & pcv_intPackageInfo_ID &" "
												
						set rs2=server.CreateObject("ADODB.RecordSet")
						set rs2=conntemp.execute(query)		
						
						if err.number<>0 then
							'// handle admin error
						end if
						
						if NOT rs2.eof then
							Do until rs2.eof	
								pcv_strProductDescription = rs2("description")
								%>
								<%=pcv_strProductDescription%><br />
								<%
							rs2.movenext
							Loop
						end if						
						%>						</td>
						<td>
							
						
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
							  <tr>
								<td align="right">Weight:</td>
								<td align="left"><b><%=pidPackageWeight%> lbs.</b></td>
							  </tr>
							  <tr>
								<td align="right" valign="top">Carrier:</td>
								<td align="left" nowrap>
									<b>
										<%=pcv_strCarrierCode%>
									</b>
								</td>
							  </tr>
							  <tr>
								<td align="right" valign="top">Method:</td>
								<td align="left" nowrap>
									<b>
                                        <%=pcv_ShipMethod%>
									</b>
								</td>
							  </tr>
							  <tr>
								<td align="right" nowrap>Net Rate:</td>
								<td align="left">
								<b>
								<%
								if pcv_strFDXRate > 0 then
									response.write scCurSign&money(pcv_strFDXRate)
								else
									response.write "Alternate Payor"
								end if
								%>
								</b></td>
							  </tr>
						  </table>						</td>
						<td nowrap>
              <%
                query = "SELECT pcPackageLabel_ID, pcPackageLabel_Name, pcPackageLabel_File, pcPackageLabel_FileType, pcPackageLabel_Resolution FROM pcPackageLabel WHERE pcPackageInfo_ID = " & pcv_intPackageInfo_ID & ";"
                set rslabel = connTemp.execute(query)
                
                If Not rslabel.eof Then
	                Do Until rslabel.eof
                    labelID = rslabel("pcPackageLabel_ID")
                    labelName = rslabel("pcPackageLabel_Name")
                    labelFile = rslabel("pcPackageLabel_File")
                    labelFileType = rslabel("pcPackageLabel_FileType")
                    labelResolution = rslabel("pcPackageLabel_Resolution")

										labelName = Replace(labelName, "''", "'")
                    %>
                      <a href="FedExWS_ManageShipmentsPrinting.asp?label=<%= labelFile & "." & labelFileType %>&res=<%= labelResolution %>" target="_blank"><%= labelName %></a><br />
                    <%
                    rslabel.MoveNext
                  Loop
								Else
            			%><a href="FedExWS_ManageShipmentsPrinting.asp?label=Label<%=pcv_strTrackingNumber%>.PNG" target="_blank">View/Print Label</a><br /><%
                End If
              %>
						</td>
						<td>
							<a href="FedExWS_ManageShipmentsCancel.asp?id=<%=pcv_intOrder%>&PackageInfo_ID=<%=pcv_intPackageInfo_ID%>">Cancel Shipment</a></td>
						<td><input type=checkbox name="check<%=mcount%>" value="<%=pcv_intPackageInfo_ID%>"></td>
					</tr>
					<% rs.MoveNext
				End If
			Next%>
			<input type=hidden name="count" value="<%=mcount%>">								
		<% End If %>
		<tr> 
			<td colspan="11">
				<%if mcount>0 then%>
					<a href="javascript:checkAll();"><b>Check All</b></a><b>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a></b><br>
					<br><input type="submit" class="btn btn-default" name="submit" value="Track All Selected Packages">
					<script type=text/javascript>
					function checkAll() {
					for (var j = 1; j <= <%=mcount%>; j++) {
					box = eval("document.checkboxform.check" + j); 
					if (box.checked == false) box.checked = true;
						 }
					}
						
					function uncheckAll() {
					for (var j = 1; j <= <%=mcount%>; j++) {
					box = eval("document.checkboxform.check" + j); 
					if (box.checked == true) box.checked = false;
						 }
					}
					</script>
				<%end if%>
			</td>
		</tr>
	</form>
              
	<tr>
		<td colspan="11"> 
			<% if pResults<>"0" Then %>
				<table width="100%" border="0" cellspacing="0" cellpadding="4">
					<tr> 
						<td> 
							<form method="post" action="" name="" class="pcForms">
							<b> 
							<% Response.Write("<font size=2 face=arial>Page "& iPageCurrent & " of "& iPageCount & "</font><P>")%>
							<% 'Display Next / Prev buttons
							if iPageCurrent > 1 then
								'We are not at the beginning, show the prev button %>
								<a href="FedEx_ManageShipmentsResults.asp?id=<%=pcv_intOrder%>&TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent-1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a> 
							<% end If
							If iPageCount <> 1 then
								For I=1 To iPageCount
									If I=iPageCurrent Then %>
										<%=I%> 
									<% Else %>
										<a href="FedExWS_ManageShipmentsResults.asp?id=<%=pcv_intOrder%>&TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=I%>&order=<%=strORD%>&sort=<%=strSort%>"><%=I%></a> 
									<% End If %>
								<% Next %>
							<% end if %>
							<% if CInt(iPageCurrent) <> CInt(iPageCount) then
								'We are not at the end, show a next link %>
								<a href="FedExWS_ManageShipmentsResults.asp?id=<%=pcv_intOrder%>&TypeSearch=<%=request.querystring("TypeSearch")%>&advquery=<%=request.querystring("advquery")%>&FromDate=<%=request.querystring ("FromDate") %>&ToDate=<%=request.querystring ("ToDate") %>&iPageCurrent=<%=iPageCurrent+1%>&order=<%=strORD%>&sort=<%=strSort%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
							<% end If 
							 %>
							</b> 
							</form>						</td>
					</tr>
				</table>			</td>
		</tr>
		<tr>
			<td colspan="11" style="text-align: center">
			<% 
			pcv_strPreviousPage = "Orddetails.asp?id=" & pcv_intOrder
			%>
				<input type="button" class="btn btn-default"  name="Button" value="Go Back To Order Details" onClick="document.location.href='<%=pcv_strPreviousPage%>'">
			<% 
			else 
			pcv_strPreviousPage = "Orddetails.asp?id=" & Request("id")
			%>
				<input type="button" class="btn btn-default"  name="Button" value="Go Back to Order Details" onClick="document.location.href='<%=pcv_strPreviousPage%>'">
			<% end if %>		
			</td>
	</tr>
		<tr>
		  <td colspan="11" style="text-align: center"><br /><%= pcf_FedExWriteLegalDisclaimers %></td>
		  </tr>
</table>
<%
'// DESTROY THE FEDEX OBJECT
set objFedExClass = nothing
%>
<!--#include file="AdminFooter.asp"-->