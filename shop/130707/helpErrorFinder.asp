<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Online Help - Error Finder" %>
<% Section="" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<% 
on error resume next

Dim intFormSubmitted

intFormSubmitted=0

if request("Action")="DEL" then
	intRefId=request("RefID")
	
	query="DELETE FROM pcErrorHandler WHERE pcErrorHandler_ID="&intRefId&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
	
	call closeDb()
response.redirect "helpErrorFinder.asp?msg=8"
end if

'Check if form has been submitted
IF request("submit") <> "" THEN
	intFormSubmitted=1
	pcIntCustRefID = trim(request("CustRefID"))
	if pcIntCustRefID="" then
		call closeDb()
response.redirect "helpErrorFinder.asp?msg=1"
	end if
	
	query="SELECT pcErrorHandler_ID, pcErrorHandler_SessionID, pcErrorHandler_RequestMethod, pcErrorHandler_ServerPort, pcErrorHandler_HTTPS, pcErrorHandler_LocalAddr, pcErrorHandler_RemoteAddr, pcErrorHandler_UserAgent, pcErrorHandler_URL, pcErrorHandler_HttpHost, pcErrorHandler_HttpLang, pcErrorHandler_ErrNumber, pcErrorHandler_ErrSource, pcErrorHandler_ErrDescription, pcErrorHandler_InsertDate, pcErrorHandler_CustomerRefID FROM pcErrorHandler WHERE (((pcErrorHandler_CustomerRefID)='" & pcIntCustRefID &"'));"

	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs = conntemp.execute(query)
	if rs.EOF then
		call closeDb()
response.redirect "helpErrorFinder.asp?msg=2"
	else
		pcErrorHandler_ID=rs("pcErrorHandler_ID")
		pcErrorHandler_SessionID=rs("pcErrorHandler_SessionID")
		pcErrorHandler_RequestMethod=rs("pcErrorHandler_RequestMethod")
		pcErrorHandler_ServerPort=rs("pcErrorHandler_ServerPort")
		pcErrorHandler_HTTPS=rs("pcErrorHandler_HTTPS")
		pcErrorHandler_LocalAddr=rs("pcErrorHandler_LocalAddr")
		pcErrorHandler_RemoteAddr=rs("pcErrorHandler_RemoteAddr")
		pcErrorHandler_UserAgent=rs("pcErrorHandler_UserAgent")
		pcErrorHandler_URL=rs("pcErrorHandler_URL")
		pcErrorHandler_HttpHost=rs("pcErrorHandler_HttpHost")
		pcErrorHandler_HttpLang=rs("pcErrorHandler_HttpLang")
		pcErrorHandler_ErrNumber=rs("pcErrorHandler_ErrNumber")
		pcErrorHandler_ErrSource=rs("pcErrorHandler_ErrSource")
		pcErrorHandler_ErrDescription=rs("pcErrorHandler_ErrDescription")
		pcErrorHandler_ErrDescription = replace(pcErrorHandler_ErrDescription,"""""","""")
		pcErrorHandler_InsertDate=rs("pcErrorHandler_InsertDate")
		pcErrorHandler_CustomerRefID=rs("pcErrorHandler_CustomerRefID")
	end if
	set rs=nothing
	
	%>
	<table class="pcCPcontent">
		<tr>
			<td>
				Customer Reference Number: <%=pcErrorHandler_CustomerRefID%>
			</td>
		</tr>
		<tr> 
			<td>
				Date/Time: <%=pcErrorHandler_InsertDate%>
			</td>
		</tr>
		<tr>
			<td>
				Session ID: <%=pcErrorHandler_SessionID%>
			</td>
		</tr>
		<tr>
			<td>
				Error Number: <%=pcErrorHandler_ErrNumber%>
			</td>
		</tr>
		<tr>
			<td>
				Source: <%=pcErrorHandler_ErrSource%>
			</td>
		</tr>
		<tr>
			<td>
				Description: <%=pcErrorHandler_ErrDescription%>
			</td>
		</tr>
		<tr>
			<td>
				Request Method: <%=pcErrorHandler_RequestMethod%>
			</td>
		</tr>
		<tr>
			<td>
				Server Port: <%=pcErrorHandler_ServerPort%>
			</td>
		</tr>
		<tr>
			<td>
				HTTPS: <%=pcErrorHandler_HTTPS%>
			</td>
		</tr>
		<tr>
			<td>
				Local Address: <%=pcErrorHandler_LocalAddr%>
			</td>
		</tr>
		<tr>
			<td>
				Host Address: <%=pcErrorHandler_RemoteAddr%>
			</td>
		</tr>
		<tr>
			<td>
				User Agent: <%=pcErrorHandler_UserAgent%>
			</td>
		</tr>
		<tr>
			<td>
				URL: <%=pcErrorHandler_URL%>
			</td>
		</tr>
		<tr>
			<td>
				HTTP Headers: <%=pcErrorHandler_HttpHost%>, <%=pcErrorHandler_HttpLang%>
			</td>
		</tr>
	</table>
<% End if

if request("submit3")<>"" then
	intFormSubmitted=1
	strViewLog=request("ViewLog")
	
	'//Today's date
	pcDtToday=Date() 'MM/DD/YYYY
	
	
	select case strViewLog
	
		case "DAY"
			'//GET ALL FOR THE CURRENT DATE
			dtFromDate=Date()
			if SQL_Format="1" then
				FromDate=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
			else
				FromDate=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			end if
			query1= query1 & " pcErrorHandler_InsertDate>='" & FromDate & "'"
			query="SELECT pcErrorHandler_ID, pcErrorHandler_SessionID, pcErrorHandler_RequestMethod, pcErrorHandler_ServerPort, pcErrorHandler_HTTPS, pcErrorHandler_LocalAddr, pcErrorHandler_RemoteAddr, pcErrorHandler_UserAgent, pcErrorHandler_URL, pcErrorHandler_HttpHost, pcErrorHandler_HttpLang, pcErrorHandler_ErrNumber, pcErrorHandler_ErrSource, pcErrorHandler_ErrDescription, pcErrorHandler_InsertDate, pcErrorHandler_CustomerRefID FROM pcErrorHandler WHERE"&query1
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

		case "WEEK"
			'//GET ALL FOR THE LAST 7 DAYS
			dtFromDate=Date()-7
			dtFromDateStr=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			if SQL_Format="1" then
				FromDate=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
			else
				FromDate=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			end if
			dtToDate=Date()
			dtToDateStr=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			if SQL_Format="1" then
				ToDate=day(dtToDate) & "/" & month(dtToDate) & "/" & year(dtToDate)
			else
				ToDate=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			end if
			
			if (FromDate<>"") and (IsDate(FromDate)) then
				query1= query1 & " pcErrorHandler_InsertDate>='" & FromDate & "'"
			end if

			if (ToDate<>"") and (IsDate(ToDate)) then
				query1= query1 & " AND pcErrorHandler_InsertDate<='" & ToDate & "'"
			end if

			query="SELECT pcErrorHandler_ID, pcErrorHandler_SessionID, pcErrorHandler_RequestMethod, pcErrorHandler_ServerPort, pcErrorHandler_HTTPS, pcErrorHandler_LocalAddr, pcErrorHandler_RemoteAddr, pcErrorHandler_UserAgent, pcErrorHandler_URL, pcErrorHandler_HttpHost, pcErrorHandler_HttpLang, pcErrorHandler_ErrNumber, pcErrorHandler_ErrSource, pcErrorHandler_ErrDescription, pcErrorHandler_InsertDate, pcErrorHandler_CustomerRefID FROM pcErrorHandler WHERE"&query1
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

		case "MONTH"
			'//GET ALL FOR THE LAST 7 DAYS
			dtFromDate=Date()-30
			dtFromDateStr=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			if SQL_Format="1" then
				FromDate=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
			else
				FromDate=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			end if
			dtToDate=Date()
			dtToDateStr=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			if SQL_Format="1" then
				ToDate=day(dtToDate) & "/" & month(dtToDate) & "/" & year(dtToDate)
			else
				ToDate=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			end if
			if (FromDate<>"") and (IsDate(FromDate)) then
				query1= query1 & " pcErrorHandler_InsertDate>='" & FromDate & "'"
			end if

			if (ToDate<>"") and (IsDate(ToDate)) then
				query1= query1 & " AND pcErrorHandler_InsertDate<='" & ToDate & "'"
			end if

			query="SELECT pcErrorHandler_ID, pcErrorHandler_SessionID, pcErrorHandler_RequestMethod, pcErrorHandler_ServerPort, pcErrorHandler_HTTPS, pcErrorHandler_LocalAddr, pcErrorHandler_RemoteAddr, pcErrorHandler_UserAgent, pcErrorHandler_URL, pcErrorHandler_HttpHost, pcErrorHandler_HttpLang, pcErrorHandler_ErrNumber, pcErrorHandler_ErrSource, pcErrorHandler_ErrDescription, pcErrorHandler_InsertDate, pcErrorHandler_CustomerRefID FROM pcErrorHandler WHERE"&query1
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

		case "YEAR"
			'//GET ALL FOR THE LAST 7 DAYS
			dtFromDate=Date()-365
			dtFromDateStr=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			if SQL_Format="1" then
				FromDate=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
			else
				FromDate=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			end if
			dtToDate=Date()
			dtToDateStr=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			if SQL_Format="1" then
				ToDate=day(dtToDate) & "/" & month(dtToDate) & "/" & year(dtToDate)
			else
				ToDate=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			end if
			if (FromDate<>"") and (IsDate(FromDate)) then
				query1= query1 & " pcErrorHandler_InsertDate>='" & FromDate & "'"
			end if

			if (ToDate<>"") and (IsDate(ToDate)) then
				query1= query1 & " AND pcErrorHandler_InsertDate<='" & ToDate & "'"
			end if

			query="SELECT pcErrorHandler_ID, pcErrorHandler_SessionID, pcErrorHandler_RequestMethod, pcErrorHandler_ServerPort, pcErrorHandler_HTTPS, pcErrorHandler_LocalAddr, pcErrorHandler_RemoteAddr, pcErrorHandler_UserAgent, pcErrorHandler_URL, pcErrorHandler_HttpHost, pcErrorHandler_HttpLang, pcErrorHandler_ErrNumber, pcErrorHandler_ErrSource, pcErrorHandler_ErrDescription, pcErrorHandler_InsertDate, pcErrorHandler_CustomerRefID FROM pcErrorHandler WHERE"&query1
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
		case "ALL"
			query="SELECT * FROM pcErrorHandler;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

	end select
	
	if NOT rs.eof then
	'//Show records returned
	%>
	<table class="pcCPcontent">
			<tr>
				<th nowrap>Date</th>
			  <th nowrap>Reference Number</th>
			  <th nowrap>Error Number</th>
			  <th nowrap>IP Number</th>
			  <th nowrap>URL</th>
			  <th nowrap align="center">Delete</th>
			</tr>
			<% do until rs.eof
				pcErrorHandler_ID=rs("pcErrorHandler_ID")
				pcErrorHandler_SessionID=rs("pcErrorHandler_SessionID")
				pcErrorHandler_RequestMethod=rs("pcErrorHandler_RequestMethod")
				pcErrorHandler_ServerPort=rs("pcErrorHandler_ServerPort")
				pcErrorHandler_HTTPS=rs("pcErrorHandler_HTTPS")
				pcErrorHandler_LocalAddr=rs("pcErrorHandler_LocalAddr")
				pcErrorHandler_RemoteAddr=rs("pcErrorHandler_RemoteAddr")
				pcErrorHandler_UserAgent=rs("pcErrorHandler_UserAgent")
				pcErrorHandler_URL=rs("pcErrorHandler_URL")
				pcErrorHandler_HttpHost=rs("pcErrorHandler_HttpHost")
				pcErrorHandler_HttpLang=rs("pcErrorHandler_HttpLang")
				pcErrorHandler_ErrNumber=rs("pcErrorHandler_ErrNumber")
				pcErrorHandler_ErrSource=rs("pcErrorHandler_ErrSource")
				pcErrorHandler_ErrDescription=rs("pcErrorHandler_ErrDescription")
				pcErrorHandler_InsertDate=rs("pcErrorHandler_InsertDate")
				pcErrorHandler_CustomerRefID=rs("pcErrorHandler_CustomerRefID")
				%>
				<tr>
					<td nowrap><%=pcErrorHandler_InsertDate%></td>
					<td nowrap><a href="helpErrorFinder.asp?submit=Y&CustRefID=<%=pcErrorHandler_CustomerRefID%>"><%=pcErrorHandler_CustomerRefID%></a></td>
					<td nowrap><%=pcErrorHandler_ErrNumber%></td>
					<td nowrap><%=pcErrorHandler_RemoteAddr%></td>
					<td nowrap><%=pcErrorHandler_URL%></td>
					<td nowrap align="center"><a href="javascript:if (confirm('Are you sure you want to delete this error log. This action is permanent and cannot be reversed.')) location='helpErrorFinder.asp?Action=DEL&RefID=<%=pcErrorHandler_ID%>'"><img src="images/delete2.gif" alt="Delete Error" width="23" height="18" border="0"></a></td>
			  </tr><% rs.moveNext
			loop %>
		</table>
	<% else %>
	<table class="pcCPcontent">
		<tr>
		  <td><div class="pcCPmessage">No error logs were found.</div></td>
		</tr>
	</table>
	<% end if
	set rs=nothing
	

end if


if request("submit2")<>"" then
	intFormSubmitted=1
	strPurgeLog=request("PurgeLog")
	
	'//Today's date
	pcDtToday=Date() 'MM/DD/YYYY
	
	
	select case strPurgeLog
	
		case "DAY"
			'//GET ALL FOR THE CURRENT DATE
			dtFromDate=Date()
			if SQL_Format="1" then
				FromDate=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
			else
				FromDate=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			end if
			query1= query1 & " pcErrorHandler_InsertDate='" & FromDate & "'"

			query="DELETE FROM pcErrorHandler WHERE"&query1
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs=nothing
			
			call closeDb()
response.redirect "helpErrorFinder.asp?msg=3"
		case "WEEK"
			'//GET ALL FOR THE LAST 7 DAYS
			dtFromDate=Date()-7
			dtFromDateStr=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			if SQL_Format="1" then
				FromDate=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
			else
				FromDate=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			end if
			dtToDate=Date()
			dtToDateStr=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			if SQL_Format="1" then
				ToDate=day(dtToDate) & "/" & month(dtToDate) & "/" & year(dtToDate)
			else
				ToDate=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			end if
			
			if (FromDate<>"") and (IsDate(FromDate)) then
				query1= query1 & " pcErrorHandler_InsertDate>='" & FromDate & "'"
			end if

			if (ToDate<>"") and (IsDate(ToDate)) then
				query1= query1 & " AND pcErrorHandler_InsertDate<='" & ToDate & "'"
			end if

			query="DELETE FROM pcErrorHandler WHERE"&query1
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs=nothing
			
			call closeDb()
response.redirect "helpErrorFinder.asp?msg=4"
		case "MONTH"
			'//GET ALL FOR THE LAST 7 DAYS
			dtFromDate=Date()-30
			dtFromDateStr=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			if SQL_Format="1" then
				FromDate=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
			else
				FromDate=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			end if
			dtToDate=Date()
			dtToDateStr=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			if SQL_Format="1" then
				ToDate=day(dtToDate) & "/" & month(dtToDate) & "/" & year(dtToDate)
			else
				ToDate=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			end if
			if (FromDate<>"") and (IsDate(FromDate)) then
				query1= query1 & " pcErrorHandler_InsertDate>='" & FromDate & "'"
			end if

			if (ToDate<>"") and (IsDate(ToDate)) then
				query1= query1 & " AND pcErrorHandler_InsertDate<='" & ToDate & "'"
			end if

			query="DELETE FROM pcErrorHandler WHERE"&query1
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs=nothing
			
			call closeDb()
response.redirect "helpErrorFinder.asp?msg=5"
		case "YEAR"
			'//GET ALL FOR THE LAST 7 DAYS
			dtFromDate=Date()-365
			dtFromDateStr=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			if SQL_Format="1" then
				FromDate=day(dtFromDate) & "/" & month(dtFromDate) & "/" & year(dtFromDate)
			else
				FromDate=month(dtFromDate) & "/" & day(dtFromDate) & "/" & year(dtFromDate)
			end if
			dtToDate=Date()
			dtToDateStr=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			if SQL_Format="1" then
				ToDate=day(dtToDate) & "/" & month(dtToDate) & "/" & year(dtToDate)
			else
				ToDate=month(dtToDate) & "/" & day(dtToDate) & "/" & year(dtToDate)
			end if
			if (FromDate<>"") and (IsDate(FromDate)) then
				query1= query1 & " pcErrorHandler_InsertDate>='" & FromDate & "'"
			end if

			if (ToDate<>"") and (IsDate(ToDate)) then
				query1= query1 & " AND pcErrorHandler_InsertDate<='" & ToDate & "'"
			end if

			query="DELETE FROM pcErrorHandler WHERE"&query1
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs=nothing
			
			call closeDb()
response.redirect "helpErrorFinder.asp?msg=6"
		case "ALL"
			query="DELETE FROM pcErrorHandler;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				
				call closeDb()
response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs=nothing
			
			call closeDb()
response.redirect "helpErrorFinder.asp?msg=7"
	end select
		
	
end if
if intFormSubmitted=0 then %>
	<form action="helpErrorFinder.asp" method="post" class="pcForms">
		<input type="hidden" name="action" value="1">
		<table class="pcCPcontent">
			<tr>
				<td class="pcCPspacer"></td>
			</tr>
			<%
			strMsg=request("msg")
			if request("msg")<>"" then
				select case strMsg
					case "1"
						msg="The Reference Number that you entered is not valid. Please enter a valid Reference Number."
					case "2"
						msg="There is no error information associated with the Reference Number that you entered."
					case "3"
						msg="All errors that had been logged earlier today have been successfully purged."
						msgtype=1
					case "4"
						msg="All errors in the selected date range have been successfully purged."
						msgtype=1
					case "5"
						msg="All errors in the selected date range have been successfully purged."
						msgtype=1
					case "6"
						msg="All errors in the selected date range have been successfully purged."
						msgtype=1
					case "7"
						msg="All errors have been successfully purged."
						msgtype=1
					case "8"
						msg="The selected error has been successfully deleted."
						msgtype=1
				end select
 %>
				<tr>
					<td>
                    <% ' START show message, if any %>
                        <!--#include file="pcv4_showMessage.asp"-->
                    <% 	' END show message %>
				</tr>
				<tr>
					<td class="pcCPspacer"></td>
				</tr>
			<% end if %>
			<tr>
				<td>When a storefront page returns an error, ProductCart saves the error information into the database. For security reasons, no details are shown to the user. To retrieve details on the error, enter the reference number provided to you in the storefront by the page that returned the error, or sent to you by the customer that reported the problem. Please keep in mind that many errors (esp. Type Mismatches) are actually the result of a security service (such as McAfee Security Scan or Trustwave) scanning the server for potential vulnerabilities or SQL injection exploits, and triggering the error. If you're not hearing directly from customers about problems on your store, then most of the errors can be disregarded, but if you are able to replicate an error yourself, please <a href="https://www.productcart.com/store/pc/custpref.asp" target="_blank"> login to your account to submit a ticket </a>.</td>
			</tr>
			<tr>
				<td>Enter Reference Number: <input type="text" value="" name="CustRefID"></td>
			</tr>
			<tr>
				<td class="pcCPspacer"></td>
			</tr>
			<tr>
				<td><input type="submit" value="Locate Error Information" name="submit" class="btn btn-primary"></td>
			</tr>
			<tr>
			  <td>&nbsp;</td>
			  </tr>
			<tr>
			  <th>Purge Error Logs </th>
			</tr>
			<tr>
				<td class="pcCPspacer"></td>
			</tr>
			<tr>
			  <td>Purge all errors from the last: 
			    <select name="PurgeLog">
			      <option value="DAY">Today</option>
			      <option value="WEEK">Week</option>
			      <option value="MONTH">Month</option>
			      <option value="YEAR">Year</option>
			      <option value="ALL">Purge entire log</option>
			      </select>			    </td>
			  </tr>
			<tr>
			  <td><input type="submit" value="Purge Error Log" name="submit2" class="btn btn-primary" onclick="return(confirm('You are about to purge the selected error logs. Are you sure you want to continue?'));"></td>
			  </tr>
			<tr>
			  <td>&nbsp;</td>
			  </tr>
			<tr>
			  <th>View Error Logs </th>
			</tr>
			<tr>
				<td class="pcCPspacer"></td>
			</tr>
			<tr>
			  <td>View all errors from the last: 
			    <select name="ViewLog">
			      <option value="DAY">Today</option>
			      <option value="WEEK">Week</option>
			      <option value="MONTH">Month</option>
			      <option value="YEAR">Year</option>
			      <option value="ALL">View entire log</option>
			      </select>
                </td>
			</tr>
			<tr>
			  <td><input type="submit" value="View Error Log" name="submit3" class="btn btn-primary"></td>
			  </tr>
		</table>
	</form>
<% END IF %>
<!--#include file="AdminFooter.asp"-->
