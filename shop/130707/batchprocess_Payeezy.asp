<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="inc_GenDownloadInfo.asp"-->
<!--#include file="inc_PayeezyFunctions.asp"-->
<%Dim pageTitle, Section
pageTitle="Batch Process Payeezy Orders"
pageIcon="pcv4_icon_orders.gif"
Section="orders"%>
<!--#include file="adminHeader.asp" -->
<%'on error resume next
Dim successCnt, successData
Dim failedCnt, failedData

successCnt=0
successData=""
failedCnt=0
failedData=""

'// How many checkboxes?
checkboxCnt=request.Form("PEYcheckboxCnt")

'////////////////////////////////////////////////////
'// START: Process Selected Orders
'////////////////////////////////////////////////////
dim r
For r=1 to checkboxCnt
	if request.Form("checkOrd"&r)="YES" then
		pOrderStatus=request.Form("orderstatus"&r)
		pCheckEmail=request.Form("checkEmail"&r)
		pIdOrder=Request.Form("idOrder"&r)  & ""
		qry_ID=pIdOrder
		pcv_CustomerReceived=0
		pcv_AdmComments=""
		pcv_SubmitType=3
		
		queryQ="SELECT pcPEYLg_Status FROM pcPayeezyLogs WHERE idOrder=" & pIdOrder & ";"
		set rsQ=connTemp.execute(queryQ)
		NeedProcess=0
		FailedRc=0
		if rsQ.eof then
			queryQ="SELECT idCustomer FROM Orders WHERE idOrder=" & pIdOrder & ";"
			set rsQ=connTemp.execute(queryQ)
			tmpIDCust=0
			if not rsQ.eof then
				tmpIDCust=rsQ("idCustomer")
			end if
			set rsQ=nothing
			queryQ="INSERT INTO pcPayeezyLogs (idOrder,idCustomer,pcPEYLg_Status) VALUES (" & pIdOrder & "," & tmpIDCust & ",0);"
			set rsQ=connTemp.execute(queryQ)
			set rsQ=nothing
			NeedProcess=1
		else
			tmpStatus=rsQ("pcPEYLg_Status")
			set rsQ=nothing
			Select Case tmpStatus
				Case "1":
					queryQ="UPDATE Orders SET pcOrd_PaymentStatus=2 WHERE idOrder=" & pIdOrder & ";"
					set rsQ=connTemp.execute(queryQ)
					set rsQ=nothing
					NeedProcess=0
					FailedRc=0
				Case "0":
					NeedProcess=1
				Case "2":
					NeedProcess=0
					failedCnt=failedCnt+1
					FailedRc=1
					failedData=failedData & "Order Number "& (int(pIdOrder)+scpre) &" cannot be processed because the Payeezy payment was voided.<BR>"
			End Select	
		end if
		set rsQ=nothing
		
		IF NeedProcess=1 THEN
			tmpResult=CaptureVoidPayeezy(pIdOrder,1)
			if tmpResult=false then
				failedCnt=failedCnt+1
				FailedRc=1
				failedData=failedData & "Order Number "& (int(pIdOrder)+scpre) &" cannot be processed because the Payeezy payment cannot be captured.<BR>"
			else
				queryQ="UPDATE Orders SET pcOrd_PaymentStatus=2 WHERE idOrder=" & pIdOrder & ";"
				set rsQ=connTemp.execute(queryQ)
				set rsQ=nothing
				FailedRc=0
			end if
		END IF
		
		IF FailedRc=0 THEN
		'// START:  Process Order and Send Notification E-mails
		%>  <!--#include file="inc_ProcessOrder.asp"-->  <%
		'// END:  Process Order and Send Notification E-mails
		
		successCnt=successCnt+1
		successData=successData&"Order Number "& (int(pIdOrder)+scpre) &" was processed successfully<BR>"
		
		END IF
	
	end if
Next
'////////////////////////////////////////////////////
'// END: Process Selected Orders
'////////////////////////////////////////////////////


%>
<table class="pcCPcontent">
<%if successCnt>0 then%>
<tr>
	<td><div class="pcCPmessageSuccess"><%=successCnt%> records were successfully processed.</div>
		<% if successData<>"" then %>
			<br><%=successData%><br>
		<% end if %>
	</td>
</tr>
<%end if%>
<%if failedCnt>0 then%>
<tr>
	<td><div class="pcCPmessageWarning"><%=failedCnt%> records cannot be processed.</div>
		<% if failedData<>"" then %>
			<br><%=failedData%><br>
		<% end if %>
	</td>
</tr>
<%end if%>
<%if successCnt+failedCnt=0 then%>
<tr>
	<td>
		<div class="pcCPmessage">Please select orders to batch process.</div>
	</td>
</tr>
<%end if%>
<tr>
	<td>
    	<p>&nbsp;</p>
	    <p><a href="resultsAdvancedAll.asp?B1=View%2BAll&dd=1">Manage Orders</a></p>
	</td>
</tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>

<!--#include file="adminFooter.asp" -->
