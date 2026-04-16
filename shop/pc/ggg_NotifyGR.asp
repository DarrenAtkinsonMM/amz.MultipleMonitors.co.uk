<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="CustLIv.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
dim pid

pIdCustomer=session("idCustomer")

query="SELECT name,lastName,email FROM customers WHERE idCustomer=" &pIdCustomer
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
CustName=rs("name") & " " & rs("lastname")
CustEmail=rs("email")

set rs=nothing

gIDEvent=getUserInput(request("IDEvent"),0)

if gIDEvent<>"" then
	query="select pcEv_Name,pcEv_Date,pcEv_Code from pcEvents where pcEv_IDCustomer=" & pIDCustomer & " and pcEv_IDEvent=" & gIDEvent
	set rstemp=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rstemp.eof then
		set rstemp=nothing
		call closedb()
		response.redirect "ggg_manageGRs.asp"
	else
		geName=rstemp("pcEv_Name")
		geDate=rstemp("pcEv_Date")
		if gedate<>"" then
			if scDateFrmt="DD/MM/YY" then
				gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
			else
				gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
			end if
		end if
		gCode=rstemp("pcEv_Code")
	
		SPath1=Request.ServerVariables("PATH_INFO")
		mycount1=0
		do while mycount1<2
			if mid(SPath1,len(SPath1),1)="/" then
				mycount1=mycount1+1
			end if
			if mycount1<2 then
				SPath1=mid(SPath1,1,len(SPath1)-1)
			end if
		loop
		SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1
	
		if Right(SPathInfo,1)="/" then
			gStr=SPathInfo & "pc/ggg_viewGR.asp?grcode=" & gCode
		else
			gStr=SPathInfo & "/pc/ggg_viewGR.asp?grcode=" & gCode
		end if
	
		gmsg=gename & vbcrlf & dictLanguage.Item(Session("language")&"_NotifyGR_10") & gedate & vbcrlf & dictLanguage.Item(Session("language")&"_NotifyGR_11") & gStr
	
	end if
	set rstemp=nothing
end if

%>
<script type=text/javascript>

function Form1_Validator(theForm)
{
	if (theForm.yourname.value == "")
  	{
			alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		    theForm.yourname.focus();
		    return (false);
	}

	if (theForm.youremail.value == "")
  	{
			alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		    theForm.youremail.focus();
		    return (false);
	}
	
	if (theForm.friendsemail.value == "")
  	{
			alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		    theForm.friendsemail.focus();
		    return (false);
	}
	
	if (theForm.title.value == "")
  	{
			alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		    theForm.title.focus();
		    return (false);
	}
	
	if (theForm.message.value == "")
  	{
			alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		    theForm.message.focus();
		    return (false);
	}
	
return (true);
}
</script>

<div id="pcMain">
	<div class="pcMainContent">
		<form name="Form1" action="ggg_NotifyGRb.asp?action=send" method="POST" onSubmit="return Form1_Validator(this)" class="pcForms">
			<h1><%= dictLanguage.Item(Session("language")&"_NotifyGR_1")%>"<%=geName%>"<%= dictLanguage.Item(Session("language")&"_NotifyGR_1a")%></h1>
			<div class="pcFormItem">
				<div class="pcFormItemFull">
					<%= dictLanguage.Item(Session("language")&"_NotifyGR_2")%>
					
					<input type="hidden" name="pid" value="<%=pid%>">
					<input type="hidden" name="pname" value="<%=productname%>">
				</div>
			</div>

			<div class="pcSpacer"></div>

			<% '// Customer Name %>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_NotifyGR_3")%></div>
				<div class="pcFormField"><input type="text" size="40" name="yourname" value="<%=CustName%>"></div>
			</div>

			<% '// Customer Email %>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_NotifyGR_4")%></div>
				<div class="pcFormField"><input type="text" size="40" name="youremail" value="<%=CustEmail%>"></div>
			</div>

			<% '// Email Address(es) %>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_NotifyGR_5")%></div>
				<div class="pcFormField">
					<textarea name="friendsemail" rows="8" cols="40"></textarea>
					<br><i><%= dictLanguage.Item(Session("language")&"_NotifyGR_6")%></i>
				</div>
			</div>

			<div class="pcSpacer"></div>
			
			<% '// Subject/Title %>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_NotifyGR_7")%></div>
				<div class="pcFormField"><input type="text" size="40" name="title"></div>
			</div>

			<% '// Message %>
			<div class="pcFormItem">
				<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_NotifyGR_8")%></div>
				<div class="pcFormField"><textarea rows="15" cols="40" name="message"><%=gmsg%></textarea></div>
			</div>
			
			<div class="pcSpacer"></div>
			
			<div class="pcFormButtons">
				<button class="pcButton pcButtonSendMessages" id="submit" name="Submit" value="<%= dictLanguage.Item(Session("language")&"_NotifyGR_12")%>">
					<img src="<%=pcf_getImagePath("",rslayout("SendMsgs"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_sendmsgs") %>">
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_sendmsgs") %></span>
				</button>

				<a class="pcButton pcButtonBack" href="ggg_manageGRs.asp">
					<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
				</a>
			</div>
		</form>

	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
