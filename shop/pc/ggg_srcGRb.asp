<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<%
%>
<!--#include file="header_wrapper.asp"-->
<% dim iPageSize
iPageSize="20"

dim iPageCurrent

if request.queryString("iPageCurrent")="" then
   iPageCurrent=1 
else
   iPageCurrent=server.HTMLEncode(request.querystring("iPageCurrent"))
	 if not validNum(iPageCurrent) then
	 	iPageCurrent=1
	 end if
end if
    
dim cname, clastname, emonth, eyear, pcint_nosearch



cname=getUserInput(request("cname"),0)
clastname=getUserInput(request("clastname"),0)
cregname=getUserInput(request("cregname"),0)
emonth=getUserInput(request("emonth"),0)
eyear=getUserInput(request("eyear"),0)

If cname="" and clastname="" and emonth="" and eyear="" and cregname="" then
	pcint_nosearch = 1
End if
%>
<div id="pcMain">
	<div class="pcMainContent">
		<h1><%= dictLanguage.Item(Session("language")&"_SrcGRb_1")%></h1>
		<%
		IF pcint_nosearch <> 1 THEN   'At least one search criteria has been provided
			query="select customers.name,customers.lastname,pcEvents.pcEv_Name,pcEvents.pcEv_Date,pcEvents.pcEv_Code from customers,pcEvents where customers.idcustomer=pcEvents.pcEv_IDCustomer and pcEvents.pcEv_Hide=0 and pcEvents.pcEv_Active=1"

			if cname<>"" then
				query=query & " AND customers.name like '%" & cname & "%' "
			end if

			if clastname<>"" then
				query=query & " AND customers.lastname like '%" & clastname & "%' "
			end if
	
			if cregname<>"" then
				query=query & " AND pcEvents.pcEv_Name like '%" & cregname & "%' "
			end if

			if emonth<>"" then
				query=query & " AND month(pcEvents.pcEv_Date)='" & emonth & "' "
			end if

			if eyear<>"" then
				query=query & " AND year(pcEvents.pcEv_Date)='" & eyear & "' "
			end if

			Set rsTemp=Server.CreateObject("ADODB.Recordset")     

			rstemp.CacheSize=iPageSize
			rstemp.PageSize=iPageSize

			rsTemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText
	
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsTemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			if rsTemp.eof then
				set rstemp=nothing
				call closeDb()
				response.redirect "msg.asp?message=97"
			end if

			dim iPageCount
			iPageCount=rstemp.PageCount

			If Cint(iPageCurrent) > Cint(iPageCount) Then Cint(iPageCurrent)=Cint(iPageCount)
			If Cint(iPageCurrent) < 1 Then iPageCurrent=Cint(1)

			rstemp.AbsolutePage=iPageCurrent

			dim count

			col_CustomerNameClass = "pcCol-3 pcSrcGRb_CustomerName"
			col_EventNameClass		= "pcCol-5 pcSrcGRb_EventName"
			col_EventDateClass		= "pcCol-2 pcSrcGRb_EventDate"
			col_ActionsClass			= "pcCol-2 pcSrcGRb_Actions"
			%>

			<div id="pcTableSrcGRb" class="pcTable">
				<div class="pcTableHeader">
					<div class="<%= col_CustomerNameClass %>"><%= dictLanguage.Item(Session("language")&"_SrcGRb_2")%></div>
					<div class="<%= col_EventNameClass %>"><%= dictLanguage.Item(Session("language")&"_SrcGRb_3")%></div>
					<div class="<%= col_EventDateClass %>"><%= dictLanguage.Item(Session("language")&"_SrcGRb_4")%></div>
					<div class="<%= col_ActionsClass %>"></div>
				</div>

				<%
				count=0
				do while not rsTemp.eof and count < rsTemp.pageSize
					count=count+1
					custname=rstemp("name") & " " & rstemp("lastname")
					gname=rstemp("pcEv_Name")
					gedate=rstemp("pcEv_Date")
					if year(gedate)="1900" then
						gedate=""
					end if
					if gedate<>"" then
						if scDateFrmt="DD/MM/YY" then
							gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
						else
							gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
						end if
					end if
					gCode=rstemp("pcEv_Code")
					%>
					<div class="pcTableRow">
						<div class="<%= col_CustomerNameClass %>"><%=custname%></div>
						<div class="<%= col_EventNameClass %>"><%=gname%></div>
						<div class="<%= col_EventDateClass %>"><%=gedate%></div>
						<div class="<%= col_ActionsClass %>">
							<a href="ggg_viewGR.asp?grcode=<%=gCode%>"><%= dictLanguage.Item(Session("language")&"_SrcGRb_5")%></a>
						</div>
					</div>
					<%
					rstemp.MoveNext
				loop
				set rstemp=nothing
				%>

			</div>

			<div class="pcFormItem">
				<div class="pcPageNav"> 
				<% if iPageCount > 1 then %>
					<%response.write(dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & iPageCount)%>
					&nbsp;-&nbsp;
					<% if iPageCurrent > 1 then %>
						<a href="ggg_srcGRb.asp?iPageSize=<%=iPageSize%>&iPageCurrent=<%=iPageCurrent - 1%>&cname=<%=cname%>&clastname=<%=clastname%>&emonth=<%=emonth%>&eyear=<%=eyear%>"><img src="<%=pcf_getImagePath("",rsIconObj("previousicon"))%>" border="0"></a> 
					<% end if
					For I=1 To iPageCount
						If I=iPageCurrent Then%>
							<b><%=I%></b>
						<% Else %>
							<a href="ggg_srcGRb.asp?iPageSize=<%=iPageSize%>&iPageCurrent=<%=I%>&cname=<%=cname%>&clastname=<%=clastname%>&emonth=<%=emonth%>&eyear=<%=eyear%>"><%=I%></a> 
						<% End If %>
					<% Next %>
					<% if cInt(iPageCurrent) <> cInt(iPageCount) then %>
						<a href="ggg_srcGRb.asp?iPageSize=<%=iPageSize%>&iPageCurrent=<%=iPageCurrent + 1%>&cname=<%=cname%>&clastname=<%=clastname%>&emonth=<%=emonth%>&eyear=<%=eyear%>"><img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>" border="0"></a> 
					<% end if 
				end if%>
				</div>
			</div>

			<div class="pcSpacer"></div>

			<div class="pcFormButtons">
				<a class="pcButton pcButtonBack" href="ggg_srcGR.asp">
					<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
				</a>
			</div>
		<%ELSE ' No search criteria has been provided %>
			<div class="pcFormItem">
				<div class="pcFormItemFull"><% response.write dictLanguage.Item(Session("language")&"_advSrcb_1") %></div>
			</div>

			<div class="pcSpacer"></div>

			<div class="pcFormButtons">
				<a class="pcButton pcButtonBack" href="ggg_srcGR.asp">
					<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
				</a>
			</div>
		<%END IF%>

	</div>
</div>

<!--#include file="footer_wrapper.asp"-->
