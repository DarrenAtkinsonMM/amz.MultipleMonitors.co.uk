<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="CustLIv.asp"-->
<% 
query="select pcEv_IDEvent,pcEv_Name,pcEv_Type,pcEv_Date,pcEv_Hide,pcEv_Active from pcEvents where pcEv_IDCustomer=" & Session("idcustomer")
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

%> 
<!--#include file="header_wrapper.asp"-->
<div id="pcMain">
	<div class="pcMainContent">
		<h1><%= dictLanguage.Item(Session("language")&"_ManageGRs_1")%></h1>
		<% if rstemp.eof then%>
			<div class="pcErrorMessage">
				<%= dictLanguage.Item(Session("language")&"_ManageGRs_3")%>
			</div>
		<%else%>
			<%
				'// Column Classes
				col_EventNameClass	= "pcCol-3 pcManageGRs_EventName"
				col_EventTypeClass	= "pcCol-2 pcManageGRs_EventType"
				col_EventDateClass	= "pcCol-2 pcManageGRs_EventDate"
				col_VisibilityClass	= "pcCol-1 pcManageGRs_Visibility"
				col_StatusClass			= "pcCol-1 pcManageGRs_Status"
				col_ItemsClass			= "pcCol-1 pcManageGRs_Items"
				col_OptionsClass		= "pcCol-2 pcManageGRs_Options"
			%>
			<div id="pcTableManageGRs" class="pcTable">
				<div class="pcTableHeader">
					<div class="<%= col_EventNameClass %>"><%= dictLanguage.Item(Session("language")&"_ManageGRs_4")%></div>
					<div class="<%= col_EventTypeClass %>"><%= dictLanguage.Item(Session("language")&"_ManageGRs_5")%></div>
					<div class="<%= col_EventDateClass %>"><%= dictLanguage.Item(Session("language")&"_ManageGRs_6")%></div>
					<div class="<%= col_VisibilityClass %>"><%= dictLanguage.Item(Session("language")&"_ManageGRs_7")%></div>
					<div class="<%= col_StatusClass %>"><%= dictLanguage.Item(Session("language")&"_ManageGRs_8")%></div>
					<div class="<%= col_ItemsClass %>"><%= dictLanguage.Item(Session("language")&"_ManageGRs_9")%></div>
					<div class="<%= col_OptionsClass %>">&nbsp;</div>
				</div>
				<div class="pcSpacer"></div>
		<%do while not rstemp.eof
			gIDEvent=rstemp("pcEv_IDEvent")
			gType=rstemp("pcEv_Type")
			if gType<>"" then
			else
				gType="N/A"
			end if
			gName=rstemp("pcEv_Name")
			gDate=rstemp("pcEv_Date")
			if year(gDate)="1900" then
				gDate=""
			end if
			if gDate<>"" then
				if scDateFrmt="DD/MM/YY" then
					gDate=(day(gDate)&"/"&month(gDate)&"/"&year(gDate))
				else
					gDate=(month(gDate)&"/"&day(gDate)&"/"&year(gDate))
				end if
					end if
					gHide=rstemp("pcEv_Hide")
						if gHide<>"" then
						else
							gHide="0"
						end if
						gActive=rstemp("pcEv_Active")
						if gActive<>"" then
						else
							gActive="0"
						end if
						query="select sum(pcEP_Qty) as gQty,sum(pcEP_HQty) as gHQty from pcEvProducts where pcEP_IDEvent=" & gIDEvent & " and pcEP_GC=0 group by pcEP_IDEvent"
						set rs1=connTemp.execute(query)
						if err.number<>0 then
					call LogErrorToDatabase()
					set rs1=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
						if not rs1.eof then
							gQty=rs1("gQty")
							gHQty=rs1("gHQty")
						else
							gQty="0"
							gHQty="0"
						end if
						set rs1=nothing
						if gQty<>"" then
						else
							gQty="0"
						end if
			        
						if gHQty<>"" then
						else
							gHQty="0"
						end if%>
						<div class="pcTableRow">
							<div class="<%= col_EventNameClass %>"><strong><%=gName%></strong></div>
							<div class="<%= col_EventTypeClass %>"><%=gType%></div>
							<div class="<%= col_EventDateClass %>"><%=gDate%></div>
							<div class="<%= col_VisibilityClass %>">
								<%if gHide="1" then%>
									<span style="color: #FF0000"><%= dictLanguage.Item(Session("language")&"_ManageGRs_7b")%></span>
								<%else%>
									<%= dictLanguage.Item(Session("language")&"_ManageGRs_7a")%>
								<%end if%>
							</div>
							<div class="<%= col_StatusClass %>">
								<%if gActive="1" then%>
									<%= dictLanguage.Item(Session("language")&"_ManageGRs_8a")%>
								<%else%>
									<span style="color: #FF0000"><%= dictLanguage.Item(Session("language")&"_ManageGRs_8b")%></span>
								<%end if%>
							</div>
							<div class="<%= col_ItemsClass %>"><%=gQty%> (<%=clng(gQty)-clng(gHQty)%>)</div>
							<div class="<%= col_OptionsClass %>">
								<div class="btn-group">
									<button type="button" class="btn btn-default dropdown-toggle" data-toggle="dropdown">
										<%= dictLanguage.Item(Session("language")&"_ManageGRs_15")%>
										<span class="caret"></span>
									</button>

									<ul class="dropdown-menu" role="menu">
										<li><a href="JavaScript:;" onClick="document.getElementById('addProductsInfo<%=gIDEvent%>').style.display='';"><span class="glyphicon glyphicon-plus"></span>&nbsp;<%= dictLanguage.Item(Session("language")&"_ManageGRs_14")%></a></li>
										<li><a href="ggg_EditGR.asp?IDEvent=<%=gIDEvent%>"><span class="glyphicon glyphicon-cog"></span>&nbsp;<%= dictLanguage.Item(Session("language")&"_ManageGRs_13")%></a></li>
										<li><a href="ggg_GRDetails.asp?IDEvent=<%=gIDEvent%>"><span class="glyphicon glyphicon-th-list"></span>&nbsp;<%= dictLanguage.Item(Session("language")&"_ManageGRs_12")%></a></li>
										<li><a href="ggg_NotifyGR.asp?IDEvent=<%=gIDEvent%>"><span class="glyphicon glyphicon-envelope"></span>&nbsp;<%= dictLanguage.Item(Session("language")&"_ManageGRs_11")%></a></li>
									</ul>
								</div>
							</div>
						</div>
							
						<div class="pcTableRowFull">                
							<div class="pcInfoMessage" id="addProductsInfo<%=gIDEvent%>" style="display: none;">
								<%= dictLanguage.Item(Session("language")&"_ManageGRs_20")%> <br /><br />
								<a href="default.asp"><%= dictLanguage.Item(Session("language")&"_ManageGRs_21")%></a> | 
								<a href="JavaScript:;" onClick="document.getElementById('addProductsInfo<%=gIDEvent%>').style.display='none';"><%= dictLanguage.Item(Session("language")&"_ManageGRs_22")%></a>
							</div>
						</div>
					<%rstemp.movenext
				loop
				set rstemp=nothing%>
				</div>
			<% end if 'not rstemp.eof
			%>
		<div class="pcFormButtons">
			<a class="pcButton pcButtonCreateRegistry" href="ggg_instGR.asp">
				<img src="<%=pcf_getImagePath("",rslayout("CreRegistry"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_creregistry") %>">
				<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_creregistry") %></span>
			</a>

			<a class="pcButton pcButtonBack" href="javascript:history.go(-1)">
				<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
				<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
			</a>
		</div>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
