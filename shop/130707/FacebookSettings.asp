<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="Facebook Store Settings"
pageIcon="pcv4_icon_settings.png"
%>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<% 
pcPageName="FacebookSettings.asp"

'/////////////////////////////////////////////////////
'// Retrieve current database data
'/////////////////////////////////////////////////////
%>
<!--#include file="pcAdminRetrieveSettings.asp"-->

<%if request("updateSettings")<>"" then
	pcStoreOn=request("StoreOn")
	pcStoreMsg=replace(request("StoreMsg"),"'","''")
	if IsNull(pcStoreMsg) OR pcStoreMsg="" then
		pcStoreMsg="This store has been temporarily closed."
	end if
	pcAppID=request("AppID")
	pcHomepage=request("HomePage")
	if IsNull(pcHomepage) OR pcHomepage="" then
		pcHomepage="viewcategories.asp"
	end if
	pcStoreHeader=replace(request("StoreHeader"),"'","''")
	pcStoreFooter=replace(request("StoreFooter"),"'","''")
	pcAppWidth=request("AppWidth")
	if IsNull(pcAppWidth) OR pcAppWidth="" OR pcAppWidth="0" then
		pcAppWidth=810
	end if
	pcCustomDisplay=request("UseDefault")
	if IsNull(pcCustomDisplay) OR pcCustomDisplay="" then
		pcCustomDisplay=1
	else
		pcCustomDisplay=0
	end if
	pcIntCatImages=request("catImages")
	if IsNull(pcIntCatImages) OR pcIntCatImages="" then
		pcIntCatImages=0
	end if
	pcIntCatRow=request("CatRow")
	if IsNull(pcIntCatRow) OR pcIntCatRow="" OR pcIntCatRow="0" then
		pcIntCatRow=3
	end if
	pcIntCatRowsPerPage=request("CatRowsperPage")
	if IsNull(pcIntCatRowsPerPage) OR pcIntCatRowsPerPage="" OR pcIntCatRowsPerPage="0" then
		pcIntCatRowsPerPage=3
	end if
	pcStrBType=request("BType")
	if IsNull(pcStrBType) OR pcStrBType="" then
		pcStrBType="h"
	else
		pcStrBType=lcase(pcStrBType)
	end if
	pcIntPrdRow=request("PrdRow")
	if IsNull(pcIntPrdRow) OR pcIntPrdRow="" OR pcIntPrdRow="0" then
		pcIntPrdRow=3
	end if
	pcIntPrdRowsPerPage=request("PrdRowsPerPage")
	if IsNull(pcIntPrdRowsPerPage) OR pcIntPrdRowsPerPage="" OR pcIntPrdRowsPerPage="0" then
		pcIntPrdRowsPerPage=3
	end if
	pcIntShowSKU=request("ShowSKU")
	if IsNull(pcIntShowSKU) OR pcIntShowSKU="" then
		pcIntShowSKU=-1
	end if
	pcIntShowSmallImg=request("ShowSmallImg")
	if IsNull(pcIntShowSmallImg) OR pcIntShowSmallImg="" then
		pcIntShowSmallImg=-1
	end if
	
	

	query="SELECT pcFBS_TurnOnOff FROM pcFacebookSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		query="UPDATE pcFacebookSettings SET pcFBS_TurnOnOff=" & pcStoreOn & ",pcFBS_OffMsg='" & pcStoreMsg & "',pcFBS_AppID='" & pcAppID & "',pcFBS_RedirectURL='" & pcHomepage & "',pcFBS_Header='" & pcStoreHeader & "',pcFBS_Footer='" & pcStoreFooter & "',pcFBS_PageWidth=" & pcAppWidth & ",pcFBS_CustomDisplay=" & pcCustomDisplay & ",pcFBS_CatImages=" & pcIntCatImages & ",pcFBS_CatRow=" & pcIntCatRow & ",pcFBS_CatRowsperPage=" & pcIntCatRowsPerPage & ",pcFBS_BType='" & pcStrBType & "',pcFBS_PrdRow=" & pcIntPrdRow & ",pcFBS_PrdRowsPerPage=" & pcIntPrdRowsPerPage & ",pcFBS_ShowSKU=" & pcIntShowSKU & ",pcFBS_ShowSmallImg=" & pcIntShowSmallImg & ";"
		set rs=connTemp.execute(query)
		set rs=nothing
	else
		query="INSERT INTO pcFacebookSettings (pcFBS_TurnOnOff,pcFBS_OffMsg,pcFBS_AppID,pcFBS_RedirectURL,pcFBS_Header,pcFBS_Footer,pcFBS_PageWidth,pcFBS_CustomDisplay,pcFBS_CatImages,pcFBS_CatRow,pcFBS_CatRowsperPage,pcFBS_BType,pcFBS_PrdRow,pcFBS_PrdRowsPerPage,pcFBS_ShowSKU,pcFBS_ShowSmallImg) VALUES (" & pcStoreOn & ",'" & pcStoreMsg & "','" & pcAppID & "','" & pcHomepage & "','" & pcStoreHeader & "','" & pcStoreFooter & "'," & pcAppWidth & "," & pcCustomDisplay & "," & pcIntCatImages & "," & pcIntCatRow & "," & pcIntCatRowsPerPage & ",'" & pcStrBType & "'," & pcIntPrdRow & "," & pcIntPrdRowsPerPage & "," & pcIntShowSKU & "," & pcIntShowSmallImg & ");"
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
	msg="Updated Facebook Store Settings successfully!"
end if%>

<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer" align="center">
			<%
			if msg<>"" then %>
				<div class="pcCPmessageSuccess"><%=msg%></div>
			<% end if %>
		</td>
	</tr>
</table>
<%



pcStoreOn=0
pcStoreMsg="This store has been temporarily closed."
pcAppID=""
pcAppWidth=790
pcHomepage="viewcategories.asp"
pcStoreHeader="<div id=""pcFacebookHeader""><a href=""viewcategories.asp"">Browse The Store</a><a href=""showbestsellers.asp"">Best Sellers</a><a href=""showspecials.asp"">Specials</a><a href=""shownewarrivals.asp"">New Arrivals</a><a href=""search.asp"">Search Products</a><a href=""viewcart.asp"" style=""border-right: none;"">Shopping Cart</a><a href=""CustPref.asp"" style=""border-right: none;"">My Account</a></div>"
pcStoreFooter=""
pcCustomDisplay=0
pcIntCatImages=0
pcIntCatRow=3
pcIntCatRowsPerPage=3
pcStrBType="h"
pcIntPrdRow=3
pcIntPrdRowsPerPage=3
pcIntShowSKU="-1"
pcIntShowSmallImg="-1"



query="SELECT pcFBS_TurnOnOff,pcFBS_OffMsg,pcFBS_AppID,pcFBS_RedirectURL,pcFBS_Header,pcFBS_Footer,pcFBS_PageWidth,pcFBS_CustomDisplay,pcFBS_CatImages,pcFBS_CatRow,pcFBS_CatRowsperPage,pcFBS_BType,pcFBS_PrdRow,pcFBS_PrdRowsPerPage,pcFBS_ShowSKU,pcFBS_ShowSmallImg FROM pcFacebookSettings;"
set rs=connTemp.execute(query)
if not rs.eof then
	pcStoreOn=rs("pcFBS_TurnOnOff")
	pcStoreMsg=rs("pcFBS_OffMsg")
	if IsNull(pcStoreMsg) OR pcStoreMsg="" then
		pcStoreMsg="This store has been temporarily closed."
	end if
	pcAppID=rs("pcFBS_AppID")
	pcHomepage=rs("pcFBS_RedirectURL")
	if IsNull(pcHomepage) OR pcHomepage="" then
		pcHomepage="viewcategories.asp"
	end if
	pcStoreHeader=rs("pcFBS_Header")
	pcStoreFooter=rs("pcFBS_Footer")
	pcAppWidth=rs("pcFBS_PageWidth")
	if IsNull(pcAppWidth) OR pcAppWidth="" OR pcAppWidth="0" then
		pcAppWidth=810
	end if
	pcCustomDisplay=rs("pcFBS_CustomDisplay")
	if IsNull(pcCustomDisplay) OR pcCustomDisplay="" then
		pcCustomDisplay=0
	end if
	pcIntCatImages=rs("pcFBS_CatImages")
	if IsNull(pcIntCatImages) OR pcIntCatImages="" then
		pcIntCatImages=0
	end if
	pcIntCatRow=rs("pcFBS_CatRow")
	if IsNull(pcIntCatRow) OR pcIntCatRow="" OR pcIntCatRow="0" then
		pcIntCatRow=3
	end if
	pcIntCatRowsPerPage=rs("pcFBS_CatRowsperPage")
	if IsNull(pcIntCatRowsPerPage) OR pcIntCatRowsPerPage="" OR pcIntCatRowsPerPage="0" then
		pcIntCatRowsPerPage=3
	end if
	pcStrBType=rs("pcFBS_BType")
	if IsNull(pcStrBType) OR pcStrBType="" then
		pcStrBType="h"
	else
		pcStrBType=lcase(pcStrBType)
	end if
	pcIntPrdRow=rs("pcFBS_PrdRow")
	if IsNull(pcIntPrdRow) OR pcIntPrdRow="" OR pcIntPrdRow="0" then
		pcIntPrdRow=3
	end if
	pcIntPrdRowsPerPage=rs("pcFBS_PrdRowsPerPage")
	if IsNull(pcIntPrdRowsPerPage) OR pcIntPrdRowsPerPage="" OR pcIntPrdRowsPerPage="0" then
		pcIntPrdRowsPerPage=3
	end if
	pcIntShowSKU=rs("pcFBS_ShowSKU")
	if IsNull(pcIntShowSKU) OR pcIntShowSKU="" then
		pcIntShowSKU=-1
	end if
	pcIntShowSmallImg=rs("pcFBS_ShowSmallImg")
	if IsNull(pcIntShowSmallImg) OR pcIntShowSmallImg="" then
		pcIntShowSmallImg=-1
	end if
end if
set rs=nothing



%>


<form name="form1" method="post" action="<%=pcPageName%>" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td valign="top">
				<div id="TabbedPanels2" class="tabbable">
	


            <ul class="nav nav-tabs">
					<li class="active"><a href="#tabs-1" data-toggle="tab"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_1")%></a></li>
					<li><a href="#tabs-2" data-toggle="tab"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_4")%></a></li>
				  </ul>
                  
       
            <div class="tab-content">
		
					<div id="tabs-1" class="tab-pane active">
						<table class="pcCPcontent">
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2">Turn Facebook Store On & Off</th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="StoreOn" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_7")%>
									<input type="radio" name="StoreOn" value="0" <% if pcStoreOn="0" then%>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_8")%>
								</td>
							</tr>
							<tr>
								<td align="right" valign="top"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_9")%></td>
								<td><textarea name="StoreMsg" cols="60" rows="6"><%=pcStoreMsg%></textarea></td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2">Facebook Application Settings</th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td align="right">Application ID:</td>
								<td align="left">
								<input type="text" name="AppID" value="<%=pcAppID%>" size="45">
								</td>
							</tr>
                            <!--
							<tr>
								<td align="right">Application Width:</td>
								<td align="left">
								<select name="AppWidth">
									<option value="810" <%if pcAppWidth<>"520" then%>selected<%end if%>>Normal (810px)</option>
									<option value="520" <%if pcAppWidth="520" then%>selected<%end if%>>Narrow (520px)</option>
								</select>
								</td>
							</tr>
                            -->
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<th colspan="2">Homepage</th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td align="right">Homepage URL:</td>
								<td align="left">
								<input type="text" name="HomePage" value="<%=pcHomepage%>" size="60">
								</td>
							</tr>
                            <!--
							<tr>
								<td align="right" valign="top">Header:</td>
								<td><textarea name="StoreHeader" cols="60" rows="6"><%=pcStoreHeader%></textarea></td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td align="right" valign="top">Footer:</td>
								<td><textarea name="StoreFooter" cols="60" rows="6"><%=pcStoreFooter%></textarea></td>
							</tr>
                            -->
						</table>
					</div>

					<div id="tabs-2" class="tab-pane">
						<table class="pcCPcontent">
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2"><input type="checkbox" name="UseDefault" value="1" <%if pcCustomDisplay<>"1" then%>checked<%end if%> onclick="javascript: if (this.checked) {document.getElementById('DisplayTab').style.display='none'} else {document.getElementById('DisplayTab').style.display=''};"> Use Store's Default Display Settings</td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
							<td colspan="2">
							<table class="pcCPcontent" id="DisplayTab" <%if pcCustomDisplay<>"1" then%>style="display: none"<%end if%>>
							<tr>
								<th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_50")%></th>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_51")%>:</td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="1" checked class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_505")%>
								</td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="0" <% If trim(pcIntCatImages)="0" then  %>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_506")%>
								</td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="4" <% If trim(pcIntCatImages)="4" then  %>checked<% end if %> class="clearBorder">Thumbnails only
								</td>
							</tr>
							<tr>
								<td colspan="2">
									<input type="radio" name="catImages" value="2" <% If trim(pcIntCatImages)="2" then  %>checked<% end if %> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_507")%>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_52")%>:</td>
							</tr>
							<tr>
								<td align="right" nowrap="nowrap"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_508")%>:</td>
								<td align="left"><input type="text" name="CatRow" value="<%=pcIntCatRow%>">
								</td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_509")%>:</td>
								<td width="556" align="left">
								<input type="text" name="CatRowsperPage" value="<%=pcIntCatRowsPerPage%>">
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_53")%>:</td>
							</tr>
							<tr>
								<td colspan="2">
									<% If ucase(trim(pcStrBType))="H" then  %>
									 <input type="radio" name="BType" value="H" checked class="clearBorder">
									<% Else %>
									 <input type="radio" name="BType" value="H" class="clearBorder">
									<% End If %>
								 <%=dictLanguageCP.Item(Session("language")&"_cpCommon_510")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrBType))="P" then  %>
								 <input type="radio" name="BType" value="P" checked class="clearBorder">
								<% Else %>
								 <input type="radio" name="BType" value="P" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_511")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrBType))="L" then  %>
									<input type="radio" name="BType" value="L" checked class="clearBorder">
								<% Else %>
									<input type="radio" name="BType" value="L" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_512")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<% If ucase(trim(pcStrBType))="M" then  %>
									<input type="radio" name="BType" value="M" checked class="clearBorder">
								<% Else %>
									<input type="radio" name="BType" value="M" class="clearBorder">
								<% End If %>
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_513")%></td>
							</tr>
							<tr>
								<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_54")%></td>
							</tr>
							<tr>
								<td align="right" width="20%" nowrap="nowrap"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_514")%>:</td>
								<td align="left" width="80%" nowrap="nowrap"><input type="text" name="PrdRow" value="<%=pcIntPrdRow%>">
								</td>
							</tr>
							<tr>
								<td align="right"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_509")%>:</td>
								<td align="left">
								<input type="text" name="PrdRowsPerPage" value="<%=pcIntPrdRowsPerPage%>">
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							<tr>
								<td colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_55")%></td>
							</tr>
							<tr>
								<td colspan="2">
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_515")%>: <input type="radio" name="ShowSKU" value="-1" <%If pcIntShowSKU="-1" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="radio" name="ShowSKU" value="0" <%If pcIntShowSKU="0" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								</td>
							</tr>
							<tr>
								<td colspan="2">
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_516")%>: <input type="radio" name="ShowSmallImg" value="-1" <%If pcIntShowSmallImg="-1" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_310")%>&nbsp;<input type="radio" name="ShowSmallImg" value="0" <%If pcIntShowSmallImg="0" then%> checked<% end if%> class="clearBorder"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_311")%>
								</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
							</table>
							</td>
							</tr>
							<tr>
								<td class="pcCPspacer" colspan="2"></td>
							</tr>
						</table>
        
        </div>  
    </div>
    <div style="clear: both;">&nbsp;</div>
				<script type=text/javascript>
					$pc( "#TabbedPanels2" ).tab('show')
				</script>
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>
				<p>
				  <input type="submit" name="updateSettings" value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_107")%>" class="btn btn-primary">
				</p>
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
