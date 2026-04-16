<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="View/Edit Option Attribute" %>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" -->
<%
Dim FacetCols
FacetCols=3
pidOption=request.Querystring("idOption")
pidOptionGroup=request.Querystring("idOptionGroup")

if not validNum(pidOption) then
   call closeDb()
response.redirect "msg.asp?message=21"
end if

query="SELECT pcFG_ID FROM pcFGOG WHERE idOptionGroup=" & pidOptionGroup & ";"
set rs=connTemp.execute(query)
tmpFGID=0
if not rs.eof then
	tmpFGID=rs("pcFG_ID")
end if
set rs=nothing


query="SELECT * FROM options WHERE idOption=" &pidOption
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	
	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error renaming option attribute on modOpta.asp") 
end If

poptionDescrip=replace(rstemp("optionDescrip"),"''","'")
poptionDescrip=replace(rstemp("optionDescrip"),"""","&quot;")

pcv_OptImg=rstemp("pcOpt_Img")
pcv_OptCode=rstemp("pcOpt_Code")

set rstemp=nothing

%>
<!--#include file="AdminHeader.asp"-->
<form action="modOptb.asp" method="post" name="modOpGr" class="pcForms">
<input type="hidden" name="idOption" size="60" value="<%=pidOption%>">
<input type="hidden" name="idOptionGroup" value="<%=pidOptionGroup%>">
<% if request.querystring("redirectURL")<>"" then %>
<input type="hidden" name="redirectURL" value="<%=request.querystring("redirectURL")%>">
<input type="hidden" name="mode" value="<%=request.querystring("mode")%>">
<% end if %>
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>             
	<tr> 
		<td colspan="2">
		Attribute: <input name="optionDescrip" type="text" value="<%=poptionDescrip%>" size="40" maxlength="250">
		</td>
	</tr>

	<% If statusAPP="1" OR scAPP=1 Then %>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td valign="top">Image File:</td>
			<td>
			<script language="JavaScript">
			<!--
					function chgWin(file,window) {
					msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
					if (msgWindow.opener == null) msgWindow.opener = self;
			}
			//-->
			</script> 
				<input type="text" name="OptImg" size="20" value="<%=pcv_OptImg%>"><br>
				<font color="#666666">Type in the file name, no file path. All images must be located in the 'pc/catalog' folder. This image is displayed on the product details page <u>only</u> when this attribute belongs to the first option group assigned to a product (e.g. color swatch). For more information, please see the Apparel Add-On User Guide. <a href="#" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">Upload a new image</a>
			| <a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=OptImg&fid=modOpGr','window2')">Find an image</a></font> </td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr> 
		<tr>
			<td nowrap valign="top">Attribute Code:</td>
			<td>
				<input type="text" name="OptCode" size="20" value="<%=pcv_OptCode%>"><br>
				<font color="#666666">Used for generating sub-product SKUs. For more information, please see the Apparel Add-On User Guide. </font>
			</td>
		</tr>
	<% End If %>
	<%IF tmpFGID>"0" THEN
		queryQ="SELECT pcFC_ID FROM pcFCAttr WHERE idOption=" & pidOption & ";"
		set rsQ=connTemp.execute(queryQ)
		MCount=-1
		if not rsQ.eof then
			tmpMArr=rsQ.getRows()
			set rsQ=nothing
			MCount=ubound(tmpMArr,2)
		end if
		set rsQ=nothing
		queryQ="SELECT pcFC_ID,pcFC_Code,pcFC_Img FROM pcFacets WHERE pcFG_ID=" & tmpFGID & ";"
		set rsQ=connTemp.execute(queryQ)
		if not rsQ.eof then
			tmpFArr=rsQ.getRows()
			set rsQ=nothing
			FCount=ubound(tmpFArr,2)%>
			<tr>
			<td valign="top" nowrap>Map Option:</td>
			<td valign="top">
			<input type="hidden" name="FCount" id="FCount" value="<%=Clng(FCount)+1%>">
			<table class="pcCPcontent">
			<tr>
			<%For ik=0 to FCount
				tmpS=""
				For m=0 to MCount
					if Clng(tmpMArr(0,m))=Clng(tmpFArr(0,ik)) then
						tmpS="checked"
						Exit For
					end if
				Next
				%>
			<td valign="top" width="10"><input type="checkbox" name="FC<%=Clng(ik)+1%>" id="FC<%=Clng(ik)+1%>" value="<%=tmpFArr(0,ik)%>" class="clearBorder" <%=tmpS%>></td>
			<td valign="top" width="100" nowrap>
				<%if tmpFArr(2,ik)<>"" then%><img src="../pc/catalog/<%=tmpFArr(2,ik)%>"  border=0 align="top"><br><%end if%>
				<%if tmpFArr(1,ik)<>"" then%><%=tmpFArr(1,ik)%><%end if%>
			</td>
			<%if ((ik+1) mod FacetCols=0) then
			if ik<FCount then
				response.write "</tr><tr>"
			else
				response.write "</tr>"
			end if
			end if
			Next
			if (FCount+1) mod FacetCols<>0 then%>
			</tr>
			<%end if%>
			</table>
			</td>
			</tr>
		<%end if
		set rsQ=nothing
	END IF%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>   
	<tr>
		<td colspan="2">
		<input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:history.back()">
		&nbsp;<input type="submit" name="modify" value="Update" class="btn btn-primary">
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr> 
</table>
</form>
<!--#include file="AdminFooter.asp"-->