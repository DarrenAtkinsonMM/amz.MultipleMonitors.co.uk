<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
tmpFGID=getUserInput(request("id"),0)
if tmpFGID="" then
	tmpFGID=0
end if
if tmpFGID="0" then
	response.redirect "ManageFacetGroups.asp"
end if

if request("action")="upd" then
	FCount=getUserInput(request("FCount"),0)
	if FCount="" then
		FCount=0
	end if
	For i=1 to FCount
		tmpFCID=getUserInput(request("ID" & i),0)
		tmpFCOrder=getUserInput(request("O" & i),0)
		
		if (tmpFCID<>"") AND (tmpFCOrder<>"") then
			query="UPDATE pcFacets SET pcFC_Order=" & tmpFCOrder & " WHERE pcFC_ID=" & tmpFCID & " AND pcFG_ID=" & tmpFGid & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
	Next
	
	msg="Facets order has beeen updated successfully!"
	msgType=1
end if		
	
tmpFGName=""

query="SELECT pcFG_Name FROM pcFacetGroups WHERE pcFG_ID=" & tmpFGID & ";"
set rs=connTemp.execute(query)

if not rs.eof then
	tmpFGName=rs("pcFG_Name")
end if
set rs=nothing

pageTitle="Manage Facets"
if tmpFGName<>"" then
	pageTitle=pageTitle & " of the group: " & tmpFGName
end if %>
<!--#include file="AdminHeader.asp"-->
<form method="post" action="ManageFacets.asp?action=upd" name="mngFC" class="pcForms">
<input type="hidden" name="id" value="<%=tmpFGID%>">
<table class="pcCPcontent">
    <tr>
    	<td colspan="6">
        	<!--#include file="pcv4_showMessage.asp"-->
        	<div class="cpOtherLinks"><a href="AddEditFC.asp?idgroup=<%=tmpFGID%>">Add New Facet</a> | <a href="AddEditFG.asp">Add New Facet Group</a></div>
        </td>
	</tr>
                    
<%
	query="SELECT pcFC_ID,pcFC_Code,pcFC_Name,pcFC_Img,pcFC_Order FROM pcFacets WHERE pcFG_ID=" & tmpFGID & " ORDER BY pcFC_Order ASC"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	intCount=-1
	if rs.EOF then
		set rs=nothing
		
%>      
      <tr> 
        <td colspan="6"><div class="pcCPmessage">No facets found</div></td>
      </tr>
      <tr>
        <td colspan="6" class="pcCPspacer"></td>
      </tr>                
<% 

	Else 
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)%>
		<tr>
			<th width="10&" align="right">ID</td>
			<th width="20%">Code</td>
			<th width="20%">Name</td>
			<th width="30%">Image</td>
			<th width="10%">Order</td>
			<th width="10%">&nbsp;<input type="hidden" name="FCount" value="<%=Clng(intCount)+1%>"></td>
		</tr>
		<%For i=0 to intCount
		tmpFCID=tmpArr(0,i)
		tmpFCCode=tmpArr(1,i)
		tmpFCName=tmpArr(2,i)
		tmpFCImg=tmpArr(3,i)
		tmpFCOrder=tmpArr(4,i)%>
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist">
			<td width="10&"><%=tmpFCID%></td>
			<td width="20%"><a href="AddEditFC.asp?idgroup=<%=tmpFGID%>&id=<%=tmpFCID%>"><%=tmpFCCode%></a></td>
			<td width="20%"><a href="AddEditFC.asp?idgroup=<%=tmpFGID%>&id=<%=tmpFCID%>"><%=tmpFCName%></a></td>
			<td width="30%"><%=tmpFCImg%>&nbsp;<%if tmpFCImg<>"" then%><img src="../pc/catalog/<%=tmpFCImg%>"  border=0 align=absbottom><%end if%></td>
			<td width="10%">
				<input type="text" name="O<%=i+1%>" id="O<%=i+1%>" size="4" value="<%=tmpFCOrder%>">
				<input type="hidden" name="ID<%=i+1%>" id="ID<%=i+1%>" value="<%=tmpFCID%>">
			</td>
			<td width="10%" nowrap class="cpLinksList">
				<%
				Mapped=0
				queryQ="SELECT TOP 1 pcFC_ID FROM pcFCAttr WHERE pcFC_ID=" & tmpFCID & ";"
				set rsQ=connTemp.execute(queryQ)
				if not rsQ.eof then
					Mapped=1
				else
					Mapped=0
				end if
				set rsQ=nothing
				%>
				<a href="AddEditFC.asp?idgroup=<%=tmpFGID%>&id=<%=tmpFCID%>">Edit</a> | <a href="javascript:<%if Mapped=1 then%>if (confirm('This Facet was linked to a Product Attribute. Are you sure you want to complete this action?')) location='delFC.asp?id=<%=tmpFCID%>&idgroup=<%=tmpFGID%>';<%else%>if (confirm('You are about to remove this facet from the facet group. Are you sure you want to complete this action?')) location='delFC.asp?id=<%=tmpFCID%>&idgroup=<%=tmpFGID%>';<%end if%>">Delete Facet</a>
			</td>
		</tr>
		<%Next%>
		<tr>
		
	<%End If
	set rs=nothing
%>
<tr>
	<td colspan="6" class="pcCPspacer">&nbsp;</td>
</tr>
	<td colspan="6">
	<%if intCount>=0 then%>
	<input type="submit" name="Submit" value="Update Facets order" class="btn btn-primary">&nbsp;
	<%end if%>
	<input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:location='ManageFacetGroups.asp';">
	</td>
</tr>     
</table>
</form>
<br /><br />
<!--#include file="AdminFooter.asp"-->