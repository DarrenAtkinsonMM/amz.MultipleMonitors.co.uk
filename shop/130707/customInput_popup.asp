<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
dim f
%>
<html>
<head>
<title>Edit Order Custom Input Fields</title>
<link href="css/pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="background-image: none;">
<div id="pcCPmain" style="width:450px; background-image: none;">
<% 
if request("action")="update" then
	pidProductOrdered=request("idProductOrdered")
	intCount=request.form("xCnt")
	tempArray=""
	xCnt=0
	For i=0 to intCount
		pxfield=getUserInput(request.form("x" & i),500)
		'// Repair HTML for the Line Return
		pxfield=replace(pxfield,"&lt;BR&gt;","<BR>")
	
		if pxfield<>"" then
			pxfield=replace(pxfield,vbCrlf,"<BR>")
			pXfieldDescrip=request.form("xDesc" & i)
			if xCnt=1 then
				tempArray=tempArray& "|"
			end if
			tempArray=tempArray&pXfieldDescrip&": "&pxfield
			xCnt=1
		end if
	Next
	
	if tempArray<>"" then
		tempArray=replace(tempArray,"'","''")
	end if
		
	query="UPDATE productsOrdered SET xfdetails=N'"&tempArray&"' WHERE idProductOrdered="&pidProductOrdered&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
%>
	<table class="pcCPcontent">
		<tr>
			<th>Edit Custom Input Field</th>
		</tr>
		<tr>
			<td><br>
				<br>
				<b>This order has been updated!<br>
				<br>
			</b><br>
				</td>
		</tr>
		<tr>
			<td>
				<p align="center"><input type="button" class="btn btn-default"  name="Back" value="Close Window" onClick="opener.location.reload(); self.close();"></td>
		</tr>
	</table>
<%ELSE%>
<form name="form1" method="post" action="customInput_popup.asp?action=update" class="pcForms">
<table class="pcCPcontent">
  <tr>
    <th>Edit Custom Input Field</th>
  </tr>
  <tr>
    <td>Edit the current field and click on the &quot;update&quot; button.</td>
  </tr>
  <tr>
    <td><table width="98%" border="0" align="center" cellpadding="4" cellspacing="0">
      <% c=request("c")
			pidProductOrdered=request("idProductOrdered")
			query="select xfdetails From ProductsOrdered WHERE idProductOrdered="&pidProductOrdered&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			pxfDetails=rs("xfdetails") 
			
			if pxfDetails<>"" then 
				xfArray=split(pxfdetails,"|")
				for xf=0 to ubound(xfArray)
					tempXf=xfArray(xf)
					xSplitArray=split(tempXf,": ") 
					if int(c)=int(xf) then
						%>
						<tr valign="top">
						  <td><%=xSplitArray(0)%>:</td>
	  </tr>
						<tr valign="top">
							<td><input type="hidden" name="xdesc<%=c%>" value="<%=xSplitArray(0)%>"><textarea name="x<%=c%>" cols="30" rows="5"><%=replace(xSplitArray(1),"<BR>",vbCrlf)%></textarea>
							</td>
						</tr>
					<% else %>
						<input type="hidden" name="xdesc<%=xf%>" value="<%=xSplitArray(0)%>">
						<input type="hidden" name="x<%=xf%>" value="<%=xSplitArray(1)%>">
					<% end if 
				next %>
				<input type="hidden" name="xCnt" value="<%=ubound(xfArray)%>">
			<% end if %>
			<input type="hidden" name="idProductOrdered" value="<%=pidProductOrdered%>">
      <tr>
        <td><input type="submit" name="Submit" value="Update">
        <input type="button" class="btn btn-default"  name="Back" value="Close" onClick="self.close();">
        </td>
        </tr>
      <tr>
        <td align="center">&nbsp;</td>
      </tr>
    </table>      
    </td>
  </tr>
</table>
</form>
<%END IF%>
</div>
</body>
</html>
