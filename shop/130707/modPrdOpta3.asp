<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_UpdateDates.asp" -->
<% 
dim f, pIdProduct

If request.Form("Submit2")<>"" then
	pIdProduct=request.form("idProduct")
	pIdOptionGroup=request.form("idOptionGroup")
	pAssignID=request.form("AssignID")

	tmpstart=0	
	query="SELECT idOption,sortOrder FROM options_optionsGroups WHERE idproduct=" & pIdProduct & " AND idoptionGroup=" & pIdOptionGroup & " order by sortOrder desc;"
	set rstemp=connTemp.execute(query)
	
	if not rstemp.eof then
		tmpstart=rstemp("sortOrder")
		if IsNull(tmpstart) or tmpstart="" then
			tmpstart=0
		end if
		if tmpstart=0 then
			do while not rstemp.eof
				tmpstart=tmpstart+1
				rstemp.MoveNext
			loop
		end if
	end if
	set rstemp=nothing

	For Each intOptionID in Request.Form("idoption")
		pPrice=request.form("price"&intOptionID)
		If pPrice="" then
			pPrice="0"
		End If
		pWPrice=request.form("Wprice"&intOptionID)
		If pWPrice="" then
			pWPrice="0"
		End If
		if scDecSign="," then
			pPrice=replacecomma(pPrice)
			pWPrice=replacecomma(pWPrice)
		end if
	
		tmpstart=tmpstart+1
	
		strSQL="INSERT INTO options_optionsGroups (idproduct, idoptionGroup, idOption, Price, Wprice, sortOrder) VALUES (" & pIdProduct &", " & pIdOptionGroup & ", " & intOptionID & ","& pPrice &","& pWPrice &"," & tmpstart & ")"
		set rstemp=conntemp.execute(strSQL)
	Next
	if request.form("NewOptions")<>"" then
		ArrNewOption=Split(request("NewOptions"), ",")
	
		tmpstart=0	
		query="SELECT idOption,sortOrder FROM options_optionsGroups WHERE idproduct=" & pIdProduct & " AND idoptionGroup=" & pIdOptionGroup & " order by sortOrder desc;"
		set rstemp=connTemp.execute(query)
		
		if not rstemp.eof then
			tmpstart=rstemp("sortOrder")
			if IsNull(tmpstart) or tmpstart="" then
				tmpstart=0
			end if
			if tmpstart=0 then
				do while not rstemp.eof
					tmpstart=tmpstart+1
					rstemp.MoveNext
				loop
			end if
		end if
		set rstemp=nothing
	
		for i=lbound(ArrNewOption) to UBound(ArrNewOption)
			NewOption=ArrNewOption(i)
			'check if option description already exists
			strSQL="SELECT optionDescrip FROM options WHERE optionDescrip='"&trim(NewOption)&"';"
			set rstemp=conntemp.execute(strSQL)
			if rstemp.eof then
				strSQL="INSERT INTO options (optionDescrip) VALUES (N'"& trim(NewOption) &"')"
				set rstemp=conntemp.execute(strSQL)
				strSQL="SELECT * FROM options WHERE optionDescrip='"& trim(NewOption) &"'"
				set rstemp=conntemp.execute(strSQL)
				pIdOption=rstemp("idoption")
				strSQL="INSERT INTO optGrps (idoptionGroup, idoption) VALUES ("& pIdOptionGroup &","& pIdOption &")"
				set rstemp=conntemp.execute(strSQL)
			
				tmpstart=tmpstart+1
			
				strSQL="INSERT INTO options_optionsGroups (idproduct, idoptionGroup, idOption, Price, Wprice, sortOrder) VALUES (" & pIdProduct &", " & pIdOptionGroup & ", " & pIdOption & ",0,0," & tmpstart & ")"
				set rstemp=conntemp.execute(strSQL)
			end if
		next
	End If
	
	call updPrdEditedDate(pIdProduct)
	
	set rstemp=nothing
	
	call closeDb()
response.redirect "modPrdOpta.asp?s=1&msg="&Server.Urlencode("You have successfully updated your product attributes.")&"&idproduct="& pIdProduct
End If

' form parameter
pIdOptionsGroups=request.Querystring("idOptionGroup") 
pIdProduct=request.Querystring("idProduct")
pAssignID=request.QueryString("AssignID")
if trim(pidproduct)="" then
   call closeDb()
response.redirect "msg.asp?message=2"
end if

' get item details from db

query="SELECT idProduct, description FROM products WHERE products.idProduct=" &pIdProduct
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	
    call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in modPrdOpta3.asp: "&Err.Description) 
end if

' charge rscordset data into local variables

pIdProduct=rstemp("idProduct")
pDescription=rstemp("description")

pageTitle="Modify Product Options for <strong>" & pDescription & "</strong>"
%>
<!--#include file="AdminHeader.asp"-->

<form method="post" name="modifyProduct" action="modPrdOpta3.asp" class="pcForms">
<input type="hidden" name="AssignID" value="<%=pAssignID%>">
<input type="hidden" name="idOptionGroup" value="<%=pIdOptionsGroups%>">
<input type="hidden" name="idproduct" value="<%response.write pIdProduct%>">
<table class="pcCPcontent">
	<tr>                     
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
	<% 
	query="SELECT * FROM optionsGroups WHERE idoptionGroup=" &pIdOptionsGroups
	set rstemp=conntemp.execute(query) 
	%>						  
	<tr>                        
		<th colspan="2" nowrap="nowrap"><%=rstemp("optionGroupDesc")%></th>
		<th nowrap="nowrap">Additional: Price</th>
		<th nowrap="nowrap">Wholesale Price</th>
	</tr>
	<tr>                     
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
						  
<%
query="SELECT options.optionDescrip, options.idoption FROM options, optGrps WHERE options.idoption=optGrps.idoption AND  optGrps.idoptionGroup="& rstemp("idoptionGroup") &" AND options.idoption NOT IN (SELECT idoption FROM options_optionsGroups WHERE idproduct="&pIdProduct&" AND idoptionGroup="& rstemp("idoptionGroup") &") ORDER BY optionDescrip"
set rstemp=conntemp.execute(query)
	noAttribute="0"
	
If rstemp.eof then 
	noAttribute="1"
%>
												
	<tr> 
		<td colspan="4" align="center">
			<div class="pcCPmessage" style="text-align:left">
				All the attributes in this Option Group have already been assigned to the selected product.
					<ul>
						<li>Select &quot;Add Attributes&quot; to add new attributes to this Option Group.</li>
						<li>Select &quot;Back&quot; to return to the options summary page.</li>
					</ul>
			</div>
		</td>
	</tr>
											
<% 
else
	Do until rstemp.eof 
	%>
						  
    <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
		<td width="5%" align="center"><input type="checkbox" name="idoption" value="<%=rstemp("idoption")%>" class="clearBorder"></td>
		<td width="80%"><%=rstemp("optionDescrip")%></td>
		<td><%=scCurSign%> <input type="text" name="price<%=rstemp("idoption")%>" size="6" maxlength="10"></td>
		<td><%=scCurSign%> <input type="text" name="Wprice<%=rstemp("idoption")%>" size="6" maxlength="10"></td>
	</tr>
						  
	<% 
	rstemp.movenext
	loop
end if 
set rstemp=nothing

%>

	<tr>                     
		<td colspan="4" class="pcCPspacer"></td>
	</tr>

<% if noAttribute="1" then %>
                      
    <tr>
        <td colspan="4" align="center"> 
        <input type="button" class="btn btn-default"  name="button" value="Add Attributes" class="btn btn-primary" onClick="location.href='modOptGrpa.asp?AssignID=<%=pAssignID%>&idProduct=<%=pIdProduct%>&idOptionGroup=<%=pIdOptionsGroups%>&page=modPrdOpta3'">&nbsp;
        <input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:history.back()">
        </td>
    </tr>

<% else %>

    <tr> 
        <td colspan="4" align="center"> 
            <input type="submit" name="Submit2" value="Continue" class="btn btn-primary">&nbsp;
            <input type="button" class="btn btn-default"  name="button" value="Add More Attributes" onClick="location.href='modOptGrpa.asp?AssignID=<%=pAssignID%>&idProduct=<%=pIdProduct%>&idOptionGroup=<%=pIdOptionsGroups%>&page=modPrdOpta3'">
            <input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:history.back()">
        </td>
    </tr>

<% end if %>
                    
</table>
</form>
<!--#include file="AdminFooter.asp"-->