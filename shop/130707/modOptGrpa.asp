<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Option Group" %>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
Dim pid, AssignID, pidProduct, pidOptionGroup
Dim	FacetCols

FacetCols=3

'START code used when loading the page after adding a new attribute
If request.form("Submit")<>"" then
	'add new attribute
	pidOptionGroup=request.form("idOptionGroup")
	AssignID=request.Form("AssignID")
	pidProduct=request.Form("idProduct")
	refpage=request.Form("refpage")
	attribute=replace(trim(request.form("attrib")),"'","''")
	if attribute="" then
		call closeDb()
		response.redirect "modOptGrpa.asp?msg="&Server.Urlencode("You need to specify an attribute to add.")&"&idOptionGroup="&pidOptionGroup
		response.end
	end if

	pcv_OptImg=request("OptImg")
	pcv_OptCode=request("OptCode")

	Dim pcv_strResults	
	pcv_strResults=""
		query="INSERT INTO options (optionDescrip,pcOpt_Img,pcOpt_Code) VALUES (N'"&attribute&"','" & pcv_OptImg & "','" & pcv_OptCode & "')"
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=connTemp.execute(query)
		query="SELECT idOption FROM options WHERE optionDescrip='"&attribute&"' ORDER BY idOption Desc;"
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=connTemp.execute(query)
		pcv_strResults="Successfully created new Attribute. "
	pidOption=rs("idOption")
	
	'CHECK IF THIS OPTION ALREADY IS ASSIGNED TO THE GROUP
	query="SELECT idOption FROM optGrps WHERE idOptionGroup="&pidOptionGroup&" AND idOption="&pidOption&";"
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=connTemp.execute(query)
	if rs.eof then
		query="INSERT INTO optGrps (idOption, idOptionGroup) VALUES ("&pidOption&", "&pidOptionGroup&")"
		Set rs=Server.CreateObject("ADODB.Recordset")
		Set rs=connTemp.execute(query)
		pcv_strResults = pcv_strResults & "Successfully added Attribute to Option Group."
		
		tFGID=getUserInput(request("tmpFGID"),0)
		if tFGID="" then
			tFGID=0
		end if
		
		if tFGID>"0" then
			FCount=getUserInput(request("FCount"),0)
			if FCount>"0" then
				For ik=1 to FCount
					tFC=getUserInput(request("FC" & ik),0)
					if tFC="" then
						tFC="0"
					end if
					if tFC>"0" then
						queryQ="INSERT INTO pcFCAttr (IdOption,pcFC_ID) VALUES (" & pidOption & "," & tFC & ");"
						set rsQ=connTemp.execute(queryQ)
						set rsQ=nothing
					end if
				Next
			end if
		end if
	else
		pcv_strResults = pcv_strResults & "Attribute already exists in Option Group."
	End if
	
	'if the admin had originally come from a product options page (modPrdOpta2.asp or modPrdOpta3.asp), go back to that page, otherwise stay on the same page	
	if pidProduct = 0 then
		set rs=nothing
		
	 	call closeDb()
response.redirect "modOptGrpa.asp?s=1&msg="&Server.Urlencode(pcv_strResults)&"&idOptionGroup="&pidOptionGroup
	else
	 if refpage = "modPrdOpta3" then
		set rs=nothing
		
	  	call closeDb()
response.redirect "modPrdOpta3.asp?AssignID="&AssignID&"&idProduct="&pidProduct&"&idOptionGroup="&pidOptionGroup
	  else
		set rs=nothing
		
	  	call closeDb()
response.redirect "modPrdOpta2.asp?AssignID="&AssignID&"&idProduct="&pidProduct&"&idOptionGroup="&pidOptionGroup
	  end if
	response.end
	end if
End if
'END code used when loading the page after adding a new attribute

'if the request is coming from modPrdOpta2.asp (specific product), get that information so that the admin can return to that page
pidOptionGroup=request.Querystring("idOptionGroup")
AssignID=request.QueryString("AssignID")
 if AssignID = "" then
 	AssignID = 0
 end if
pidProduct=request.QueryString("idProduct")
 if pidProduct = "" then
 	pidProduct = 0
 end if
refpage=request.QueryString("page")

if trim(pidOptionGroup)="" then
   call closeDb()
response.redirect "msg.asp?message=22"
end if

	
Function GetAttributes()

	'// Gets group assignments
	query="SELECT options.optionDescrip, options.idOption, options.pcOpt_Img, options.pcOpt_Code FROM options INNER JOIN optGrps ON options.idOption=optGrps.idoption WHERE (((optGrps.idOptionGroup)="&pidOptionGroup&")) ORDER BY options.optionDescrip;"
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs=connTemp.execute(query)
	
	If Err Then
		CleanUp
		GetAttributes=False
		Exit Function
	ElseIf rs.EOF OR rs.BOF Then
		CleanUp
		GetAttributes=False
		Exit Function
	Else
		GetAttributes=True
	End If
End Function
Sub CleanUp
	Set rs=Nothing
	Set connTemp=Nothing
	
End Sub


query="SELECT * FROM optionsGroups WHERE idOptionGroup=" &pidOptionGroup
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	
 	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error loading option group information on modOptGrpa.asp") 
end If

query="SELECT pcFG_ID FROM pcFGOG WHERE idOptionGroup=" & pidOptionGroup & ";"
set rs=connTemp.execute(query)
tmpFGID=0
if not rs.eof then
	tmpFGID=rs("pcFG_ID")
end if
set rs=nothing

if tmpFGID>"0" then
	tmpcolspan="6"
else
	tmpcolspan="4"
end if


%>
<!--#include file="AdminHeader.asp"-->
<script type=text/javascript>
function Form1_Validator(theForm)
{
	<%if tmpFGid>"0" then%>
	if (theForm.pcFGID.value != theForm.saveFG.value)
  	{
	    return (confirm('You are about to change the facet group of this option group. If you do this, all linked facets will be removed. Are you sure you want to complete this action?'));
	}
	<%end if%>

	return (true);
}
</script>
<form method="post" name="modOpGr" action="modOptGrpb.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
<input type="hidden" name="idOptionGroup" size="60" value="<%=pidOptionGroup%>">
<table class="pcCPcontent">
    <tr>
        <td colspan="<%=tmpcolspan%>" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
	<tr>                       
		<th colspan="<%=tmpcolspan%>">Edit Option Group</th>
	</tr>
	<tr>
		<td colspan="<%=tmpcolspan%>" class="pcCPspacer"></td>
	</tr>              
	<tr> 
		<td width="20%">Option Group:</td>
		<td width="80%"><input type="text" name="optionGroupDesc" value="<%=rstemp("optionGroupDesc")%>" size="35"></td>
	</tr>
    
	<%	
    If scSearch_IsEnabled = True Then
    	
        query="SELECT pcFG_ID,pcFG_Name FROM pcFacetGroups ORDER BY pcFG_Name ASC;"
        set rs=connTemp.execute(query)
        if not rs.eof then
            tmpArr=rs.getRows()
            set rs=nothing
            intCount=ubound(tmpArr,2)
            %>
            <tr>
            <td width="20%">Facet Group:</td>
            <td width="80%">
            <select name="pcFGID" id="pcFGID">
                <option value="0"></option>
            <%For i=0 to intCount%>
                <option value="<%=tmpArr(0,i)%>" <%if Clng(tmpArr(0,i))=Clng(tmpFGid) then%>selected<%end if%> ><%=tmpArr(1,i)%></option>
            <%Next%>
            </select>
            <input type="hidden" name="saveFG" id="saveFG" value="<%=tmpFGid%>">
            </td>
            </tr>
        <%
        end if
        set rs=nothing
    
    End If
    %>
    
	<tr> 
		<td colspan="<%=tmpcolspan%>">     
		<input type="submit" name="modify" value="Update Group" class="btn btn-primary">&nbsp;
		<input type="button" class="btn btn-default"  name="Button" value="Back" onClick="document.location.href='ManageOptions.asp';">
		</td>
	</tr>
	<tr>
		<td colspan="<%=tmpcolspan%>" class="pcCPspacer"></td>
	</tr>          
</table>
</form>

<table class="pcCPcontent">
	<tr>
		<td colspan="<%=tmpcolspan%>" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th colspan="<%=tmpcolspan%>">Manage Attributes in this Group</th>
	</tr>
	<tr>
		<td colspan="<%=tmpcolspan%>" class="pcCPspacer"></td>
	</tr>               
	                     

		<tr>
			<td colspan="<%=tmpcolspan%>">
	        		<h2>Add new attribute</h2>
			        <form name="form1" action="modOptGrpa.asp" method="post" class="pcForms"> 
               
							<table class="pcCPcontent">
				            <tr>
				            	<td nowrap>New Attribute Name:</td>
				                <td><input type="text" name="attrib"></td>
				            </tr>
							<tr>
								<td nowrap valign="top">Image File:</td>
								<td valign="top">
									<input type="text" name="OptImg" size="20"><br>
									<font color="#666666">Type in the file name, no file path. All images must be located in the 'pc/catalog' folder. This image is displayed on the product details page <u>only</u> when this attribute belongs to the first option group assigned to a product (e.g. color swatch). For more information, please see the Apparel Add-On User Guide. <a href="#" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">Upload a new image</a>.</font>
								</td>
							</tr>
							<tr>
								<td valign="top" nowrap>Attribute Code:</td>
								<td valign="top">
									<input type="text" name="OptCode" size="20"><br>
									<font color="#666666">Used when generating sub-product SKUs. For more information please see the Apparel Add-On User Guide.</font>
								</td>
							</tr>
							<%IF tmpFGID>"0" THEN
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
									<%For ik=0 to FCount%>
									<td valign="top" width="10"><input type="checkbox" name="FC<%=Clng(ik)+1%>" id="FC<%=Clng(ik)+1%>" value="<%=tmpFArr(0,ik)%>" class="clearBorder"></td>
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

							</table>

						<input type="hidden" name="idOptionGroup" value="<%=pidOptionGroup%>">
						<input type="hidden" name="tmpFGID" value="<%=tmpFGID%>">
						<input type="hidden" name="AssignID" value="<%=AssignID%>">
						<input type="hidden" name="idProduct" value="<%=pidProduct%>">
						<input type="hidden" name="refpage" value="<%=refpage%>">
						&nbsp;<input type="submit" name="Submit" value="Add" class="btn btn-primary">
					</form>

	        </td>
		</tr>
	<%
	If NOT GetAttributes() Then
		noattrb=1
		%> 
 		<tr>
			<td colspan="<%=tmpcolspan%>">
				<h2>Existing attributes:</h2>
				No attributes found
			</td>
		</tr>
	<%
	Else
	%>
		<tr>
			<td colspan="<%=tmpcolspan%>">
	        <h2>Existing attributes:</h2>
	        </td>
		</tr> 

		<% If statusAPP="1" OR scAPP=1 Then %> 
			<tr>
				<td>Attribute Name</td>
				<td>Attribute Image</td>
				<td>Attribute Code</td>
				<%IF tmpFGID>"0" THEN%>
				<td nowrap>Map Code</td>
				<td nowrap>Map Image</td>
				<%END IF%>
				<td>&nbsp;</td>		
			</tr>
		<% End If %>
		<%
		noattrb=0
		do while not rs.eof
			
				poptionID=rs("idOption")
				poptionDescrip=replace(rs("optionDescrip"),"''","'")
				poptionDescrip=replace(rs("optionDescrip"),"""","&quot;")
				pcv_OptImg=rs("pcOpt_Img")
				if pcv_OptImg="" or IsNull(pcv_OptImg) then
					pcv_OptImg="N/A"
				end if
				pcv_OptCode=rs("pcOpt_Code")
				if pcv_OptCode="" or IsNull(pcv_OptCode) then
					pcv_OptCode="N/A"
				end if
				%>
	
				<% If statusAPP="1" OR scAPP=1 Then %>                      
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist" valign="top"> 
		        		<td width="30%"><%=poptionDescrip%></td>
						<td nowrap="nowrap"><a href="../pc/catalog/<%=pcv_OptImg%>" target="_blank"><%=pcv_OptImg%></a></td>
						<td><%=pcv_OptCode%></td>
						<%IF tmpFGID>"0" THEN
						tmpIDOpt=rs("idOption")
						queryQ="SELECT pcFC_Code,pcFC_Img FROM pcFacets INNER JOIN pcFCAttr ON pcFacets.pcFC_ID=pcFCAttr.pcFC_ID WHERE pcFCAttr.idOption=" & tmpIDOpt & ";"
						set rsQ=connTemp.execute(queryQ)
						tmpFCode=""
						tmpFImg=""
						if not rsQ.eof then
							pcFArr=rsQ.getRows()
							set rsQ=nothing
							FCount=ubound(pcFArr,2)
							For ik=0 to FCount
								if tmpFCode<>"" then
									tmpFCode=tmpFCode & "<br>"
								end if
								if pcFArr(0,ik)<>"" then
								tmpFCode=tmpFCode & pcFArr(0,ik)
								end if
								if tmpFImg<>"" then
									tmpFImg=tmpFImg & "<br>"
								end if
								if pcFArr(1,ik)<>"" then
									tmpFImg=tmpFImg & "<img src=""../pc/catalog/" & pcFArr(1,ik) & """  border=0 align=""top"">"
								end if
							Next
						end if
						set rsQ=nothing
						%>
						<td><%=tmpFCode%></td>
						<td><%=tmpFImg%></td>
						<%END IF%>
		        		<td nowrap align="right" class="cpLinksList"> 
							<a href="modOpta.asp?idOption=<%=rs("idOption")%>&idOptionGroup=<%=pidOptionGroup%>">View/Edit</a>
							&nbsp;|&nbsp; 
							<a href="javascript:if (confirm('This attribute may have been assigned to one or more products: are you sure you want to delete it?')) location='actionOptions.asp?delete=<%=rs("idOption")%>&idOptionGroup=<%=pidOptionGroup%>'">Delete</a>
		        		</td>
					</tr>

				<% Else %>

					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				        <td width="50%"><%= rs("optionDescrip") %></td>
				        <td width="50%" nowrap align="right" class="cpLinksList"> 
				            <a href="modOpta.asp?idOption=<%=rs("idOption")%>&idOptionGroup=<%=pidOptionGroup%>">Rename</a> | <a href="javascript:if (confirm('This attribute may have been assigned to one or more products: are you sure you want to delete it?')) location='actionOptions.asp?delete=<%=rs("idOption")%>&idOptionGroup=<%=pidOptionGroup%>'">Delete</a>
				        </td>
					</tr>

				<% End If %> 
	                      
		<%
	    rs.MoveNext
	    Loop
	    CleanUp
    End If
    %>
	<tr>
		<td colspan="<%=tmpcolspan%>" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="<%=tmpcolspan%>" align="center">
		<form class="pcForms">
		<input type="button" class="btn btn-default"  value="Manage Options" onClick="location.href='ManageOptions.asp'">
		<% If noattrb=0 Then %>
        	&nbsp;<input type="button" class="btn btn-default"  value="Add to Multiple Products" onClick="location.href='AssignMultiOptions.asp?idOptionGroup=<%=pidOptionGroup%>'">
        	&nbsp;<input type="button" class="btn btn-default"  value="Remove from Multiple Products" onClick="location.href='RevMultiOptions.asp?idOptionGroup=<%=pidOptionGroup%>'">
		<% end if %>
		</form>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->