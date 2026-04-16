<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2014. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Remove Option Attributes from Multiple Products" %>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<% 
dim f, pIdProduct

If request.Form("Submit2")<>"" then
	
	pIdProductArray=request.form("idProduct")
	pIdOptionGroup=request.form("idOptionGroup")
	pIdProduct=split(pIdProductArray,",")
	cnt=0
	
	'//////////////////////////////////////////////////
	'// START LOOP THROUGH ALL THE PRODUCT ID(S)
	'//////////////////////////////////////////////////
	For i=lBound(pIdProduct) to UBound(pIdProduct)
		if trim(pIdProduct(i)<>"") then
			contgo=0
			
			'// START: Loop Through all the selected Options Id(s)
			For Each intOptionID in Request.Form("idoption")
				
				'// Check if the Option has been deleted already
				strSQL="SELECT * FROM options_optionsGroups WHERE idproduct="&pIdProduct(i)&" AND idoptionGroup="&pIdOptionGroup&" AND idOption="&intOptionID&";"
				set rstemp=conntemp.execute(strSQL)
				if not rstemp.eof then
					'// It is NOT deleted
					cnt=1
				end if
				
				'// It is NOT deleted, Generate the Delete Statement
				if cnt=1 then					
				
					If statusAPP="1" OR scAPP=1 Then
						query="SELECT idoptoptgrp FROM options_optionsGroups WHERE idproduct="&pIdProduct(i)&" AND idoptionGroup="&pIdOptionGroup&" AND idOption="&intOptionID&";"
						set rstemp=connTemp.execute(query)
						do while not rstemp.eof
							idoptoptgrp=rstemp("idoptoptgrp")
							
							query="UPDATE Products SET removed=-1,active=0 WHERE pcprod_ParentPrd="& pIdProduct(i) & " AND removed=0 AND ((pcprod_Relationship like '%[_]" & idoptoptgrp & "[_]%') OR (pcprod_Relationship like '%[_]" & idoptoptgrp & "'))"
							set rstemp1=conntemp.execute(query)
							set rstemp1=nothing
	
							rstemp.MoveNext
						loop
						set rstemp=nothing
					End If
				
					strSQL="DELETE FROM options_optionsGroups WHERE idproduct="&pIdProduct(i)&" AND idoptionGroup="&pIdOptionGroup&" AND idOption="&intOptionID&";"
					set rstemp=conntemp.execute(strSQL)
				end if	
			Next
			'// END: Loop Through all the selected Options Id(s)
			
			'// Check if all options have been removed.
			strSQL="SELECT * FROM options_optionsGroups WHERE idproduct="&pIdProduct(i)&" AND idoptionGroup="&pIdOptionGroup&";"
			set rstemp=conntemp.execute(strSQL)							
			if rstemp.eof then
				'// It is NOT related
				contgo=1
			end if	
			
			'// It all Options have been removed then delete the corrisponding record in pcProductOptions
			if contgo=1 then				
				strSQL="DELETE FROM pcProductsOptions WHERE idproduct="&pIdProduct(i)&" AND idoptionGroup="&pIdOptionGroup&";"
				set rstemp=conntemp.execute(strSQL)
			end if			
		
		
		end if		
	Next
	'//////////////////////////////////////////////////
	'// END LOOP
	'//////////////////////////////////////////////////
	
	set rstemp=nothing
	
	
	if cnt=0 then	
		call closeDb()
response.redirect "manageOptions.asp?msg="&server.urlencode("No products needed to be updated.")
	else
		call closeDb()
response.redirect "manageOptions.asp?s=1&msg="&server.urlencode("Successfully removed Option Attributes from the selected products.")
	end if	
	response.end
End If

' form parameter
pIdOptionsGroups=request("idOptionGroup")
pIdProduct=request("prdlist")

if trim(pidproduct)="" then
   call closeDb()
response.redirect "msg.asp?message=2"
end if
%>
<form method="post" name="modifyProduct" action="RevMultiOptions1.asp" class="pcForms">
<table class="pcCPcontent">
<tr> 
	<td colspan="5">
		<input type="hidden" name="idproduct" value="<%response.write pIdProduct%>">
		<table width="100%">
		<tr>
			<td colspan="2"> 
				<input type="hidden" name="idOptionGroup" value="<%=pIdOptionsGroups%>">
				<% query="SELECT * FROM optionsGroups WHERE idoptionGroup=" &pIdOptionsGroups
				set rstemp=conntemp.execute(query) %>
				Removing option group: <b><%=rstemp("optionGroupDesc")%></b>. Select the attributes that you would like to remove.
			</td>
		</tr>
		</table>
	</td>
</tr>                  
<tr>          
	<th colspan="2">Attributes</th>
</tr>
<tr>
	<td colspan="5" class="pcSpacer"></td>
</tr>
<% query="SELECT options.optionDescrip, options.idoption FROM options, optGrps WHERE options.idoption=optGrps.idoption AND  optGrps.idoptionGroup="& rstemp("idoptionGroup") &" ORDER BY optionDescrip"
set rstemp=conntemp.execute(query)
noAttribute="0"
If rstemp.eof then 
	noAttribute="1"%>
<tr> 
	<td colspan="2"><div class="pcCPmessage">There are currently no attributes assigned to this Option Group.</div></td>
</tr>
<% else
	Do until rstemp.eof %>
<tr> 
	<td width="1%"> 
		<input type="checkbox" name="idoption" value="<%=rstemp("idoption")%>">
	</td>
	<td width="99%"><%=rstemp("optionDescrip")%></td>
</tr>
                      
<% rstemp.movenext
loop 
end if%>
<tr>
	<td colspan="5" class="pcSpacer"></td>
</tr>
<tr> 
<td colspan="5" align="center"> 
	<input type="submit" name="Submit2" value="Continue" class="btn btn-primary">
</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->