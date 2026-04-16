<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Product Options" %>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
    <tr>
    	<td colspan="2">
        	<!--#include file="pcv4_showMessage.asp"-->
        	<div class="cpOtherLinks"><a href="instOptGrpa.asp">Add New Option Group</a> | <a href="ApplyOptionsMulti1.asp">Copy Options from one Product to N other Products</a></div>
        </td>
	</tr>
                    
<%
	Dim pid
	

	' gets group assignments
	query="SELECT * FROM OptionsGroups WHERE idOptionGroup>1 ORDER BY OptionGroupDesc ASC"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	
	if rs.EOF then
		set rs=nothing
		
%>      
      <tr> 
        <td colspan="2"><div class="pcCPmessage">No option groups found</div></td>
      </tr>
      <tr>
        <td colspan="2" class="pcCPspacer"></td>
      </tr>                
<% 

	Else 
		Do While NOT rs.EOF %>         
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
			<td width="60%"><a href="modOptGrpa.asp?idOptionGroup=<%=rs("idOptionGroup")%>"><%=rs("OptionGroupDesc")%></a></td>
			<td width="40%" nowrap class="cpLinksList">
				<%
				If statusAPP="1" OR scAPP=1 Then

				pcv_myTest=0

					query="SELECT idproduct FROM Products WHERE removed=0 AND pcProd_ParentPrd IN (SELECT DISTINCT products.idproduct FROM products,options_optionsGroups where products.pcProd_Apparel=1 and products.removed=0 and options_optionsGroups.idproduct=products.idproduct and options_optionsGroups.idOptionGroup=" & rs("idOptionGroup") & ");"
					set rsT=connTemp.execute(query)
					
					if not rsT.eof then
						pcv_myTest=1
					end if
					set rsT=nothing

				End If
				%>
				<a href="modOptGrpa.asp?idOptionGroup=<%=rs("idOptionGroup")%>">Manage Attributes</a> | <a href="javascript:<%if pcv_myTest=1 then%>alert('You cannot remove an option group from an apparel product as it would cause its sub-products to malfunction. Do the following: (a) remove the sub-products; (b) remove the option group from the apparel product; (c) add a new option group; (d) recreate the sub-products.');<%else%>if (confirm('You are about to remove this option group from your database. Are you sure you want to complete this action?')) location='delOptGrpb.asp?idOptionGroup=<%= rs("idOptionGroup") %>';<%end if%>">Delete Group</a> | Products: <a href="AssignMultiOptions.asp?idOptionGroup=<%=rs("idOptionGroup")%>">Assign To</a>
				<%
					query="SELECT pcProductsOptions.idProduct FROM pcProductsOptions INNER JOIN products ON pcProductsOptions.idProduct = products.idProduct WHERE pcProductsOptions.idOptionGroup = " & rs("idOptionGroup") & " AND products.removed = 0"
					set rstemp=Server.CreateObject("ADODB.Recordset")
					set rstemp=connTemp.execute(query)
					if not rstemp.eof then
				%>
                    &nbsp;:&nbsp;<a href="ManageOptionsProducts.asp?idOptionGroup=<%=rs("idOptionGroup")%>">Used By</a>
                    &nbsp;:&nbsp;<a href="RevMultiOptions.asp?idOptionGroup=<%=rs("idOptionGroup")%>">Remove From</a>
				<%
					end if
					set rstemp=nothing	
				%>
			</td>
		</tr>
<% 
		rs.MoveNext
    Loop
		set rs=nothing
		
    End If
%>     
</table>
<br /><br />
<!--#include file="AdminFooter.asp"-->