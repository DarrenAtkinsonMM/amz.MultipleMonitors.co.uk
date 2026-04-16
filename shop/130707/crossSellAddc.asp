<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Cross Selling - Add new relationship: set product order" %>
<% Section="products" %>
<%PmAdmin="2*3*"%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
idmain=request.QueryString("idmain") 
if request.Form("SubmitOrder")<>"" then
	idmain=request.Form("idmain") 
	icnt=request.Form("icnt")
	
	set rs=Server.CreateObject("ADODB.RecordSet")
	For i=0 to (cint(icnt)-1)
		idcrosssell=request.Form("idcrosssell"&i)
		order=request.Form("num"&i)
		query="UPDATE cs_relationships SET num="&order&" WHERE idcrosssell="&idcrosssell&";"
		set rs=conntemp.execute(query)
		set rs=nothing
	Next

	call closeDb()
    response.redirect "crossSellEdit.asp?idmain="&idmain
	response.end
end if
%>
	<form name="form1" method="post" action="crossSellAddc.asp" class="pcForms">
		<input name="idmain" type="hidden" value="<%=idmain%>">
		<% 
		query="SELECT * FROM cs_relationships WHERE idproduct="&idmain&" ORDER BY num ASC;"
		set rstemp=Server.CreateObject("ADODB.Recordset") 
		set rstemp=conntemp.execute(query) %>

		<table class="pcCPcontent">
			<% cnt=0
			do until rstemp.eof
		 	query="SELECT * FROM products WHERE idproduct="&rstemp("idrelation")&";"
			set rsRelation=Server.CreateObject("ADODB.Recordset") 
			set rsRelation=conntemp.execute(query) %>
                
			<tr> 
				<td width="70%"><%=rsRelation("description")%></td>
				<td width="30%">                               
					<input name="num<%=cnt%>" type="text" value="<%=rstemp("num")%>" size="3">
					<input name="idcrosssell<%=cnt%>" type="hidden" value="<%=rstemp("idcrosssell")%>">
				</td>
			</tr>
			<% cnt=cnt+1
			rstemp.moveNext
			loop
			set rstemp=nothing
			 %>
			<input name="icnt" type="hidden" value="<%=cnt%>">
			<tr> 
				<td height="10" colspan="2"></td>
			</tr>
			<tr> 
				<td align="center" colspan="2"> 
					<input type="submit" name="SubmitOrder" value="Finish" class="btn btn-primary">&nbsp;
					<input type="button" class="btn btn-default"  name="back" value="Back" onClick="location.href='crossSellAddb.asp'">
				</td>
			</tr>
		</table>
	</form>
<!--#include file="AdminFooter.asp"-->
