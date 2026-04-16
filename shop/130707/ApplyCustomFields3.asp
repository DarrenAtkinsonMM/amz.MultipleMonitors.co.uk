<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Copy custom fields to other products" %>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_UpdateDates.asp" -->
<%
Dim rsOrd, pid

CustFieldCopy=session("CustFieldCopy")
idproduct=session("ACidproduct")

if request("action")="apply" then

  CN1=""
  CC1=""
  
  CFieldType=0
 
  if instr(CustFieldCopy,"custom")>0 then
	  CC1=replace(CustFieldCopy,"custom","")
	  query="SELECT idSearchField FROM pcSearchData WHERE idSearchData=" & CC1 & ";"
	  set rsQ=connTemp.execute(query)
	  CN1=0
	  if not rsQ.eof then
	  	CN1=rsQ("idSearchField")
	  end if
	  set rsQ=nothing
	  CFieldType=1
  end if
  
  if instr(CustFieldCopy,"xfield")>0 then
	  CN1=replace(CustFieldCopy,"xfield","")
	  CC1=0
	  query="SELECT pcPXF_XReq FROM pcPrdXFields WHERE IdXField=" & CN1 & " AND IdProduct=" & idproduct & ";"
	  set rstemp=connTemp.execute(query)
	  if not rstemp.eof then
		CC1=rstemp("pcPXF_XReq")
	  end if
	  CFieldType=2
  end if
  
 RSu=0
 RFa=0

If (request("prdlist")<>"") and (request("prdlist")<>",") then
	prdlist=split(request("prdlist"),",")
	For i=lbound(prdlist) to ubound(prdlist)
		id=prdlist(i)
		IF (id<>"0") AND (id<>"") THEN
			IF CFieldType=2 THEN
				query="DELETE FROM pcPrdXFields WHERE idProduct=" & id & " AND IdXField=" & CN1 & ";"
				set rstemp=connTemp.execute(query)
				set rstemp=nothing

				query="INSERT INTO pcPrdXFields (IdProduct,IdXField,pcPXF_XReq) VALUES (" & id & "," & CN1 & "," & CC1 & ");"
				set rstemp=connTemp.execute(query)
				set rstemp=nothing

				RSu=RSu+1
			ELSE
				query="DELETE FROM pcSearchFields_Products WHERE idproduct=" & id & " AND idSearchData IN (SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & CN1 & ");"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing

				query="INSERT INTO pcSearchFields_Products (idproduct,idSearchData) VALUES (" & id & "," & CC1 & ");"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing
				
				RSu=RSu+1
			END IF
			
			call updPrdEditedDate(id)
			
		END IF 'have id product
	Next
End if



end if 'action=apply
%>
<!--#include file="AdminHeader.asp"-->

	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>

			<div class="pcCPmessageSuccess">
            	The selected custom field was copied to <%=RSu%> products.
				<%if RFa>0 then%>
                    <br>
                    <%=RFa%> of the selected products could not be updated because they already had the maximum allowed number of search or input fields assigned to them.
                <%end if%>
            	<div><a href="AdminCustom.asp?idproduct=<%=idproduct%>">Return to product's Custom Fields page</a>.</div>
            </div>                   
<!--#include file="AdminFooter.asp"-->
