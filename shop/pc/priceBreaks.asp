<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<% 
dim pIdProduct, query2, rsTemp2, pDescription

pIdProduct=getUserInput(request.QueryString("idProduct"),0)
if not validNum(pIdProduct) then
	response.redirect "msg.asp?message=74"
end if

	'// Load icons
	Set conlayout=Server.CreateObject("ADODB.Connection")
	conlayout.Open scDSN
	set rsIconObj = server.CreateObject("ADODB.Recordset")
	Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

	'// Load discounts
	query="SELECT quantityFrom,quantityUntil,percentage,discountPerWUnit,discountPerUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" ORDER BY num"
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	set rsTemp=conntemp.execute(query)
	
	if rsTemp.eof then
	   set rsTemp=nothing
	   call closeDb()
	   response.redirect "msg.asp?message=74"
	end if

	if err.number<>0 then
		call LogErrorToDatabase()
		set rsTemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	'// Load product description
	query="SELECT description FROM products WHERE idProduct="& pIdProduct
	Set rsTemp2=Server.CreateObject("ADODB.Recordset")
	Set rsTemp2=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsTemp2=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	pDescription=rsTemp2("description")
	Set rsTemp2 = nothing
%> 
<div class="modal-header">
    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
    <h3 class="modal-title" id="pcDialogTitle"><%= dictLanguage.Item(Session("language")&"_pricebreaks_5")%><%=pDescription%></h3>
</div>
<div class="modal-body">
	<div class="pcShowProductQtyDiscounts">
    <div class="pcCartLayout pcShowList container-fluid">
      <div class="pcTableHeader row">
        <div class="col-xs-8"><%=dictLanguage.Item(Session("language")&"_pricebreaks_1")%></div>
        <div class="col-xs-4"><%=dictLanguage.Item(Session("language")&"_pricebreaks_2")%>&nbsp;<img src="<%=pcf_getImagePath("",rsIconObj("discount"))%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_6")%>"></div>
      </div>
    
			<% 
      do until rstemp.eof
      %>
      <div class="row">
        <div class="col-xs-8">
          <% if rstemp("quantityFrom")=rstemp("quantityUntil") then %>
            <%=rstemp("quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")%>
          <% else %>
            <%=rstemp("quantityFrom")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_3")&"&nbsp;"&rstemp("quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")%>
        <% end if %>
        </div>
        <div class="col-xs-4">
          <% If (request.querystring("Type")="1") or (session("CustomerType")="1") Then %>
            <% If rstemp("percentage")="0" then %>
            <%=scCurSign & money(rstemp("discountPerWUnit"))%> 
            <% else %>
            <%=rstemp("discountPerWUnit")%>%
            <% End If %>
          <% else %>
            <% If rstemp("percentage")="0" then %>
            <%=scCurSign & money(rstemp("discountPerUnit"))%> 
            <% else %>
            <%=rstemp("discountPerUnit")%>%
            <% End If %>
          <% end If %>
        </div>
      </div>
      <% 
      rstemp.moveNext
      loop
      set rsTemp = nothing
      %>
    </div>
    <div class="pcClear"></div>
	</div>
</div>
<div class="modal-footer">
    <button class="btn btn-default" data-dismiss="modal" type="button"><%=dictLanguage.Item(Session("language")&"_AddressBook_5")%></button>
</div>
<%
call closeDb()
%>
