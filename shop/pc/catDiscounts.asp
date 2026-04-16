<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<% 
dim pIdCategory, query2, rsTemp2, pDescription

pIdCategory=request.QueryString("idCategory")

if trim(pIdCategory)="" or not validNum(pIdCategory) then
   response.redirect "msg.asp?message=86"
end if

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

query="SELECT pcCD_quantityfrom,pcCD_quantityUntil,pcCD_percentage,pcCD_discountPerWUnit,pcCD_discountPerUnit FROM pcCatDiscounts WHERE pcCD_idcategory="& pIdCategory &" ORDER BY pcCD_num"
set rsTemp=Server.CreateObject("ADODB.Recordset")
set rsTemp=conntemp.execute(query)

query="SELECT categoryDesc,idcategory FROM categories WHERE idcategory="&pIdCategory
Set rsTemp2=Server.CreateObject("ADODB.Recordset")
Set rsTemp2=conntemp.execute(query)
pDescription=rsTemp2("categoryDesc")
Set rsTemp2 = nothing
%> 
<div class="modal-header">
    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
    <h3 class="modal-title" id="pcDialogTitle"><%= dictLanguage.Item(Session("language")&"_pricebreaks_5")%><%=pDescription%></h3>
</div>
<div class="modal-body">

  	<p><%= dictLanguage.Item(Session("language")&"_pricebreaks_6")%></p>
    
    <div class="pcTable">
    	<div class="pcTableHeader">
    		<div class="pcQtyDiscQuantity">
					<%= dictLanguage.Item(Session("language")&"_pricebreaks_1")%>
        </div>
    		<div class="pcQtyDiscSave">
					<%= dictLanguage.Item(Session("language")&"_pricebreaks_2")%>
					<img src="<%=pcf_getImagePath("",rsIconObj("discount"))%>" border="0">
        </div>
      </div>
    
			<% do until rstemp.eof %>
     
      <div class="pcTableRow">
        
        <div class="pcQtyDiscQuantity">
          <% if rstemp("pcCD_quantityFrom")=rstemp("pcCD_quantityUntil") then %>
            <%=rstemp("pcCD_quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")%>
          <% else %>
            <%=rstemp("pcCD_quantityFrom")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_3")&"&nbsp;"&rstemp("pcCD_quantityUntil")&"&nbsp;"& dictLanguage.Item(Session("language")&"_pricebreaks_4")%>
        	<% end if %>
        </div>
        
        <div class="pcQtyDiscSave">
          <% If (request.querystring("Type")="1") or (session("CustomerType")="1") Then %>
            <% If rstemp("pcCD_percentage")="0" then %>
            <%=scCurSign & money(rstemp("pcCD_discountPerWUnit"))%> 
            <% else %>
            <%=rstemp("pcCD_discountPerWUnit")%>%
            <% End If %>
          <% else %>
            <% If rstemp("pcCD_percentage")="0" then %>
            <%=scCurSign & money(rstemp("pcCD_discountPerUnit"))%> 
            <% else %>
            <%=rstemp("pcCD_discountPerUnit")%>%
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
<div class="modal-footer">
    <button class="btn btn-default" data-dismiss="modal" type="button"><%=dictLanguage.Item(Session("language")&"_AddressBook_5")%></button>
</div>
<%
call closeDb()
conlayout.Close
Set conlayout=nothing
Set rsIconObj = nothing 
%>
