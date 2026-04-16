<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<% 
dim pIdProduct, pDescription, pPrice, pDetails, pListPrice, pLgimageURL, pImageUrl, pWeight, pcv_strProductsArray

Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

categoryDescName = getUserInput(request.QueryString("cd"),0)
pcv_strProductsArray = getUserInput(request.QueryString("SIArray"),0)

'// Trim the last comma so we can use this feature with one item
if instr(pcv_strProductsArray,",")>0 then
	xStringLength = len(pcv_strProductsArray)
	if xStringLength>0 then
		pcv_strProductsArray = left(pcv_strProductsArray,(xStringLength-1))
	end if
end if
ProductArray = Split(pcv_strProductsArray,",")
%>
<div class="modal-header">
    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
    <h3 class="modal-title" id="pcDialogTitle"><%= dictLanguage.Item(Session("language")&"_pricebreaks_7")%></h3>
</div>
<div class="modal-body">
	<div class="pcShowProductQtyDiscounts">
    <%    
    For i = lbound(ProductArray) To UBound(ProductArray)
    
        pIdProduct=ProductArray(i)
        If validNum(pIdProduct) Then
        
            query="SELECT description FROM products WHERE idProduct="& pIdProduct
            Set rsTemp=Server.CreateObject("ADODB.Recordset")
            Set rsTemp=conntemp.execute(query)
            If Not rsTemp.Eof Then
                pDescription=rsTemp("description")
            End If
            Set rsTemp = Nothing
        
            query="SELECT quantityFrom,quantityUntil,percentage,discountPerWUnit,discountPerUnit FROM discountsperquantity WHERE idProduct="& pIdProduct &" ORDER BY num"
            Set rsTemp=Server.CreateObject("ADODB.Recordset")
            Set rsTemp=conntemp.execute(query)
            If Not rsTemp.Eof Then
                %>
                <h4><%=pDescription%></h4>
    
    			<div class="pcCartLayout pcShowList container-fluid">
      				<div class="pcTableHeader row">
                        <div class="col-xs-8"><%response.write dictLanguage.Item(Session("language")&"_pricebreaks_1")%></div>
                        <div class="col-xs-4"><%response.write dictLanguage.Item(Session("language")&"_pricebreaks_2")%>&nbsp;<img src="<%=pcf_getImagePath("",rsIconObj("discount"))%>"></div>
                    </div>
                    <%
                    Do Until rsTemp.Eof
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
                    rsTemp.MoveNext
                Loop
                %>
                </div>
            <% 
            End If 
            Set rsTemp = Nothing
            
        End If '// If validNum(pIdProduct) Then
    
    Next '// For i = lbound(ProductArray) To UBound(ProductArray)
    %>
    <div class="pcClear"></div>
	</div>
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
