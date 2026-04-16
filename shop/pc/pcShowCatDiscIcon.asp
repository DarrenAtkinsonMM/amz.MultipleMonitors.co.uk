<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
        
<%
' *******************************
' * START Category discounts
' *******************************	
' check for discount per category
query="SELECT pcCD_idDiscount FROM pcCatDiscounts WHERE pcCD_idcategory="& intIdCategory
if session("CustomerType")<>"1" then
	query=query & " and pcCD_discountPerUnit<>0"
else
	query=query & " and pcCD_discountPerWUnit<>0"
end if
set rsDisc=Server.CreateObject("ADODB.Recordset")
set rsDisc=conntemp.execute(query)

if not rsDisc.eof then
	pCatDiscountPerQuantity=-1
else
	pCatDiscountPerQuantity=0
end if
set rsDisc = nothing
				
if pCatDiscountPerQuantity=-1 then %>
	<div class="pcShowProductCatDiscount">
    <a href="javascript:openbrowser('catDiscounts.asp?type=<%=Session("customerType")%>&idcategory=<%=intIdCategory%>&type=1')">
			<img src="<%=pcf_getImagePath("",rsIconObj("discount"))%>" alt="<%response.write dictLanguage.Item(Session("language")&"_viewPrd_16")%>">
    </a>
	</div>
<% end if
		
' *******************************
' * END Category discounts
' *******************************
%>
