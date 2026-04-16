<%@ LANGUAGE="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "category-bulk-add-to-cart-db.asp"
' This page adds the saved wishlist products into the cart.
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/common.asp"-->
<%
    call opendb()  
    dim intIsActive, strCurSkuPos, strCurSkuVal, intCurSkuStock, strRtnMessage, strCurSkuQty, strAllAdded, intReset, intAddToCart, strPostIt, intAddCnt
    
    'add to cart
    intAddToCart=getUserInput(request("addtocart"),1)
    if len(intAddToCart)>0 then
        strAllAdded=session("bulkcategoryadd")
        arrAllAdded=split(strAllAdded,"||")
        intAllAddedCount=ubound(arrAllAdded)-1
        intAddCnt=0
        y=0
        for x=0 to 6
             if y<=intAllAddedCount then 
                strCurSkuVal=arrAllAdded(y) 
                strCurSkuQty=arrAllAdded(y+1) 
             y=y+1
                if strCurSkuQty="" then
                    strCurSkuQty=1
                end if
            else 
                strCurSkuVal=""
                strCurSkuQty="1"
            end if 
            if len(strCurSkuVal)>0 then
                query="select idproduct from products where sku='"&strCurSkuVal&"'"
                set rs=server.CreateObject("ADODB.RecordSet")
	            set rs=conntemp.execute(query)
                if not rs.eof then
                    strPostIt=strPostIt&"idproduct"&x+1&"="&rs("idproduct")&"&QtyM"&rs("idProduct")&"="&strCurSkuQty&"&"
                    intAddCnt=intAddCnt+1
                end if
                set rs=nothing
            else
                
            end if
            y=y+1
        next 
        strPostIt=strPostIt&"pcnt="&intAddCnt
        response.write "0|"&strPostIt
        response.End
    end if


    'clear form 
    intReset=getUserInput(request("reset"),1)
    if len(intReset)>0 then
         session("bulkcategoryadd")=""
         response.end
    end if

    strCurSkuPos=getUserInput(request("curskupos"),25)
    strCurSkuVal=getUserInput(request("curskuval"),25)
    strCurSkuQty=getUserInput(request("curskuqty"),5)
    strCurSkuQty=parseInt(strCurSkuQty)
    strAllAdded=getUserInput(request("all"),0)

    if len(strCurSkuVal)>0 then
        query="select sku, nostock, stock, active, formQuantity from products where sku='"&strCurSkuVal&"'"
        set rs=server.CreateObject("ADODB.RecordSet")
	    set rs=conntemp.execute(query)
	    if rs.EOF then
            strRtnMessage="1|Product Not Found"
        else
            if rs("active")=0 or rs("formQuantity")=-1 then
                strRtnMessage="1|Currently Unavailable"
            else
                strRtnMessage="0|Product Found"
                session("bulkcategoryadd")=strAllAdded
            end if

        end if
        set rs=nothing
    end if
    call closedb()  

   response.write strRtnMessage
  %>
