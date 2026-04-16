<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
Server.ScriptTimeout=5400%>
<%PmAdmin=7%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
Dim AddrList()
Dim purchaseType

mCount=-1

SOptedIn=getUserInput(request("SOptedIn"),0)

SIDProduct=getUserInput(request("SIDProduct"),0)
SIDCategory=getUserInput(request("SIDCategory"),0)
	if not validNum(SIDProduct) then SIDProduct=0
	if not validNum(SIDCategory) then SIDCategory=0
	if SIDCategory>0 then SIDProduct=0

SCustType=getUserInput(request("SCustType"),0)

SStartDate=getUserInput(request("SStartDate"),0)
SEndDate=getUserInput(request("SEndDate"),0)
	if not IsDate(SStartDate) then SStartDate=""
	if not IsDate(SEndDate) then SEndDate=""

purchaseType=getUserInput(request("purchaseType"),0)
	if purchaseType="" then
		purchaseType=0
	end if
	'// If the customer's purchases are to be ignored, clear product and category IDs
	if purchaseType="0" then
		SIDProduct=0
		SIDCategory=0
	end if


Function checkNoOrderCustomer(COptedIn,CCustType)

Dim mrs4,query4, myTemp4, tmpArr, intCount, i

	myTemp4=" WHERE (idcustomer NOT IN (SELECT DISTINCT idcustomer FROM orders)) "

	if (COptedIn<>"") and (COptedIn<>"0") then
		if myTemp4="" then
			myTemp4=" WHERE "
		else
			myTemp4=myTemp4 & " AND "
		end if
		myTemp4=myTemp4 & " Customers.RecvNews=" & COptedIn
	end if
	
	if (CCustType<>"") and (CCustType<>"0") then
		if myTemp4="" then
			myTemp4=" WHERE "
		else
			myTemp4=myTemp4 & " AND "
		end if
		if CCustType="1" then
			myTemp4=myTemp4 & " Customers.customerType<>1 AND Customers.idCustomerCategory=0"
		end if
		if CCustType="2" then
			myTemp4=myTemp4 & " Customers.customerType=1 AND Customers.idCustomerCategory=0"
		end if
		if instr(CCustType,"CC_")>0 then
			tmp_Arr=split(CCustType,"CC_")
			if tmp_Arr(1)<>"" then
				myTemp4=myTemp4 & " Customers.idCustomerCategory=" & tmp_Arr(1)
			end if
		end if
	end if

	query4="SELECT DISTINCT Customers.email FROM Customers" & myTemp4 & " ORDER BY Customers.Email ASC;"
	set mrs4=server.CreateObject("ADODB.RecordSet")
	set mrs4=connTemp.execute(query4)
	
	if not mrs4.eof then
		tmpArr=mrs4.getRows()
		set mrs4=nothing
		intCount=ubound(tmpArr,2)
		For i=0 to intCOunt
			mCount=mCount+1
			ReDim Preserve AddrList(mCount)
			AddrList(mCount)=tmpArr(0,i)
		Next
	end if
	set mrs4=nothing

End Function


Function checkCustomer(CIDCustomer,COptedIn,CCustType)

Dim mrs4,query4, myTemp4, tmpArr, intCount, i

	myTemp4=""

	if (CIDCustomer<>"") and (CIDCustomer<>"0") then
		if myTemp4="" then
			myTemp4=" WHERE "
		else
			myTemp4=myTemp4 & " AND "
		end if
		myTemp4=myTemp4 & " Customers.idCustomer=" & CIDCustomer
	end if

	if purchaseType="1" or purchaseType="3" then
		if myTemp4="" then
			myTemp4=" WHERE "
		else
			myTemp4=myTemp4 & " AND "
		end if
		myTemp4=myTemp4 & " (Customers.idcustomer IN (SELECT DISTINCT idcustomer FROM orders)) "
	end if
	
	if (COptedIn<>"") and (COptedIn<>"0") then
		if myTemp4="" then
			myTemp4=" WHERE "
		else
			myTemp4=myTemp4 & " AND "
		end if
		myTemp4=myTemp4 & " Customers.RecvNews=" & COptedIn
	end if
	
	if (CCustType<>"") and (CCustType<>"0") then
		if myTemp4="" then
			myTemp4=" WHERE "
		else
			myTemp4=myTemp4 & " AND "
		end if
		if CCustType="1" then
			myTemp4=myTemp4 & " Customers.customerType<>1 AND Customers.idCustomerCategory=0"
		end if
		if CCustType="2" then
			myTemp4=myTemp4 & " Customers.customerType=1 AND Customers.idCustomerCategory=0"
		end if
		if instr(CCustType,"CC_")>0 then
			tmp_Arr=split(CCustType,"CC_")
			if tmp_Arr(1)<>"" then
				myTemp4=myTemp4 & " Customers.idCustomerCategory=" & tmp_Arr(1)
			end if
		end if
	end if

	query4="SELECT DISTINCT Customers.email FROM Customers" & myTemp4 & " ORDER BY Customers.Email ASC;"
	set mrs4=server.CreateObject("ADODB.RecordSet")
	set mrs4=connTemp.execute(query4)

	if not mrs4.eof then
		tmpArr=mrs4.getRows()
		set mrs4=nothing
		intCount=ubound(tmpArr,2)
		For i=0 to intCOunt
			mCount=mCount+1
			ReDim Preserve AddrList(mCount)
			AddrList(mCount)=tmpArr(0,i)
		Next
	end if

	set mrs4=nothing

End Function

Function checkOrder(CIDOrder,COptedIn,CStartDate,CEndDate)

Dim mrs3,query3, myTemp3, tmpArr, intCount, i

	myTemp3=""

	if (CIDOrder<>"") and (CIDOrder<>"0") then
		if myTemp3="" then
			myTemp3=" WHERE "
		else
			myTemp3=myTemp3 & " AND "
		end if
		myTemp3=myTemp3 & " Orders.idOrder=" & CIDOrder
	end if

	if (CStartDate<>"") and (IsDate(CStartDate)) then
		if myTemp3="" then
			myTemp3=" WHERE "
		else
			myTemp3=myTemp3 & " AND "
		end if
		myTemp3=myTemp3 & " Orders.orderDate>='" & CStartDate & "'"
	end if

	if (CEndDate<>"") and (IsDate(CEndDate)) then
		if myTemp3="" then
			myTemp3=" WHERE "
		else
			myTemp3=myTemp3 & " AND "
		end if
		myTemp3=myTemp3 & " Orders.orderDate<='" & CEndDate & "'"
	end if
	
	if (COptedIn<>"") and (COptedIn<>"0") then
		if myTemp3="" then
			myTemp3=" WHERE "
		else
			myTemp3=myTemp3 & " AND "
		end if
		myTemp3=myTemp3 & " Customers.RecvNews=" & COptedIn
	end if
	
	if (CCustType<>"") and (CCustType<>"0") then
		if myTemp3="" then
			myTemp3=" WHERE "
		else
			myTemp3=myTemp3 & " AND "
		end if
		if CCustType="1" then
			myTemp3=myTemp3 & " Customers.customerType<>1 AND Customers.idCustomerCategory=0"
		end if
		if CCustType="2" then
			myTemp3=myTemp3 & " Customers.customerType=1 AND Customers.idCustomerCategory=0"
		end if
		if instr(CCustType,"CC_")>0 then
			tmp_Arr=split(CCustType,"CC_")
			if tmp_Arr(1)<>"" then
				myTemp3=myTemp3 & " Customers.idCustomerCategory=" & tmp_Arr(1)
			end if
		end if
	end if

	query3="SELECT DISTINCT Customers.email FROM Customers INNER JOIN Orders ON Customers.idCustomer=Orders.idCustomer " & myTemp3 & " ORDER BY Customers.Email ASC;"
	set mrs3=server.CreateObject("ADODB.RecordSet")
	set mrs3=connTemp.execute(query3)

	if not mrs3.eof then
		tmpArr=mrs3.getRows()
		set mrs3=nothing
		intCount=ubound(tmpArr,2)
		For i=0 to intCOunt
			mCount=mCount+1
			ReDim Preserve AddrList(mCount)
			AddrList(mCount)=tmpArr(0,i)
		Next
	end if

	set mrs3=nothing

End Function

Function checkproduct(CIDProduct,COptedIn)

dim mrs3,query3, myTemp3, tmpArr, intCount, i
	
	myTemp3=""

	if (CStartDate<>"") and (IsDate(CStartDate)) then
		myTemp3=myTemp3 & " AND "
		myTemp3=myTemp3 & " Orders.orderDate>='" & CStartDate & "'"
	end if

	if (CEndDate<>"") and (IsDate(CEndDate)) then
		myTemp3=myTemp3 & " AND "
		myTemp3=myTemp3 & " Orders.orderDate<='" & CEndDate & "'"
	end if
	
	if (COptedIn<>"") and (COptedIn<>"0") then
		myTemp3=myTemp3 & " AND "
		myTemp3=myTemp3 & " Customers.RecvNews=" & COptedIn
	end if
	
	if (CCustType<>"") and (CCustType<>"0") then
		myTemp3=myTemp3 & " AND "
		if CCustType="1" then
			myTemp3=myTemp3 & " Customers.customerType<>1 AND Customers.idCustomerCategory=0"
		end if
		if CCustType="2" then
			myTemp3=myTemp3 & " Customers.customerType=1 AND Customers.idCustomerCategory=0"
		end if
		if instr(CCustType,"CC_")>0 then
			tmp_Arr=split(CCustType,"CC_")
			if tmp_Arr(1)<>"" then
				myTemp3=myTemp3 & " Customers.idCustomerCategory=" & tmp_Arr(1)
			end if
		end if
	end if

	APPquery="(SELECT idproduct FROM Products WHERE ((IDProduct=" & CIDProduct & ") OR (pcprod_ParentPrd=" & CIDProduct & ")) AND removed=0 AND (((pcprod_ParentPrd>0) AND (active=0) AND (pcProd_SPInActive=0)) OR (active<>0)))"
	
	if purchaseType="1" then
		query3="SELECT DISTINCT Customers.email FROM Customers INNER JOIN (Orders INNER JOIN ProductsOrdered ON Orders.idOrder=ProductsOrdered.idOrder) ON Customers.idCustomer=Orders.idCustomer WHERE (ProductsOrdered.idproduct IN " & APPquery & ") " & myTemp3 & " ORDER BY Customers.Email ASC;"
	elseif purchaseType="2" then
		query3="SELECT DISTINCT Customers.email FROM Customers INNER JOIN Orders ON Customers.idCustomer=Orders.idCustomer WHERE (Orders.idcustomer NOT IN (SELECT DISTINCT idcustomer FROM Orders INNER JOIN ProductsOrdered ON Orders.idorder=ProductsOrdered.idOrder WHERE ProductsOrdered.idproduct IN " & APPquery & ")) " & myTemp3 & " ORDER BY Customers.Email ASC;"
	else
		query3=""
	end if

	if query3<>"" then
		set mrs3=connTemp.execute(query3)
		if not mrs3.eof then
			tmpArr=mrs3.getRows()
			set mrs3=nothing
			intCount=ubound(tmpArr,2)
			For i=0 to intCOunt
				mCount=mCount+1
				ReDim Preserve AddrList(mCount)
				AddrList(mCount)=tmpArr(0,i)
			Next
		end if
		
		set mrs3=nothing
		
	end if

End function

Function checkcategory(CIDCategory,COptedIn)

	dim mrs3,query3, myTemp3, tmpArr, intCount, i

	IF purchaseType="1" THEN
	
		myTemp3=""

		if (CStartDate<>"") and (IsDate(CStartDate)) then
			myTemp3=myTemp3 & " AND "
			myTemp3=myTemp3 & " Orders.orderDate>='" & CStartDate & "'"
		end if
	
		if (CEndDate<>"") and (IsDate(CEndDate)) then
			myTemp3=myTemp3 & " AND "
			myTemp3=myTemp3 & " Orders.orderDate<='" & CEndDate & "'"
		end if
		
		if (COptedIn<>"") and (COptedIn<>"0") then
			myTemp3=myTemp3 & " AND "
			myTemp3=myTemp3 & " Customers.RecvNews=" & COptedIn
		end if
		
		if (CCustType<>"") and (CCustType<>"0") then
			myTemp3=myTemp3 & " AND "
			if CCustType="1" then
				myTemp3=myTemp3 & " Customers.customerType<>1 AND Customers.idCustomerCategory=0"
			end if
			if CCustType="2" then
				myTemp3=myTemp3 & " Customers.customerType=1 AND Customers.idCustomerCategory=0"
			end if
			if instr(CCustType,"CC_")>0 then
				tmp_Arr=split(CCustType,"CC_")
				if tmp_Arr(1)<>"" then
					myTemp3=myTemp3 & " Customers.idCustomerCategory=" & tmp_Arr(1)
				end if
			end if
		end if
		
		queryPrd="SELECT DISTINCT Customers.email FROM Customers,Orders,ProductsOrdered,categories_products WHERE (ProductsOrdered.idproduct=categories_products.idproduct) AND (categories_products.idcategory=" & CIDCategory & ") AND (Orders.idOrder=ProductsOrdered.idOrder) AND (Customers.idCustomer=Orders.idCustomer)" & myTemp3 & ""
		
		querySubPrd="SELECT DISTINCT Customers.email FROM Customers,Orders,ProductsOrdered,categories_products,products WHERE (ProductsOrdered.idproduct=products.idproduct) AND (categories_products.idproduct=products.pcProd_ParentPrd) AND (categories_products.idcategory=" & CIDCategory & ") AND (Orders.idOrder=ProductsOrdered.idOrder) AND (Customers.idCustomer=Orders.idCustomer)" & myTemp3 & ""
		
		query3 = queryPrd & " UNION " & querySubPrd & " ORDER BY email ASC;"
		set mrs3=connTemp.execute(query3)

		if not mrs3.eof then
			tmpArr=mrs3.getRows()
			set mrs3=nothing
			intCount=ubound(tmpArr,2)
			For i=0 to intCOunt
				mCount=mCount+1
				ReDim Preserve AddrList(mCount)
				AddrList(mCount)=tmpArr(0,i)
			Next
		end if
		
		set mrs3=nothing
		
	ELSEIF purchaseType="2" THEN
		myTemp3=""

		if (CStartDate<>"") and (IsDate(CStartDate)) then
			myTemp3=myTemp3 & " AND "
			myTemp3=myTemp3 & " Orders.orderDate>='" & CStartDate & "'"
		end if
	
		if (CEndDate<>"") and (IsDate(CEndDate)) then
			myTemp3=myTemp3 & " AND "
			myTemp3=myTemp3 & " Orders.orderDate<='" & CEndDate & "'"
		end if
		
		if (COptedIn<>"") and (COptedIn<>"0") then
			myTemp3=myTemp3 & " AND "
			myTemp3=myTemp3 & " Customers.RecvNews=" & COptedIn
		end if
		
		if (CCustType<>"") and (CCustType<>"0") then
			myTemp3=myTemp3 & " AND "
			if CCustType="1" then
				myTemp3=myTemp3 & " Customers.customerType<>1 AND Customers.idCustomerCategory=0"
			end if
			if CCustType="2" then
				myTemp3=myTemp3 & " Customers.customerType=1 AND Customers.idCustomerCategory=0"
			end if
			if instr(CCustType,"CC_")>0 then
				tmp_Arr=split(CCustType,"CC_")
				if tmp_Arr(1)<>"" then
					myTemp3=myTemp3 & " Customers.idCustomerCategory=" & tmp_Arr(1)
				end if
			end if
		end if
	
		queryPrd="SELECT DISTINCT idcustomer FROM Orders INNER JOIN ProductsOrdered ON Orders.idorder=ProductsOrdered.idOrder WHERE (ProductsOrdered.idproduct IN (SELECT DISTINCT idproduct FROM categories_products WHERE idcategory=" & CIDCategory & "))"
		
		querySubPrd="SELECT DISTINCT idcustomer FROM Orders INNER JOIN ProductsOrdered ON Orders.idorder=ProductsOrdered.idOrder INNER JOIN products ON products.idproduct=ProductsOrdered.idproduct WHERE (products.pcProd_ParentPrd IN (SELECT DISTINCT idproduct FROM categories_products WHERE idcategory=" & CIDCategory & "))"
		query3="SELECT DISTINCT Customers.email FROM Customers INNER JOIN Orders ON Customers.idCustomer=Orders.idCustomer WHERE (Orders.idcustomer NOT IN ((" & queryPrd & ") UNION (" & querySubPrd & ")) " & myTemp3 & " ORDER BY Customers.Email ASC;"
		set mrs3=connTemp.execute(query3)

		if not mrs3.eof then
			tmpArr=mrs3.getRows()
			set mrs3=nothing
			intCount=ubound(tmpArr,2)
			For i=0 to intCOunt
				mCount=mCount+1
				ReDim Preserve AddrList(mCount)
				AddrList(mCount)=tmpArr(0,i)
			Next
		end if
		
		set mrs3=nothing
	END IF

End Function


'// Determine query to run
if purchaseType="3" then
	if SEndDate="" and SStartDate="" then
		call checkCustomer("0",SOptedIn,SCustType)
	else
		call checkOrder("0",SOptedIn,SStartDate,SEndDate)
	end if
elseif purchaseType="4" then
	call checkNoOrderCustomer(SOptedIn,SCustType)
else
	if (SIDCategory<>"") and (SIDCategory<>"0") then
		call checkcategory(SIDCategory,SOptedIn)
	else
		if (SIDProduct<>"") and (SIDProduct<>"0") then
			call checkproduct(SIDProduct,SOptedIn)
		else
			if (SStartDate<>"" AND SEndDate<>"") then
				call checkOrder("0",SOptedIn,SStartDate,SEndDate)
			else
				if purchaseType="2" then
					call checkNoOrderCustomer(SOptedIn,SCustType)
				else
					call checkCustomer("0",SOptedIn,SCustType)
				end if
			end if
		end if
	end if
	if purchaseType="2" and SEndDate="" and SStartDate="" then
		call checkNoOrderCustomer(SOptedIn,SCustType)
	end if
end if

'Sort e-mail address list
if mCount>=0 then

	AList=AddrList
	
	session("AddrList")=AList
	session("AddrCount")=ubound(AList)

else

	dim BList(1)
	session("AddrList")=BList
	session("AddrCount")=0

end if


call closeDb()
response.redirect "newsWizStep2.asp?from=1"
%>