<%
query="SELECT pcProd_ParentPrd AS IDParent,SUM(sales) AS totalSum FROM Products WHERE (pcProd_ParentPrd IN (SELECT DISTINCT Products.pcProd_ParentPrd FROM Products INNER JOIN ProductsOrdered ON Products.idProduct=ProductsOrdered.idProduct WHERE Products.pcProd_ParentPrd>0 AND ProductsOrdered.idOrder=" & pIdOrder & ")) GROUP BY pcProd_ParentPrd;"
set rsQ=connTemp.execute(query)
if not rsQ.eof then
	pcArrayQ=rsQ.getRows()
	intCountQ=ubound(pcArrayQ,2)
	set rsQ=nothing
	
	For kQ=0 to intCountQ
		if clng(pcArrayQ(0,kQ))<>0 then
			query="UPDATE Products SET sales=" & pcArrayQ(1,kQ) & " WHERE IDProduct=" & pcArrayQ(0,kQ)
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		end if
	Next
end if
set rsQ=nothing
%>
