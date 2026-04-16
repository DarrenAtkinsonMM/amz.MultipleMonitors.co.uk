<%
    '// This feature deprecated in v5.1.00.  

	'query="UPDATE Products SET stock=0 WHERE (Products.pcprod_Apparel=1);"
	'set rsU=Server.CreateObject("ADODB.Recordset")
	'set rsU=connTemp.execute(query)
	'set rsU=nothing
	
	'query="UPDATE A SET A.stock=1 FROM Products A,Products B WHERE (A.IdProduct=B.pcProd_ParentPrd) AND (A.pcprod_Apparel=1) AND (A.active<>0) AND (A.removed=0) AND (B.pcProd_ParentPrd>0) AND (B.Stock>0) AND (B.removed=0) AND (B.active=0) AND (B.pcProd_SPInActive=0);"
	'set rsU=Server.CreateObject("ADODB.Recordset")
	'set rsU=connTemp.execute(query)
	'set rsU=nothing
%>
