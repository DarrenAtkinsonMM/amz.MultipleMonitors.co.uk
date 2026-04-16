<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
dim initialize
initialize=0

pIdDbSession=session("pcSFIdDbSession")
pRandomKey=session("pcSFRandomKey")

HaveToRefeshCustomerCache=""

' if dbSession was not defined
if pIdDbSession="" or pRandomKey="" then
	initialize=-1
end if

' check if current pcCustomerSessions is valid
if initialize=0 AND HaveToRefeshCustomerCache<>"1" then
	pcCustSession_Date=Date()
	if SQL_Format="1" then
		pcCustSession_Date=Day(pcCustSession_Date)&"/"&Month(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
	else
		pcCustSession_Date=Month(pcCustSession_Date)&"/"&Day(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
	end if
	
	TmpQuery="SELECT idDbSession FROM pcCustomerSessions WHERE randomKey="&pRandomKey& " AND pcCustSession_Date='" &pcCustSession_Date& "' ORDER BY pcCustomerSessions.idDbSession DESC;"
	set rsTmpObj=Server.CreateObject("ADODB.Recordset")
	set rsTmpObj=conntemp.execute(TmpQuery)
	if rsTmpObj.eof then
		' invalid pcCustomerSessions
		session("pcSFIdDbSession") = ""
		session("pcSFRandomKey") = ""
		response.redirect "msg.asp?message=38"
	end if
	set rsTmpObj=nothing

end if

if initialize=-1 OR HaveToRefeshCustomerCache="1" then
	pRandomKey=randomNumber(99999999)
	session("pcSFRandomKey")=pRandomKey
	pcCustSession_Date=Date()
	if SQL_Format="1" then
		pcCustSession_Date=Day(pcCustSession_Date)&"/"&Month(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
	else
		pcCustSession_Date=Month(pcCustSession_Date)&"/"&Day(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
	end if
	
	query="INSERT INTO pcCustomerSessions (randomKey, idCustomer, pcCustSession_Date) VALUES (" &pRandomKey& ","&session("idCustomer")&", '" &pcCustSession_Date& "')"
	set rs=Server.CreateObject("ADODB.Recordset")
 	set rs=conntemp.execute(query)

 	if err.number <> 0 then
        call LogErrorToDatabase()
        set rs = Nothing
        call closeDb()
        response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

 	' get pcCustomerSessions 
	query="SELECT idDbSession FROM pcCustomerSessions WHERE randomKey="&pRandomKey& " AND idCustomer="&session("idCustomer")&" AND pcCustSession_Date='" &pcCustSession_Date& "' ORDER BY idDbSession DESC;"
 	set rs=conntemp.execute(query)
 	pIdDbSession=rs("idDbSession")
	session("pcSFIdDbSession")=pIdDbSession
	set rs=nothing
	
end if

if session("idCustomer")>"0" then
	
	query="UPDATE pcCustomerSessions SET idCustomer="&session("idCustomer")&" WHERE randomKey="&pRandomKey& " AND idDbSession=" & pIdDbSession & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
	
end if

'// When customer logs in, then transfer cart codes from memory to customer session table
if session("pcSFCust_discountcode")<>"" then
	
    query="UPDATE pcCustomerSessions SET pcCustSession_total=" & session("pcSFCust_total") & ",pcCustSession_DiscountCodeTotal=" & session("pcSFCust_DiscountCodeTotal") & ", pcCustSession_discountAmount='" & session("pcSFCust_discountAmount") & "', pcCustSession_discountcode='" & session("pcSFCust_discountcode") & "' WHERE randomKey="&pRandomKey& " AND idDbSession=" & pIdDbSession
	set rs=connTemp.execute(query)
	set rs=nothing
    
    If (Not (session("pcSFIdDbSession") = "" OR session("pcSFRandomKey") = "")) AND (session("idCustomer")>"0") Then
        session("pcSFCust_FromCart")=""
        session("pcSFCust_discountcode")=""
        session("pcSFCust_DiscountCodeTotal")=""
        session("pcSFCust_discountAmount")=""
        session("pcSFCust_total")=""
    End If
    
end if

' randomNumber function, generates a number between 1 and limit
function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+2
end function
%>