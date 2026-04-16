<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
pcv_CategoryName=getUserInput(request("CategoryName"),0)
If pcv_CategoryName="" then
	response.write "<div class=pcCPmessage style='margin: -40px 0 0 100px; width: 600px;'>Please specify a name for this new category</div>"
	response.End()
end if
pcv_ParentCatID=getUserInput(request("ParentCatID"),0)
if NOT validNum(pcv_ParentCatID) then
	response.write "<div class=pcCPmessage style='margin: -40px 0 0 100px; width: 600px;'>Please select parent category</div>"
	response.End()
end if

' Open connection to the database
dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open scDSN

'// See if product already exists
queryA="Select idCategory from categories where categorydesc='" & pcv_CategoryName & "' and idParentCategory="&pcv_ParentCatID
set rs = conn.execute(queryA)

if rs.eof then
	queryB = "INSERT INTO categories (categorydesc,idParentCategory) VALUES(N'"&pcv_CategoryName&"', "&pcv_ParentCatID&");"
	set rs = conn.execute(queryB)
	set rs = conn.execute(queryA)
	response.write "<div class=pcCPmessageSuccess style='margin: -40px 0 0 100px; width: 600px;'>The new category was successfully added and is now available in the list below.</div>"
else
	response.write "<div class=pcCPmessage style='margin: -40px 0 0 100px; width: 600px;'>The category was not added because it already exists in the store database.</div>"
end if
set rs=nothing
conn.Close

set conn=nothing

%>