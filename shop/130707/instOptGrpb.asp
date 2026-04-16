<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% 
dim f,g,z, prdFrom

poptionGroupDesc=request.querystring("optionGroupDesc")
poptionGroupDesc = replace(poptionGroupDesc,"'","''")

'get product ID of the referring product page, if any
prdFrom = request.QueryString("prdFrom")
if not validNum(prdFrom) then prdFrom = 0
AssignID = request.QueryString("AssignID")

tmpFGID=getUserInput(request("pcFGid"),0)
if tmpFGID="" then
	tmpFGID=0
end if

if trim(poptionGroupDesc)="" then
   call closeDb()
    response.redirect "msg.asp?message=16"
end if

' insert in to db new option group
query="INSERT INTO optionsGroups (optionGroupDesc) VALUES (N'" &poptionGroupDesc& "')"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	
	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in addoptiongroupexec: "&Err.Description) 
end If

if tmpFGID>"0" then
	query="SELECT idOptionGroup FROM optionsGroups WHERE optionGroupDesc like '" &poptionGroupDesc& "' ORDER BY idOptionGroup DESC;"
	set rstemp=connTemp.execute(query)
	if not rstemp.eof then
		tmpidOptGrp=rstemp("idOptionGroup")
		set rstemp=nothing
		
		query="INSERT INTO pcFGOG (pcFG_ID,idOptionGroup) VALUES (" & tmpFGID & "," & tmpidOptGrp & ");"
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
	end if
	set rstemp=nothing
end if		

query="SELECT Count(*) As TotalGrp FROM optionsGroups WHERE optionGroupDesc like '" &poptionGroupDesc& "';"
set rstemp=connTemp.execute(query)

Count=0
if not rstemp.eof then
	Count=rstemp("TotalGrp")
	if Count<>"" then
	else
		Count=0
	end if
	Count=Cint(Count)
end if
set rstemp=nothing

tmpStr=""
if Count>1 then
	tmpStr="<br /><br />Note: Option Group(s) with the same name already exist."
end if

' if the referring product exist, go back to that page, otherwise go to Manage Options
if prdFrom = 0 then
  call closeDb()
response.redirect "ManageOptions.asp?s=1&msg="&Server.Urlencode("Successfully added new Option Group." & tmpStr)
 else
  call closeDb()
response.redirect "modPrdOpta1.asp?AssignID="& AssignID &"&idproduct="& prdFrom
end if
%>
<!--#include file="AdminFooter.asp"-->
