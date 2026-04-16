<%PmAdmin=2%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
idRootCat=request("idRootCategory")
if idRootCat="" OR (NOT isNumeric(idRootCat)) then
	idRootCat=1
end if

myCats=request("strCats")
if myCats="" then
	myCats="0,"
end if

pcv_CP=request("CP")
if pcv_CP="" then
	pcv_CP="0"
end if
%>
<!--#include file="inc_pcRequestCategories.asp"-->