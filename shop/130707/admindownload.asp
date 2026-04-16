<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
on error resume next

Dim lngIDFile,intTest,lngIDFeedback,strFileName,lngIDOrder,strFileN,strFileName1

lngIDFile=getUserInput(request("IDFile"),0)

intTest=0
	
MySQL="select * from pcUploadFiles where pcUpld_IDFile=" & lngIDFile
set rstemp=connTemp.execute(mySQL)

if rstemp.eof then
	intTest=1
else
	lngIDFeedback=rstemp("pcUpld_IDFeedback")
	strFileName=rstemp("pcUpld_FileName")
end if
	
if intTest=0 then

	MySQL="select * from pcComments where pcComm_IDFeedback=" & lngIDFeedback
	set rstemp=connTemp.execute(mySQL)

	if rstemp.eof then
		intTest=1
	else
		lngIDOrder=rstemp("pcComm_IDOrder")
	end if

end if
	
if intTest=0 then

	MySQL="Select * from Orders where IDOrder=" & lngIDOrder
	set rstemp=connTemp.execute(mySQL)

	if rstemp.eof then
		intTest=1
	end if
end if
 
if intTest=1 then
	 call closeDb()
response.redirect "about:blank"
end if

MySQL="select * from pcUploadFiles where pcUpld_IDFile=" & lngIDFile
set rstemp=connTemp.execute(mySQL)

lngIDFeedback=rstemp("pcUpld_IDFeedback")
strFileName=rstemp("pcUpld_FileName")

'Downloadable file name

strFileN="../pc/Library/" & strFileName
strFileName1=mid(strFileName,instr(strFileName,"_")+1,len(strFileName))

set rstemp = nothing


call closeDb()
response.redirect strFileN

%>
