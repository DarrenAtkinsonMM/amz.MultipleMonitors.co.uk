<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%Response.ContentType = "text/xml"%>
<?xml version="1.0" ?>
<%
totalrecords=0

iPageSize=20
iPageCurrent=request("iPageCurrent")
if iPageCurrent="" then
	iPageCurrent=1
end if

rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open session("sm_query"), Conntemp, adOpenStatic, adLockReadOnly, adCmdText
rstemp.AbsolutePage=iPageCurrent

Dim strCol, Count, HTMLResult
HTMLResult=""
Count = 0
strCol = "#E1E1E1"

HTMLResult=HTMLResult & "<form action=""sm_xmlgetPrds.asp"" name=""srcresult"" class=""pcForms"">" & vbcrlf
HTMLResult=HTMLResult & "<table class=""pcCPcontent"">" & vbcrlf
HTMLResult=HTMLResult & "<tr>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""10%"">SKU</th>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""60%"">Product</th>" & vbcrlf
HTMLResult=HTMLResult & "<th width=""30%"">&nbsp;</th>" & vbcrlf
HTMLResult=HTMLResult & "</tr><tr><td colspan='3' class='pcCPSpacer'></td></tr>" & vbcrlf

do while (not rsTemp.eof) and (count < rsTemp.pageSize)
				
	If strCol <> "#FFFFFF" Then
		strCol = "#FFFFFF"
	Else 
		strCol = "#E1E1E1"
	End If
	count=count + 1
	pidProduct=trim(rstemp("idProduct"))
	psku=rstemp("sku")
	pDescription=rstemp("description")
	pcApparel=rstemp("pcProd_Apparel")
	HTMLResult=HTMLResult & "<tr onMouseOver=""this.className='activeRow'"" onMouseOut=""this.className='cpItemlist'"" class=""cpItemlist"">" & vbcrlf
	HTMLResult=HTMLResult & "<td>" & psku & "</td>" & vbcrlf
	HTMLResult=HTMLResult & "<td><a href='FindProductType.asp?id=" & pidProduct & "' target='_blank'>"
	if pcApparel="1" then
		HTMLResult=HTMLResult & pdescription & " and its sub-product(s)</a></td>" & vbcrlf
	else
		HTMLResult=HTMLResult & pdescription & "</a></td>" & vbcrlf
	end if
	HTMLResult=HTMLResult & "<td>&nbsp;</td>" & vbcrlf
	HTMLResult=HTMLResult & "</tr>" & vbcrlf

rsTemp.MoveNext
loop
HTMLResult=HTMLResult & "</table>" & vbcrlf
HTMLResult=HTMLResult & "<input type=hidden name=count value=""" & count & """>" & vbcrlf
HTMLResult=HTMLResult & "</form>" & vbcrlf

set rstemp=nothing
call closedb()

'*** Fixed FireFox issues
Dim tmpData,tmpData1
Dim tmp1,tmp2,i,Count1
tmpData=Server.HTMLEncode(HTMLResult)
Count1=0
tmpData1=""
tmp1=split(tmpData,"&lt;/tr&gt;")
For i=lbound(tmp1) to ubound(tmp1)
	if i>lbound(tmp1) then
		tmp2="&lt;/tr&gt;" & tmp1(i)
	else
		tmp2=tmp1(i)
	end if
	Count1=Count1+1
	tmpData1=tmpData1 & "<data" & Count1 & ">" & tmp2 & "</data" & Count1 & ">" & vbcrlf
Next
%><note>
<data0><%=Count1%></data0>
<%=tmpData1%>
</note>