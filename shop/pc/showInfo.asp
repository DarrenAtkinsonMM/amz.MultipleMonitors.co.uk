<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<%  
Dim ProductArray, i

categoryDescName = getUserInput(Request.QueryString("cd"),200)
IDBTO = getUserInput(Request.QueryString("IDBTO"),0)
If not IsNumeric(IDBTO) then
	response.redirect "default.asp"
End if

IDCAT = getUserInput(Request.QueryString("IDCAT"),0)
If not IsNumeric(IDCAT) then
	response.redirect "default.asp"
End if

IDPROD = getUserInput(Request.QueryString("IDPROD"),0)
If IDPROD<>"" Then
    If not IsNumeric(IDPROD) then
        response.redirect "default.asp"
    End if
End If

pcv_PageName="More details for " & categoryDescName 
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!-- #include file="inc_headerV5.asp" -->
<style>
    body {
        background: #fff !important;
        margin: 0px
    }
</style>
<script type=text/javascript>
imagename = '';
function enlrge(imgnme) {
	lrgewin = window.open("about:blank","","height=200,width=200")
	imagename = imgnme;
	setTimeout('update()',500)
}
function win(fileName)
	{
	myFloater = window.open('','myWindow','scrollbars=auto,status=no,width=400,height=300')
	myFloater.location.href = fileName;
	}
	function viewWin(file)
	{
	myFloater = window.open('','myWindow','scrollbars=yes,status=no,width=400,height=400')
	myFloater.location.href = file;
	}
function update() {
doc = lrgewin.document;
doc.open('text/html');
doc.write('<HTML><HEAD><TITLE>Enlarged Image<\/TITLE><\/HEAD><BODY bgcolor="white" onLoad="if (document.all || document.layers) window.resizeTo((document.images[0].width + 10),(document.images[0].height + 80))" topmargin="4" leftmargin="0" rightmargin="0" bottommargin="0"><table width=""' + document.images[0].width + '" height="' + document.images[0].height +'"cellspacing="0" cellpadding="0"><tr><td>');
doc.write('<IMG SRC="' + imagename + '"><\/td><\/tr><tr><td><form name="viewn"><A HREF="javascript:window.close()"><img  src="<%=pcf_getImagePath("images","close.gif")%>" align="right" border=0><\/a><\/td><\/tr><\/table>');
doc.write('<\/form><\/BODY><\/HTML>');
doc.close();
}
</script>
</head>
<body id="pcPopup">
<div id="pcMain">
<div class="pcMainContent">
<!--#include file="../includes/javascripts/pcWindowsViewPrd.asp"-->
<%
query="SELECT products.idproduct,products.description, products.smallImageUrl, products.largeImageURL, products.details FROM products INNER JOIN configSpec_products ON products.IDProduct=configSpec_products.configProduct WHERE configSpec_products.configProductCategory="& IDCAT &" AND configSpec_products.specProduct=" & IDBTO & " "
If IDPROD<>"" Then
    query = query & " AND products.idproduct=" & IDPROD 
End If 
set rs=server.CreateObject("ADODB.Recordset")

set rs=conntemp.execute(query)
If NOT rs.eof then
	pcArr=rs.getRows()
	intCount=ubound(pcArr,2)
	set rs = nothing
	For i=0 to intCount
	pcv_productName=pcArr(1,i)
	pcv_strShowImage_Url = pcArr(2,i)
	pcv_strShowImage_LargeUrl = pcArr(3,i)
	If len(pcv_strShowImage_LargeUrl)>0 Then		
		pcv_strLargeUrlPopUp= "javascript:pcAdditionalImages('" & pcf_getImagePath("catalog",pcv_strShowImage_LargeUrl)&"','"&pcArr(0,i)&"')" 
	Else
		pcv_strShowImage_LargeUrl = pcv_strShowImage_Url '// we dont have one, show the regular size
		pcv_strLargeUrlPopUp= "javascript:pcAdditionalImages('" & pcf_getImagePath("catalog",pcv_strShowImage_Url)&"','"&pcArr(0,i)&"')" 
	End If
	pcv_productDetails=pcArr(4,i)
	%>
	<h1><%=pcv_productName%></h1>
		<div class="pcBTOpopup">
			<div class="pcTableRowFull">
				<% if iBTOPopImage=1 then %>
					<%=pcv_productDetails%>
				<% else %>
				<div style="width:65%; float: left;">
					<%=pcv_productDetails%>
				</div>
				<div style="width:32%; float: left;"> 
					<% if pcv_strShowImage_Url<>"" then %>
						<% if pcv_strShowImage_LargeUrl<>"" then %>
							<a href="javascript:enlrge('<%=pcf_getImagePath("catalog",pcv_strShowImage_LargeUrl)%>')">
							<img class="ProductThumbnail" alt="<%=pcv_productName%>" src="<%=pcf_getImagePath("catalog",pcv_strShowImage_Url)%>" style="text-align: right; max-width: 100%">
							</a> 
						<% else %>
							<img src="<%=pcf_getImagePath("catalog",pcv_strShowImage_Url)%>" alt="<%=pcv_productName%>" style="text-align: right; max-width: 100%">
						<% end if %>	
					<% end if %>
				</div>
				<% end if %>
			</div>
		</div>
		<div style="clear:both"></div>
	<%Next	
	end if
set rs = nothing
call closeDB()
%>
	<div align="right" style="clear: both;">
		<A HREF="javascript:window.close()"><img  src="<%=pcf_getImagePath("images","close.gif")%>" border="0"></a>
	</div>
</div>
</div>
</body>
</html>
