<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="../includes/common.asp"--> 
<%
pidProduct=request("idProduct")

pcv_Test=1
if pidProduct="" then
	pcv_Test=0
else
    if not IsNumeric(pidProduct) then
        pcv_Test=0
    end if
end if

If pcv_Test=1 Then

    query = "Select pcprod_SizeLink,pcprod_SizeInfo,pcprod_SizeImg,pcprod_SizeURL from products where idproduct=" & pIDProduct
    Set rstemp=connTemp.execute(query)
    If Not rstemp.Eof Then
        pcv_SizeLink=rstemp("pcprod_SizeLink")
        pcv_SizeInfo=rstemp("pcprod_SizeInfo")
        pcv_SizeImg=rstemp("pcprod_SizeImg")
        pcv_SizeURL=rstemp("pcprod_SizeURL")
    End If
    Set rstemp = Nothing
    
    If (pcv_SizeInfo="") And (pcv_SizeImg="") And (pcv_SizeURL<>"") Then
        response.redirect pcv_SizeURL
    End If
    %>
    <html>
    
    <head>
    <title><%=pcv_SizeLink%></title>
    <link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath("","pcStorefront.css")%>" />
    </head>
    <body>
        <div id="pcMain">
    
            <%if (pcv_SizeInfo<>"") and (pcv_SizeImg<>"") then%>
            
                <p><%=pcv_SizeInfo%></p>
                <p><img src="<%=pcf_getImagePath("catalog",pcv_SizeImg)%>"></p>
            
            <%else%>
            
                <%if (pcv_SizeInfo<>"") then%>
                    <p><%=pcv_SizeInfo%></p>
                <%end if%>
                
                <%if (pcv_SizeImg<>"") then%>
                    <p><img src="<%=pcf_getImagePath("catalog",pcv_SizeImg)%>" border=0></p>
                <%end if%>
            
            <%end if%>
    
        </div>
    </body>
    </html>
    
<%
End If
call closeDb()
%>