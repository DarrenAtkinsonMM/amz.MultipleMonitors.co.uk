<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
dim pcIntParent, pcvParentBrandName

pcIntParent=request("parent")

If Not validNum(pcIntParent) Then
    pcIntParent = 0
End If
	
If pcIntParent>0 Then
    ' Load Parent Brand Name		
    query="SELECT BrandName FROM Brands WHERE idBrand="&pcIntParent
    set rs=server.CreateObject("ADODB.RecordSet")
    set rs=conntemp.execute(query)
    pcvParentBrandName=rs("BrandName")
    pageTitle="Manage Brands under " & pcvParentBrandName
    set rs=nothing    
Else
    pcvParentBrandName=""
    pageTitle="Manage Brands"
End IF

dim i, idBrand, priority

sMode=request.form("submitForm")
If sMode<>"" Then
	
	pcIntParent=request.form("parent")
	iCnt=request.form("iCnt")
	set rs=server.CreateObject("ADODB.RecordSet")
	for i=1 to iCnt
		idBrand=request.form("idBrand"&i)
		priority=request.form("priority"&i)
		query="UPDATE Brands SET "
		query=query & "pcBrands_Order=" &priority
		query=query & " WHERE idBrand="&idBrand&";"
		on error resume next
		rs.open query,conntemp
	next
	set rs=nothing
		
	call closeDb()
    response.redirect "BrandsManage.asp?parent="&pcIntParent
	response.end
End If


'// Load Brands
query = "SELECT A.idBrand, A.BrandName, A.pcBrands_Order, A.BrandLogo, "
query = query & "( "
query = query & "	SELECT Count(*) As totalSubBrands "
query = query & "	FROM Brands B "
query = query & "	WHERE B.pcBrands_Parent = A.IDBrand "
query = query & ") AS SubQty "
query = query & ",( "
query = query & "	SELECT Count(*) As totalProducts "
query = query & "	FROM products C "
query = query & "	WHERE active=-1 AND configOnly=0 AND removed=0 AND C.IDBrand = A.IDBrand "
query = query & ") AS PrdQty "
query = query & "FROM Brands A "
query = query & "WHERE A.pcBrands_Parent = " & pcIntParent & " "
query = query & "ORDER BY A.pcBrands_Order, A.BrandName ASC "

Set rs=Server.CreateObject("ADODB.Recordset")
Set rs=connTemp.execute(query)
if err.number <> 0 then
	set rs=nothing
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error loading brands") 
end If

Dim iCnt, pcIntNoResults
iCnt=0
pcIntNoResults=0

If rs.eof Then
	pcIntNoResults=1
End if
%>
<!--#include file="AdminHeader.asp"-->

<form name="form1" method="post" action="BrandsManage.asp" class="pcForms">
    <input type="hidden" name="parent" value="<%=pcIntParent%>">
    <table class="pcCPcontent">
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
        <tr>
            <td colspan="3">
                <%
                if pcvParentBrandName="" then
                %>
                These are the top level brands. They are displayed in the storefront on the <a href="../pc/viewbrands.asp" target="_blank">View Brands</a> page, based on the settings entered on the <a href="AdminSettings.asp?tab=4#brandSettings">Display Settings</a> page. 
<%
                else
                %>
                These are the brands under <strong><a href="BrandsEdit.asp?idbrand=<%=pcIntParent%>"><%=pcvParentBrandName%></a></strong>. Back to the <a href="BrandsManage.asp">Top Level Brands</a>.
<%
                end if
                %>
            </td>
        </tr>
  </table>

    <div id="pcCPbrandList">
        <div class="pcCPsortableTableHeader">
              <div class="pcCPsortableTableIndex">#</div>
              <div class="pcCPbrandName">Brand</div>
              <div class="pcCPbrandImage">Logo</div>
        </div>

        <% if pcIntNoResults=1 then %>
            <div class="pcCPmessage">No brands found.</div>
        <% else %>
            <ul class="pcCPsortable pcCPsortableTable">

                <%
                pcArr=rs.getRows()
                set rs=nothing
                intCount=ubound(pcArr,2)
              
                For i=0 to intCount
                    pcvBrandName=pcArr(1,i)
                    tidBrand=pcArr(0,i)
                    tpriority=pcArr(2,i)
                    timage=pcArr(3,i)
                    iBrandsCount=pcArr(4,i)
                    if IsNull(iBrandsCount) OR iBrandsCount="" then
                        iBrandsCount=0
                    end if
                    iProductCount=pcArr(5,i)
                    if IsNull(iProductCount) OR iProductCount="" then
                        iProductCount=0
                    end if
                    iCnt=iCnt+1
                    %>
                    <li class="cpItemlist"> 
                        <div class="pcCPsortableTableIndex"> 
                            <span class="pcCPsortableIndex"><%= iCnt %></span>
                            <input type="hidden" name="priority<%=iCnt%>" class="pcCPsortableOrder" value="<%=tpriority%>">
                            <input type="hidden" name="idBrand<%=iCnt%>" value="<%=tidBrand%>">
                        </div>
                        <div class="pcCPbrandName"> 
                            <a href="BrandsEdit.asp?idbrand=<%=tidBrand%>"><%=pcvBrandName%></a>
                        </div>
                        <div class="pcCPbrandImage">
                          <% If Len(timage) > 0 Then %>
                            <img src="../pc/catalog/<%= timage %>" alt="<%= pcvBrandName %>" />
                          <% End If %>
                        </div>
        
                        <div class="pcCPbrandLinks cpLinksList">
                            <a href="BrandsEdit.asp?idbrand=<%=tidBrand%>">Edit</a> | <a href="BrandsManage.asp?parent=<%=tidBrand%>">Sub-Brands (<%=iBrandsCount%>)</a> | <a href="BrandsProducts.asp?idbrand=<%=tidBrand%>">Products (<%=iProductCount%>)</a> | <a href="javascript:if (confirm('You are about to permanently remove this Brand from the database. This action CANNOT be undone. You might want to consider making the Brand \'Inactive\' instead (this setting is on the Edit Brand page). Click OK to confirm the removal or CANCEL to keep the current data.')) location='BrandsDel.asp?idbrand=<%=tidBrand%>';">Delete</a>
                        </div>
                    </li>
                    <%
                    Next
                    %>
                </ul>
                <%
                end if
                set rs=nothing
                %>
                <input type="hidden" name="iCnt" value="<%=iCnt%>">
    </div>

    <table class="pcCPcontent">
        <tr> 
            <td colspan="3" style="padding-top: 20px;"> 
            <% if pcIntNoResults=0 then %>
            <input name="submitForm" type="submit" class="btn btn-primary" value="Update Brands Order">&nbsp;
            <% end if %>
            <input name="addNew" type="button" class="btn btn-default"  value="Add New" onClick="document.location.href='BrandsAdd.asp'">
            <input name="back" type="button" class="btn btn-default"  value="Back" onClick="javascript:history.go(-1);">
            </td>
        </tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->
