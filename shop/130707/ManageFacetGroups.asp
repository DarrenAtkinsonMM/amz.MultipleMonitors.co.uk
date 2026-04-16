<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="ProductCart Search" %>
<% section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
    <tr>
    	<td colspan="2">
        	<!--#include file="pcv4_showMessage.asp"-->        	
        </td>
	</tr>
    <% If scSearch_IsEnabled <> True Then %>
        <tr>
            <td colspan="2">
                <div class="bs-callout bs-callout-info">
                    <h4>Overview</h4>    
                    <p>
                        This app integrates ProductCart with the powerful Apache Solr search platform. ProductCart Search can be used as a replacement for the default search, but comes with extra features and better performance.
                    </p>
                    <p>
                    <a target="_blank" href="#" class="btn btn-info btn-xs">Learn More</a>                      
                    </p>
                </div>
            </td>
        </tr> 
    <% Else %>
        <tr>
            <td colspan="2">
                <div class="cpOtherLinks"><a href="AddEditFG.asp">Add New Facet Group</a><%if HaveRC=1 then%> | <a href="AddEditFC.asp">Add New Facet</a><%end if%></div>      	
            </td>
        </tr>  
    <% End If %>             

    <%
    If scSearch_IsEnabled = True Then
    
        query="SELECT pcFG_ID,pcFG_Name FROM pcFacetGroups ORDER BY pcFG_Name ASC"
        set rs=Server.CreateObject("ADODB.Recordset")
        set rs=connTemp.execute(query)
        HaveRC=0
        if rs.EOF then
            set rs=nothing
            HaveRC=0
        else
            HaveRC=1
        end if
     
        If HaveRC=0 then
        %> 
             
          <tr> 
            <td colspan="2"><div class="pcCPmessage">No facet groups found</div></td>
          </tr>
          <tr>
            <td colspan="2" class="pcCPspacer"></td>
          </tr> 
                         
        <% 
        Else 
            tmpArr=rs.getRows()
            set rs=nothing
            intCount=ubound(tmpArr,2)
            For i=0 to intCount
            tmpFGID=tmpArr(0,i)
            tmpFGName=tmpArr(1,i)%>
            <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                <td width="60%"><a href="AddEditFG.asp?id=<%=tmpFGID%>"><%=tmpFGName%></a></td>
                <td width="40%" nowrap class="cpLinksList">
                    <%
                    Mapped=0
                    queryQ="SELECT pcFG_ID FROM pcFGOG WHERE pcFG_ID=" & tmpFGID & " AND idOptionGroup>0;"
                    set rsQ=connTemp.execute(queryQ)
                    if not rsQ.eof then
                        Mapped=1
                    else
                        Mapped=0
                    end if
                    set rsQ=nothing
                    %>
                    <a href="AddEditFG.asp?id=<%=tmpFGID%>">Edit</a> | <a href="manageFacets.asp?id=<%=tmpFGID%>">Manage Facets</a> | <a href="javascript:<%if Mapped=1 then%>if (confirm('This Facet Group was linked to a Product Option Group. Are you sure you want to complete this action?')) location='delFG.asp?id=<%=tmpFGID%>';<%else%>if (confirm('You are about to remove this facet group from your database. Are you sure you want to complete this action?')) location='delFG.asp?id=<%=tmpFGID%>';<%end if%>">Delete Group</a>
                </td>
            </tr>
            <%Next
        End If
        set rs=nothing
        
    End IF
    %>     
</table>
<br /><br />
<!--#include file="AdminFooter.asp"-->