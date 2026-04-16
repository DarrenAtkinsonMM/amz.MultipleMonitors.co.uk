<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<% 

%>
<div class="modal-header">
    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
    <h3 class="modal-title" id="pcDialogTitle"><%= dictLanguage.Item(Session("language")&"_catering_15")%></h3>
</div>
<div class="modal-body">

    <div class="pcInfoMessage"><%= dictLanguage.Item(Session("language")&"_catering_16")%></div>
    <div class="pcSpacer"></div> 
    
    <div class="pcTable">
    	<div class="pcTableHeader">
      	<div class="pcBlackoutDates_Date">
        	<%= dictLanguage.Item(Session("language")&"_catering_17")%>
        </div>
        <div class="pcBlackoutDates_Reason">
            <%= dictLanguage.Item(Session("language")&"_catering_18")%>
        </div>
    </div>
     		  
    <%
    query="select * from Blackout order by Blackout_Date asc"
    set rstemp=conntemp.execute(query)
    If rstemp.eof Then
        %>
        <div class="pcTableRowFull">
            <%= dictLanguage.Item(Session("language")&"_catering_19")%>
        </div>                
    <% Else
    
        Dim strCol
        strCol="#E1E1E1"
        Do While NOT rstemp.EOF
            Blackout_Date=rstemp("Blackout_Date")
            Blackout_Message=rstemp("Blackout_Message")
            If strCol <> "#FFFFFF" Then
                strCol="#FFFFFF"
            Else 
                strCol="#E1E1E1"
            End If
            %>          
            <div class="pcTableRow" style="background-color: <%= strCol %>">
                <div class="pcBlackoutDates_Date"><%=Blackout_Date%></div>
                <div class="pcBlackoutDates_Reason"><%=Blackout_Message%></div>
            </div>
                
            <%
            rstemp.MoveNext
        Loop

    End If
    %>

    <div class="pcClear"></div>
</div>
<div class="modal-footer">
    <button class="btn btn-default" data-dismiss="modal" type="button"><%=dictLanguage.Item(Session("language")&"_AddressBook_5")%></button>
</div>
<%
call closeDb()
%>
