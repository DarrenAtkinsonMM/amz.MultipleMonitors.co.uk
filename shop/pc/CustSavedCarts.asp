<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="CustLIv.asp"-->
<% 
IF request("action")="del" THEN
	tmpID=getUserInput(request("id"),0)
	if tmpID="" or IsNull(tmpID) then
		tmpID=0
	end if
	if not IsNumeric(tmpID) then
		tmpID=0
	end if
	IF tmpID>0 THEN
		query="SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartID=" & tmpID & " AND IDCustomer=" & session("IDCustomer") & ";"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
		  query="DELETE FROM pcSavedCartArray WHERE SavedCartID=" & tmpID & ";"
		  set rsQ=connTemp.execute(query)
		  set rsQ=nothing
		  query="DELETE FROM pcSavedCarts WHERE SavedCartID=" & tmpID & ";"
		  set rsQ=connTemp.execute(query)
		  set rsQ=nothing
		end if
		set rsQ=nothing
	END IF
ELSE
	IF request("action")="res" THEN
		tmpID=getUserInput(request("id"),0)
		if tmpID="" or IsNull(tmpID) then
			tmpID=0
		end if
		if not IsNumeric(tmpID) then
			tmpID=0
		end if
		IF tmpID>0 THEN
			query="SELECT SavedCartGUID FROM pcSavedCarts WHERE SavedCartID=" & tmpID & " AND IDCustomer=" & session("IDCustomer") & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				Response.Cookies("SavedCartGUID")=rsQ("SavedCartGUID")
				set rsQ=nothing        
				Response.Cookies("SavedCartGUID").Expires=Date()+365
				dim pcCartArray(100,45)
				session("pcCartSession")=pcCartArray
				Session("pcCartIndex")=0
				HaveToRestore="yes"
				%>
				<!--#include file="inc_RestoreShoppingCart.asp"-->
				<%
				call closedb()
				if Session("pcCartIndex")=0 then
					response.redirect "msg.asp?message=316"
				else
					response.redirect "viewcart.asp"
				end if
			end if
			set rsQ=nothing
		END IF
	END IF
	If Request("SaveCart")="0" Then
		response.redirect "viewcart.asp?SaveCart=0"
	End If
END IF
%>
<!--#include file="header_wrapper.asp"-->
<div id="pcMain">
  <div class="pcMainContent">
    
    <h1><%= dictLanguage.Item(Session("language")&"_CustPref_50")%></h1>
    
    <%      
      query="SELECT SavedCartID,SavedCartDate,SavedCartName FROM pcSavedCarts WHERE IDCustomer=" & session("IDCustomer") & " ORDER BY SavedCartID DESC;"
      set rsQ=Server.CreateObject("ADODB.Recordset")
      set rsQ=connTemp.execute(query)
      If rsQ.eof then 
        %>
        
        <div class="pcErrorMessage"><%= dictLanguage.Item(Session("language")&"_CustSavedCarts_1")%></div>
      
      <% Else %>
      
       
          <div class="pcCartLayout container-fluid">
            <div class="row pcTableHeader">
              <div class="col-xs-3 col-sm-2"><%= dictLanguage.Item(Session("language")&"_CustSavedCarts_2")%></div>
              <div class="col-xs-5 col-sm-8"><%= dictLanguage.Item(Session("language")&"_CustSavedCarts_8")%></div>
              <div class="col-xs-4 col-sm-2">&nbsp;</div>
            </div>
                    
            <%
              pcArr=rsQ.getRows()
              intCount=ubound(pcArr,2)
              For i=0 to intCount
                Rev_Date = pcArr(1,i)
                If scDateFrmt="DD/MM/YY" then 
                  Rev_Date = day(Rev_Date) & "/" & month(Rev_Date) & "/" & year(Rev_Date)
                Else
                  Rev_Date = month(Rev_Date) & "/" & day(Rev_Date) & "/" & year(Rev_Date)
                End If
                
                %>
                <div class="row">
                    <div class="col-xs-3 col-sm-2"><%=Rev_Date%></div>
                    <div class="col-xs-5 col-sm-8">
                        <%
                        pcv_orderName = pcArr(2,i)
                        If session("Mobile")="1" Then
                            If len(pcv_orderName)>15 Then
                                pcv_orderName = Left(pcv_orderName, 15)
                            End If  
                        End If
                        response.Write(pcv_orderName)
                        %>
                    </div>
                    <div class="col-xs-4 col-sm-2">
                    
                        <div class="btn-group">
                          <button type="button" class="btn btn-default dropdown-toggle" data-toggle="dropdown">
                            <%= dictLanguage.Item(Session("language")&"_CustviewPast_10")%> <span class="caret"></span>
                          </button>
                          <ul class="dropdown-menu pull-right" role="menu">
                            <li>
                                <a href="CustSavedCarts.asp?action=res&amp;id=<%=pcArr(0,i)%>"><span class="glyphicon glyphicon-shopping-cart"></span>&nbsp;<%= dictLanguage.Item(Session("language")&"_CustSavedCarts_3")%></a>                        
                            </li>
                            <li>
                                <a href="CustSavedCartsRename.asp?id=<%=pcArr(0,i)%>"><span class="glyphicon glyphicon-pencil"></span>&nbsp;<%= dictLanguage.Item(Session("language")&"_CustSavedCarts_9")%></a>
                            </li>
                            <li>
                                <a href="javascript:if (confirm('<%= dictLanguage.Item(Session("language")&"_CustSavedCarts_10")%>')) location='CustSavedCarts.asp?action=del&amp;id=<%=pcArr(0,i)%>';">
                                    <span class="glyphicon glyphicon-trash"></span>&nbsp;<%= dictLanguage.Item(Session("language")&"_CustSavedCarts_4")%></a>
                            </li>
                          </ul>
                        </div>

                    </div>
                </div>
                <div class="row">
                <div class="col-xs-12">
                <hr>
                </div>
                </div>
                <%
              Next
            %>


        <div class="pcFormButtons">
        	<a class="pcButton pcButtonBack" href="CustPref.asp">
                <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
            </a>          
        	<a class="pcButton pcButtonViewCart" href="viewCart.asp">
                <img src="<%=pcf_getImagePath("",rslayout("viewcartbtn"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_viewcartbtn") %>" />
                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_viewcartbtn") %></span>
            </a>
        </div>
        
        <div class="pcClear"></div>
        </div>
        
        <%
        end if
        set rsQ=nothing
        %>

    </div>
</div>
<!--#include file="footer_wrapper.asp"-->
