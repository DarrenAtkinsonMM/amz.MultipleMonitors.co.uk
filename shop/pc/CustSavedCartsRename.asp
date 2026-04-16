<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="CustLIv.asp"-->
<% dim pcv_SavedCartName, tmpID

	tmpID=getUserInput(request("id"),10)
	if tmpID="" or IsNull(tmpID) then
		tmpID=0
	end if
	if not IsNumeric(tmpID) then
		tmpID=0
	end if
	if tmpID=0 then response.Redirect "CustSavedCarts.asp"
	

	if request("submit")<>"" then
		pcv_SavedCartName=getUserInput(request("SavedCartName"),250)
		pcv_SavedCartName=pcf_ReplaceQuotes(pcv_SavedCartName)
		
		set rs=Server.CreateObject("ADODB.Recordset")
		query="UPDATE pcSavedCarts SET SavedCartName=N'" & pcv_SavedCartName & "' WHERE SavedCartID=" & tmpID & ";"
		set rs=conntemp.execute(query)
		set rs=nothing
		
		response.Redirect("CustSavedCarts.asp")
		response.End()
	end if
	
	
	query="SELECT SavedCartName FROM pcSavedCarts WHERE SavedCartID=" & tmpID & ";"
	set rs=connTemp.execute(query)
	If rs.eof then 
		response.Redirect "CustSavedCarts.asp"
	else
		pcv_SavedCartName=rs("SavedCartName")
	End if
	set rs=nothing
%>
<!--#include file="header_wrapper.asp"-->
<div id="pcMain">
	<form action="CustSavedCartsRename.asp" method="post" class="form" role="form">
  <input type="hidden" value="<%=tmpID%>" name="id">
	<div class="pcMainContent">
		<h1><%= dictLanguage.Item(Session("language")&"_CustPref_16")%></h1>
		<div class="pcShowContent">
        
          <div class="form-group">
            <label for="SavedCartName"><%=dictLanguage.Item(Session("language")&"_CustSavedCarts_8") %>:</label>
            <input type="text" value="<%=pcv_SavedCartName%>" name="SavedCartName" class="form-control">
          </div>

		</div>
    
    <div class="pcFormButtons">
      <button class="pcButton pcButtonRename" id="submit" name="submit" value="save">
        <%= dictLanguage.Item(Session("language")&"_CustSavedCarts_9") %>
      </button>

      <a href="CustSavedCarts.asp" class="pcButton pcButtonBack">
        <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_msg_back") %>" />
      	<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_msg_back")%></span>
      </a>
    </div>
	</div>
    </form>
</div>
<!--#include file="footer_wrapper.asp"-->
