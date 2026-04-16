<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="CustLIv.asp"-->
<!--#include file="header_wrapper.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<% 
'// Check if store is turned off and return message to customer
%> 
<div id="pcMain">   
  <div class="pcMainContent">
  	<h1><%= dictLanguage.Item(Session("language")&"_CustSAmanage_1")%></h1>
    
    <a href="CustAddShip.asp"><%=dictLanguage.Item(Session("language")&"_CustSAmanage_5")%></a>
    
		<%if request("msg")<>"" then%>
      <div class="pcSuccessMessage">
      <%if request("msg")="1" then
          response.write dictLanguage.Item(Session("language")&"_CustSAmanage_7")
        elseif request("msg")="2" then
          response.write dictLanguage.Item(Session("language")&"_CustSAmanage_8")
        elseif request("msg")="3" then
          response.write dictLanguage.Item(Session("language")&"_CustSAmanage_9")
        end if %>
      </div>
    <%end if%>
    
    <div class="pcShowContent">
    	<div class="pcTable">
				<% 
          query="SELECT address, city, state, stateCode, shippingaddress, shippingcity, shippingState, shippingStateCode FROM customers WHERE (((idcustomer)="&session("idCustomer")&"));"
  
          set rs=server.CreateObject("ADODB.RecordSet")
          set rs=conntemp.execute(query)
          if err.number<>0 then
            call LogErrorToDatabase()
            set rs=nothing
            call closedb()
            response.redirect "techErr.asp?err="&pcStrCustRefID
          end if
          
          pcDefaultAddress=rs("address")
          pcDefaultCity=rs("city")
          pcDefaultState=rs("state")
          pcDefaultStateCode=rs("stateCode")
          pcStrDefaultShipAddress=rs("shippingAddress")
          If len(pcStrDefaultShipAddress)<1 then
            pcStrDefaultShipAddress=pcDefaultAddress
            pcStrDefaultShipCity=pcDefaultCity
            pcStrDefaultShipState=pcDefaultState
            pcStrDefaultShipStateCode=pcDefaultStateCode
          Else
            pcStrDefaultShipCity=rs("shippingCity")
            pcStrDefaultShipState=rs("shippingState")
            pcStrDefaultShipStateCode=rs("shippingStateCode") 
          End if
          pcStrDefaultShipState=pcStrDefaultShipState & pcStrDefaultShipStateCode
          set rs=nothing
          %>
          
          <div class="pcTableRowFull"><hr></div>
          
          <div class="pcTableRow">
          	<div class="pcCustSA_Name"><%=dictLanguage.Item(Session("language")&"_CustSAmanage_10")%></div>
            <div class="pcCustSA_Edit"><a href="CustModShip.asp?reID=0"><%=dictLanguage.Item(Session("language")&"_CustSAmanage_3")%></a></div>
          </div>
          
				<% 
					query="SELECT idRecipient, recipient_NickName, recipient_FullName, recipient_Address, recipient_City, recipient_State, recipient_StateCode FROM recipients WHERE (((idCustomer)="&session("idCustomer")&"));"
					set rs = Server.CreateObject("ADODB.Recordset")
					set rs=conntemp.execute(query)
					
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
					If rs.eof then
						intShipAddressExist=0
					end if
					
					do while not rs.eof
						intShipAddressExist=1
						IDre=rs("idRecipient")
						reNickName=trim(rs("recipient_NickName"))
						reFullName=trim(rs("recipient_FullName"))
						reShipAddr=ucase(rs("recipient_Address"))
						reShipCity=ucase(rs("recipient_City"))
						reShipState=ucase(rs("recipient_State") & rs("recipient_StateCode"))        	
	
						if len(reNickName)<1 then
							reNickName=dictLanguage.Item(Session("language")&"_CustSAmanage_12")
						end if %>
            
            <div class="pcTableRow pcCustSA_Row">
              <div class="pcCustSA_Name"><%=reNickName%></div>
              <div class="pcCustSA_Edit">
                <a href="CustModShip.asp?reID=<%=IDre%>"><%=dictLanguage.Item(Session("language")&"_CustSAmanage_3")%></a> | <a href="javascript:if (confirm('<%=dictLanguage.Item(Session("language")&"_CustSAmanage_11")%>')) location='CustDelShip.asp?reID=<%=IDre%>'"><%=dictLanguage.Item(Session("language")&"_CustSAmanage_4")%></a>
              </div>
            </div>

					<% 					
						rs.movenext
					loop
					set rs = nothing
				%>	
      	<div class="pcTableRowFull"><hr></div>
      </div>
      
      <div class="pcFormButtons">
        <a class="pcButton pcButtonBack" href="custPref.asp">
          <img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="Back">
          <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
        </a>
      </div>
    </div>
  </div>
</div>
<!--#include file="footer_wrapper.asp"-->
