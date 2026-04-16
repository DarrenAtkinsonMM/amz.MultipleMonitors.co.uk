<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="header_wrapper.asp"-->
<% 
Dim pshipmentDetails

pcv_IdOrder=request("idorder")
if pcv_IdOrder="" then
	pcv_IdOrder=0
end if
pcv_PrdList=""
pcv_count=request("count")
if pcv_count="" then
	pcv_count=0
end if

if (pcv_IdOrder=0) or (pcv_count=0) then
	response.redirect "default.asp"
end if

For i=1 to pcv_count
	if request("C" & i)="1" then
		pcv_PrdList=pcv_PrdList & request("IDPrd" & i) & ","
	end if
Next
	%>
<div id="pcMain">
	<div class="pcMainContent">
    <h1><%= dictLanguage.Item(Session("language")&"_sds_viewpast_1c")%> - <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_1")%> <%=(scpre+int(pcv_IdOrder))%></h1>
      
    <ul class="pcShipWizardHeader">
      <li class="pcShipWizardStep1">
        <img src="<%=pcf_getImagePath("images","step1.gif")%>">
        <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_3")%>
      </li>
      <li class="pcShipWizardStep2 active">
        <img src="<%=pcf_getImagePath("images","step2a.gif")%>">
        <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_4")%>
      </li>
      <li class="pcShipWizardStep3">
        <img src="<%=pcf_getImagePath("images","step3.gif")%>">
        <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_5")%>
      </li>
    </ul>
    
    <div class="pcClear"></div>

    <form name="form1" method="post" action="sds_ShipOrderWizard3.asp?action=add" class="pcForms">
      <%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_14")%>
  
      <div class="pcSpacer"></div>
          
      <%
      ' START Shipment type
      query="SELECT shipmentDetails FROM orders WHERE idOrder = " & pcv_IdOrder
      set rs=Server.CreateObject("ADODB.Recordset")
      set rs=conntemp.execute(query)
      if not rs.EOF then
        pshipmentDetails=rs("shipmentDetails")
      end if
      set rs=nothing
      
      if pshipmentDetails<>"" and not isNull(pshipmentDetails) then
      %>
				<%= dictLanguage.Item(Session("language")&"_sds_custviewpastD_16")%>
			<% 
      Service=""
      If pSRF="1" then
      	response.write ship_dictLanguage.Item(Session("language")&"_noShip_b")
      else
          'get shipping details...
          shipping=split(pshipmentDetails,",")
          if ubound(shipping)>1 then
              if NOT isNumeric(trim(shipping(2))) then
                  varShip="0"
                  response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
              else
                  Service=shipping(1)
              end if
              if len(Service)>0 then
                  response.write Service
              End If
          else
              varShip="0"
              response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
          end if
				end if
				%>
				<%
					if pOrdShipType=0 then
							pDisShipType=dictLanguage.Item(Session("language")&"_sds_custviewpastD_18")
					else
							pDisShipType=dictLanguage.Item(Session("language")&"_sds_custviewpastD_19")
					end if
					if varShip<>"0" then
				%>
					<%=dictLanguage.Item(Session("language")&"_sds_custviewpastD_17")%><%=pDisShipType%>
				<%
					end if
				end if
				' END Shipment Type
			%>
  
  
      <div class="pcSpacer"></div>
      
      <hr>
      
      <div class="pcShowContent">
      	<% 'Shipment Method %>
      	<div class="pcFormItem">
        	<div class="pcFormLabel">
						<%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_15")%>
          </div>
          <div class="pcFormField">
          	<input type="text" name="pcv_method" value="" size="30">
          </div>
        </div>
        
      	<% 'Tracking Number %>
      	<div class="pcFormItem">
        	<div class="pcFormLabel">
						<%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_16")%>
          </div>
          <div class="pcFormField">
          	<input type="text" name="pcv_tracking" value="" size="30">
          </div>
        </div>
        
				<%
          Dim varMonth, varDay, varYear
          varMonth=Month(Date)
          varDay=Day(Date)
          varYear=Year(Date) 
          dim dtInputStr
          dtInputStr=(varMonth&"/"&varDay&"/"&varYear)
          if scDateFrmt="DD/MM/YY" then
              dtInputStr=(varDay&"/"&varMonth&"/"&varYear)
          end if
        %>
    
      	<% 'Shipped Date %>
      	<div class="pcFormItem">
        	<div class="pcFormLabel">
						<%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_17")%>
          </div>
          <div class="pcFormField">
          	<input type="text" name="pcv_shippedDate" value="<%=dtInputStr%>" size="30"> <i><%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_17a")%></i>
          </div>
        </div>
        
      	<% 'Comments %>
      	<div class="pcFormItem">
        	<div class="pcFormLabel">
						<%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_18")%>
          </div>
          <div class="pcFormField">
						<textarea name="pcv_AdmComments" size="40" rows="8" cols="40"></textarea>          
          </div>
        </div>
        
        <hr>
        
        <div class="pcFormItem">
        	<div class="pcFormLabel">
          </div>
          <div class="pcFormField">
            <div class="pcFormButtons">
            	<button class="pcButton pcButtonFinalizeShip" name="submit1" id="submit">
              	<img src="<%=pcf_getImagePath("",rslayout("pcLO_finalShip"))%>" alt="<%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_19") %>" />
                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_sds_shiporderwizard_19") %></span>
              </button>
              
              <a class="pcButton pcButtonBack" href="javascript:history.go(-1);">
              	<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
              </a>
              
              <input type="hidden" name="PrdList" value="<%=pcv_PrdList%>">
              <input type="hidden" name="idorder" value="<%=pcv_IdOrder%>">
              <input type="hidden" name="count" value="<%=pcv_count%>">
            </div>
          </div>
        </div>
    	</div>
    </Form>
    
    <div class="pcSpacer"></div>

  	<div style="text-align: center"><a href="sds_MainMenu.asp"><%response.write(dictLanguage.Item(Session("language")&"_CustPref_1"))%></a> - <a href="sds_ViewPast.asp"><%response.write(dictLanguage.Item(Session("language")&"_sdsMain_3"))%></a></div>

	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
