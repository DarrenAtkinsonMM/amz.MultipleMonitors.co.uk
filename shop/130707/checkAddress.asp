<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="../includes/pcUSPSClass.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/validation.asp" -->

<%
Set pcAddress = Server.CreateObject("Scripting.Dictionary")
pcAddress.Add "Address", Request("address")
pcAddress.Add "Address2", Request("address2")
pcAddress.Add "Country", Request("country")
pcAddress.Add "City", Request("city")
pcAddress.Add "Region", Request("state")
pcAddress.Add "PostalCode", Request("zip")
pcAddress.Add "Type", Request("type")

Set Session("origAddress") = pcAddress
If USPS_AddressValidation = "1" then
	Dim USPS_postdata, USPS_result, srvUSPSXmlHttp, objOutputXMLDoc
	call USPS_validateAddress(pcAddress)
elseif ptaxAvalaraAddressValidation = 1 then
	call Avalara_validateAddress(pcAddress)
end if
%>

<html>
<head>
	<title>Validate Address</title>
    <link href="../pc/css/bootstrap.min.css" rel="stylesheet" type="text/css">
    <!--#include file="inc_jquery.asp" -->
    <script src="../includes/javascripts/bootstrap.min.js"></script>
</head>
<body style="margin:10px">
<div class="table-responsive">
    <div class="modal-header">
        <h3 class="modal-title" id="pcDialogTitle">Validate Address</h3>
    </div>
    <div class="modal-body">
        <div class="pcMainContent">
            <div class="pcShowContent">
                <div id="error"></div>
            <%
            '// If any error display user-friendly generic message with orginal address.
            If Session("validAddress").Item("ErrorDesc") <> "" Then
                %>              
                <div class="alert alert-warning"><%=Session("validAddress").Item("ErrorDesc")%></div>
                <%
            End If
            %>
            <%
            '// If any return text display user-friendly generic message.
            If Session("validAddress").Item("ReturnText") <> "" Then  
                %>              
                <div class="alert alert-warning"><%=Session("validAddress").Item("ReturnText") %></div>
                <%
            End If
            %>
            
            <% If Session("origAddress").Item("Address") <> "" Then %>
            <div class="well">
                <%=Session("origAddress").Item("Address") %> <br />                
                <% If Session("origAddress").Item("Address2") <> "" Then %>
                    <%=Session("origAddress").Item("Address2") %> <br />
                <% End If %>                
                <%=Session("origAddress").Item("City") %>&nbsp;
                <%=Session("origAddress").Item("Region") %>&nbsp;
                <%=Session("origAddress").Item("PostalCode") %>                
            </div>
            <button class="btn btn-default" onclick="window.close()" type="button">Keep using current address</button>
            <% End If %>
            
            <% If (Session("validAddress").Item("ErrorDesc") = "") Then %>
            
                <hr />
                <h4><strong>OR try this suggestion:</strong></h4>
                <div class="well">
                    <strong>
                        <%=Session("validAddress").Item("Address") %> <br />                
                        <% If Session("validAddress").Item("Address2") <> "" Then %>
                            <%=Session("validAddress").Item("Address2") %> <br />
                        <% End If %>                
                        <%=Session("validAddress").Item("City") %>&nbsp;
                        <%=Session("validAddress").Item("Region") %>&nbsp;
                        <%=Session("validAddress").Item("PostalCode") %>  
                    </strong>             
                </div>                
                <button class="btn btn-primary" onclick="acceptAddress('<%=Session("validAddress").Item("Type")%>')" type="button">Accept Suggestion</button>
            
            <% End If %>
            </div>
        </div>
        <div class="pcClear"></div>
    </div>
</div>
<script type="text/javascript">
function acceptAddress(type) {
	
	var parent = opener.$("#modCust");
	
	if (type=='B') {
		
		// Update fields with new values from validation service:
		parent.find("select[name='pcBillingCountryCode'] option[value=<%=Session("validAddress").Item("Country")%>]").prop('selected', true);
		parent.find("input[name='pcBillingAddress']").val('<%=Session("validAddress").Item("Address")%>');
		parent.find("input[name='pcBillingAddress2']").val('<%=Session("validAddress").Item("Address2")%>');
		parent.find("input[name='pcBillingCity']").val('<%=Session("validAddress").Item("City")%>');
		parent.find("select[name='pcBillingStateCode'] option[value=<%=Session("validAddress").Item("Region")%>]").prop('selected', true);
		parent.find("input[name='pcBillingPostalCode']").val('<%=Session("validAddress").Item("PostalCode")%>');

	} else {

		// Update fields with new values from validation service:
		parent.find("select[name='ShipCountryCode'] option[value=<%=Session("validAddress").Item("Country")%>]").prop('selected', true);
		parent.find("input[name='ShipAddress']").val('<%=Session("validAddress").Item("Address")%>');
		parent.find("input[name='ShipAddress2']").val('<%=Session("validAddress").Item("Address2")%>');
		parent.find("input[name='ShipCity']").val('<%=Session("validAddress").Item("City")%>');
		parent.find("select[name='ShipStateCode'] option[value=<%=Session("validAddress").Item("Region")%>]").prop('selected', true);
		parent.find("input[name='ShipZip']").val('<%=Session("validAddress").Item("PostalCode")%>');
	
	}
	
	parent.find("input[name='Modify']").click();
	window.close();
	
}
</script>
<%
set Session("validAddress") = nothing
%>
</body>