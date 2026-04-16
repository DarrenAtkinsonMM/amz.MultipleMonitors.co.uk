
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="opc_pageLoad.asp"-->

<div class="modal-header">
  	<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
  	<h3 class="modal-title" id="pcDialogTitle"><%=dictLanguage.Item(Session("language")&"_pcAVTitle") %></h3>
</div>
<div class="modal-body">
	<div class="pcMainContent">
		<div class="pcShowContent">
			<div id="error"></div>
            
            <!--
            <%
            '// If any error display user-friendly generic message with orginal address.
            If Session("validAddress").Item("ErrorDesc") <> "" Then  
                %>              
                <div class="alert alert-warning"><%=dictLanguage.Item(Session("language")&"_pcAVTitle") %></div>
                <%
            End If
            %>
            <%
            '// If any return text display user-friendly generic message.
            If Session("validAddress").Item("ReturnText") <> "" Then  
                %>              
                <div class="alert alert-warning"><%=Session("validAddress").Item("ReturnText") %>&nbsp;<strong>Please review the address carefully before you confirm it.</strong></div>
                <%
            End If
            %>
            -->
            
            <div class="alert alert-warning"><%=dictLanguage.Item(Session("language")&"_pcAVInfo") %></div>
            <div class="well">
                <%=Session("origAddress").Item("Address") %> <br />                
                <% If Session("origAddress").Item("Address2") <> "" Then %>
                    <%=Session("origAddress").Item("Address2") %> <br />
                <% End If %>                
                <%=Session("origAddress").Item("City") %>&nbsp;
                <%=Session("origAddress").Item("Region") %>&nbsp;
                <%=Session("origAddress").Item("PostalCode") %>                
            </div>

            <button class="btn btn-default" onclick="confirmAddress('<%=Session("validAddress").Item("Type")%>')" type="button"><%=dictLanguage.Item(Session("language")&"_pcAVBtnConfirm") %></button>
            <button class="btn btn-default" data-dismiss="modal" type="button"><%=dictLanguage.Item(Session("language")&"_pcAVBtnEdit") %></button>
            
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
                <button class="btn btn-primary" onclick="acceptAddress('<%=Session("validAddress").Item("Type")%>')" type="button"><%=dictLanguage.Item(Session("language")&"_pcAVBtnAccept") %></button>
            
            <% End If %>
            
      	</div>
  	</div>
  	<div class="pcClear"></div>
</div>


<script>

	function confirmAddress(Type) {
        // Click "Confirm address" Button:
        var doc = angular.element($pc("#OrderPreviewCtrl")).scope(); 
        if (Type=='B') {
            $pc("#billingAddressToken").val('<%=Session("origAddress").Item("Token")%>');
            $pc("#IsBillingAddressValidated").val('true');
            doc.updateBilling(); // Re-run the Billing Panel's "Continue" button exactly.
            $pc("#IsBillingAddressValidated").val('false');
        } else {
            $pc("#shippingAddressToken").val('<%=Session("origAddress").Item("Token")%>');
            $pc("#IsShippingAddressValidated").val('true');
  		    doc.updateShipping() // Re-run the Shipping Panel's "Continue" button exactly.
            $pc("#IsShippingAddressValidated").val('false');
        }
        $('#QuickViewDialog').modal('hide');          
		return(false);
	}

	function acceptAddress(Type) {

        var doc = angular.element($pc("#OrderPreviewCtrl")).scope();
        if (Type=='B') {
            
            // Update fields with new values from validation service:
            $pc("#billcountry").val('<%=Session("validAddress").Item("Country")%>');
            $pc("#billaddr").val('<%=Session("validAddress").Item("Address")%>');
            $pc("#billaddr2").val('<%=Session("validAddress").Item("Address2")%>');
            $pc("#billcity").val('<%=Session("validAddress").Item("City")%>');
            $pc("#billstate").val('<%=Session("validAddress").Item("Region")%>');
            $pc("#billzip").val('<%=Session("validAddress").Item("PostalCode")%>');
            $pc("#billingAddressToken").val('<%=Session("validAddress").Item("Token")%>');
            $pc("#IsBillingAddressValidated").val('true');
            
            // Click "Confirm address" Button:             
            doc.updateBilling(); // Re-run the Billing Panel's "Continue" button exactly.
            $pc("#IsBillingAddressValidated").val('false'); // Reset


        } else {

            // Update fields with new values from validation service:
            $pc("#shipcountry").val('<%=Session("validAddress").Item("Country")%>');
            $pc("#shipaddr").val('<%=Session("validAddress").Item("Address")%>');
            $pc("#shipaddr2").val('<%=Session("validAddress").Item("Address2")%>');
            $pc("#shipcity").val('<%=Session("validAddress").Item("City")%>');
            $pc("#shipstate").val('<%=Session("validAddress").Item("Region")%>');
            $pc("#shipzip").val('<%=Session("validAddress").Item("PostalCode")%>');
            $pc("#shippingAddressToken").val('<%=Session("validAddress").Item("Token")%>');
            $pc("#IsShippingAddressValidated").val('true');
            
            // Click "Confirm address" Button:
            doc.updateShipping(); // Re-run the Shipping Panel's "Continue" button exactly.
            $pc("#IsShippingAddressValidated").val('false'); // Reset
        
        }
        $('#QuickViewDialog').modal('hide');          
        return(false);
	}
</script>
<%
Set Session("origAddress") = Nothing
Set Session("validAddress") = Nothing
call closeDb()
%>