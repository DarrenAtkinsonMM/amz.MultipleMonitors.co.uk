<script type=text/javascript>

    /////////////////////////////////////////////
    // START: LANGUAGE
    /////////////////////////////////////////////

    var opc_icon_error = '<img src="<%=pcf_getImagePath("images","pc_icon_error_small.png")%>" align="absmiddle">';
    var opc_icon_loader = '<img src="<%=pcf_getImagePath("images","ajax-loader1.gif")%>" width="20" height="20" align="absmiddle">';    
    var OPCReady="NO";
    var tmpchkFree="";

    /////////////////////////////////////////////
    // END: LANGUAGE
    /////////////////////////////////////////////





    /////////////////////////////////////////////
    // START: GLOBAL STATIC
    /////////////////////////////////////////////
    
    var shoppingCart;
    var CountriesRequireZipCode="US, CA, GB, " 
    var CountriesRequireStateCode="US, CA, AU, "
    var CountriesRequireProvince=""
    var OPCCheck = 1;
    var OPCFinal = 0;
    var OPCFree = 0;
    var AskEnterPass=0;
    var AcceptEnterPass=0;
    var SBSkip = 0;
    var HaveShipTypeAreaBilling=0;
    var HaveShipTypeArea=0;
    
    /////////////////////////////////////////////
    // END: GLOBAL STATIC
    /////////////////////////////////////////////





    /////////////////////////////////////////////
    // START: GLOBAL DYNAMIC
    /////////////////////////////////////////////
    
    <%
    '// Vat Settings
    pcv_ShowVatId = false
    pcv_isVatIdRequired = false
    pcv_ShowSSN = false
    pcv_isSSNRequired = false
    
    if pshowVatID="1" then pcv_ShowVatId = true
    if pVatIdReq="1" then pcv_isVatIdRequired = true
    if pshowSSN="1" then pcv_ShowSSN = true
    if pSSNReq="1" then pcv_isSSNRequired = true
    %>
    
    // SHIPPING CHARGES CONTENT FLAG (Used on the payment panel button click)
    <% If pcv_NOShippingAtAll = "1" Then %>
        var NeedLoadShipChargeContent=1;
    <% Else %>
        var NeedLoadShipChargeContent=0;
    <% End If %>
    
    <%if session("Cust_IDEvent")<>"" then%>
        var HaveGRAddress=1;
        <%if gDelivery=1 then%>
            var GRAddrOnly=1;
        <%else%>
            var GRAddrOnly=0;
        <%end if%>
    <%else%>
        var HaveGRAddress=0;
        var GRAddrOnly=0;
    <%end if%>

    // Phone is always required in v5
    var IsBillingPhoneReq=true;
	
	<% If SPhoneReq = 1 Then %>
		var IsShippingPhoneReq=true;
	<% Else %>
		var IsShippingPhoneReq=false;
	<% End If %>
    
    <% If pcv_isSSNRequired = true Then %>
        var IsSSNRequired=true;
    <% Else %>
        var IsSSNRequired=false;
    <% End If %>
    
    <% If pcv_isVatIdRequired = true Then %>
        var IsVatIdRequired=true;
    <% Else %>
        var IsVatIdRequired=false;
    <% End If %>
    
    <% If pcv_ShowVatId Then %>
        var ShowVatId=true;
    <% Else %>
        var ShowVatId=false;
    <% End If %>
    
    <% If pcv_ShowSSN Then %>
        var ShowSSN=true;
    <% Else %>
        var ShowSSN=false;
    <% End If %>


    <% If (scGuestCheckoutOpt=0 Or scGuestCheckoutOpt="" Or scGuestCheckoutOpt=1) And (Not (Session("idCustomer")>0 And session("CustomerGuest")="0")) Then %>
        var displayEmail=true;
    <% Else %>
        var displayEmail=false;
    <% End If %> 

    <% If displayShippingAddress Then %>
        var displayShippingAddress = true;
    <% Else %>
        var displayShippingAddress = false;
    <% End If %>

    var ComResShipAddress='<%=pcComResShipAddress%>'; 
        
    <% If session("PPSA")="1" Then %>
            var PPSAID=true;
    <% Else %> 
            var PPSAID=false;   
    <% End If %>
    
    var PPSAIDvalue = '<%=session("PPSA") %>';   
    
    
    <% If pcv_strPayPanel = "1" Then %>
            var goToPaymentPanel=true;
    <% Else %> 
            var goToPaymentPanel=false;   
    <% End If %> 
    
    var scGuestCheckoutOpt = '<%=scGuestCheckoutOpt%>';
    var debuggingEnabled = false;


    /////////////////////////////////////////////
    // END: GLOBAL DYNAMIC
    /////////////////////////////////////////////
    
</script>