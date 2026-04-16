<script src="https://ajax.googleapis.com/ajax/libs/angularjs/1.2.4/angular.min.js"></script>

<% If scPaypalECInContext = "1" then %>
<script src="https://www.paypalobjects.com/api/checkout.js" async></script>
<% End If %>

<% If scEnableBundling = "1" Then %>

    <script src="<%=pcf_getJSPath("js","combined.min.js")%>"></script>

<% Else %>

    <script src="<%=pcf_getJSPath("/shop/includes/jquery","jquery.validate.min.js")%>"></script>
    <script src="<%=pcf_getJSPath("/shop/includes/jquery","jquery.form.js")%>"></script>
    <script src="<%=pcf_getJSPath("/shop/includes/jquery","jquery.touchSwipe.min.js")%>"></script>
    <script src="<%=pcf_getJSPath("/shop/includes/jquery/smoothmenu","ddsmoothmenu.js")%>"></script>
        
    <script src="<%=pcf_getJSPath("/shop/includes/javascripts","bootstrap.js")%>"></script>
    <script src="<%=pcf_getJSPath("/shop/includes/javascripts","bootstrap-tabcollapse.js")%>"></script>
    <script src="<%=pcf_getJSPath("/shop/includes/javascripts","bootstrap-datepicker.js")%>"></script>
    <% if not lcase(Request.ServerVariables("SCRIPT_NAME")) = "/default.asp" then %>
    <script src="<%=pcf_getJSPath("/shop/includes/javascripts","highslide.html.packed.js")%>"></script>
    <% end if %>
    <script src="<%=pcf_getJSPath("/shop/includes/javascripts","jquery.blockUI.js")%>"></script>
    <script src="<%=pcf_getJSPath("/shop/includes/javascripts","json3.js")%>"></script>
    <script src="<%=pcf_getJSPath("/shop/includes/javascripts","accounting.min.js")%>"></script>
    <script src="<%=pcf_getJSPath("/shop/includes/javascripts","productcart.js")%>"></script>
    
    <script src="<%=pcf_getJSPath("/shop/pc/service/app","service.js")%>"></script>
    <script src="<%=pcf_getJSPath("/shop/pc/service/app","quickcart.js")%>"></script>
    <script src="<%=pcf_getJSPath("/shop/pc/service/app","viewcart.js")%>"></script>
    <script src="<%=pcf_getJSPath("/shop/pc/service/app","search.js")%>"></script>
    
    <% If pcStrPageName = "onepagecheckout.asp" Then %>  
        <script src="<%=pcf_getJSPath("/shop/includes/javascripts","opc_validation.js")%>"></script>
        <script src="<%=pcf_getJSPath("/shop/pc/service/app","onepagecheckout.js")%>"></script>
    <% End If %>
    
    <% If pcStrPageName = "OrderComplete.asp" OR pcStrPageName="custviewpastd.asp" Then %>  
        <script src="<%=pcf_getJSPath("/shop/pc/service/app","order.js")%>"></script>
    <% End If %>
    
    <script type="text/javascript" src="<%=pcf_getJSPath("/shop/includes/mojozoom","mojozoom.js")%>"></script>
    <script type="text/javascript" src="<%=pcf_getJSPath("/shop/pc/js","bulkaddtocart.js")%>"></script>    


<% End If %>


<% If pcStrPageName = "onepagecheckout.asp" Then %>  
<script>
    $pc('.collapse').collapse();
</script>
<% End If %>

<script type=text/javascript>
    // $pc("#prdtabs").tabCollapse('show');
</script>

<% 
If pcStrPageName = "configureprd.asp" Or _
      pcStrPageName = "reconfigure.asp" Or _
      pcStrPageName = "prdaddcharges.asp" Or _
      pcStrPageName = "reprdaddcharges.asp" Or _
      InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "configureprd.asp") Or _
      InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "reconfigure.asp") Or _
      InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "prdaddcharges.asp") Or _
      InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "reprdaddcharges.asp") Then %>  
<script type="text/javascript">

	//Display Add to Cart button on the Pricing Panel
	var DisplayCartOnPanel=1;

    if ($pc("#pcBTOfloatPrices").is("div")) {
        $pc("body").addClass("footerFixedPricing");
    }
    $pc(document).ready(function () {
        if ($pc("#pcBTOfloatPrices").is("div")) {

            if ($pc("#addtocart").is("button")) {
				if (DisplayCartOnPanel==1)
                $pc("#addtocart").parent().parent().addClass("addtocart");
                $pc("#pcBTOfloatPrices").addClass("addtocart");
            }

            if ($pc("#add").is("button")) {
				if (DisplayCartOnPanel==1)
                $pc("#add").parent().parent().addClass("addtocart");
                $pc("#pcBTOfloatPrices").addClass("addtocart");
            }

            var defaultPrice = $("input[name='TLcurPrice'][type='TEXT'],input[name='curPrice'][type='TEXT']");
            var customizationPrice = $("input[name='TLtotal'][type='TEXT'],input[name='CMWQD'][type='TEXT']");
            var additionalCharges = $("input[name='total'][type='TEXT']");
            var discountPrice = $("input[name='Discounts'][type='TEXT']");
			var QdiscountPrice = $("input[name='QDiscounts'][type='TEXT']");
            var grandTotal = $("input[name='TotalWithQD'][type='TEXT'],input[name='GrandTotalQD'][type='TEXT']");

            if (defaultPrice.length > 0) defaultPrice.parent().parent().addClass("defaultPrice").addClass("priceItem");
            if (customizationPrice.length > 0) customizationPrice.parent().parent().addClass("customizationPrice").addClass("priceItem");
            if (additionalCharges.length > 0) additionalCharges.parent().parent().addClass("additionalCharges").addClass("priceItem");
            if (discountPrice.length > 0) discountPrice.parent().parent().addClass("discountPrice").addClass("priceItem");
			if (QdiscountPrice.length > 0) QdiscountPrice.parent().parent().addClass("QdiscountPrice").addClass("priceItem");

            if (grandTotal.length > 0) {
                grandTotal.parent().parent().addClass("grandTotal");
                if ((discountPrice.length < 1) && (additionalCharges.length < 1) && (QdiscountPrice.length<1)) {
                    grandTotal.parent().parent().addClass("pcBTOsmall");
                    $pc("body").addClass("pcBTOsmall");
                }
            }

            $pc("#pcBTOfloatPrices").addClass("animated bounceInUp initial");
            $pc(".addtocart").addClass("animated bounceInUp initial");

            $pc("#pcBTOfloatPrices").append("<div class='pcBTOfloatPricesControls'><a href='#openPricing' id='openPricing'><span class='glyphicon glyphicon-collapse-up'></span></a><a href='#closePricing' id='closePricing'><span class='glyphicon glyphicon-collapse-down'></span></a></div>");
            $pc("#closePricing").hide();

            $pc("#openPricing").click(function (e) {
                e.preventDefault();
                $pc("#pcBTOfloatPrices").addClass("open");
                $pc("#closePricing").fadeIn(300);
                $pc("#openPricing").delay(300).fadeOut(300);
            });

            $pc("#closePricing").click(function (e) {
                e.preventDefault();
                $pc("#pcBTOfloatPrices").removeClass("open");
                $pc("#openPricing").fadeIn(300);
                $pc("#closePricing").delay(300).fadeOut(300);
            });

        }
    });
</script>
<% End If %>

<script type="text/javascript"> 
<% If session("Facebook")="1" Then %>
    var facebookActive = true;
<% Else %>
    var facebookActive = false;
<% End If %>
</script>

<%
if Session("idCustomer")>0 AND pcv_strSaveCart="0" then
%>
<script>
	showSaveCartModal();
</script>
<%
end if
%>