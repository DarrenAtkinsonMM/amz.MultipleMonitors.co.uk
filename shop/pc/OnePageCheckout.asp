<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
response.Buffer=true
Response.Expires = -1
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
pcStrPageName="OnePageCheckout.asp"
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/USPSconstants.asp"-->

<!--#include file="opc_init.asp"-->

<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="opc_pageLoad.asp"-->

<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="pcValidateQty.asp"-->

	<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
                        				
							<h3 class="color marginbot-0 h-semi">Checkout</h3><a name="opcShippingAnchor"></a><a name="opcPaymentAnchor"></a>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
<div id="pcMain" data-ng-cloak class="ng-cloak pcOnePageCheckout">
<section id="pc-multisection" class="pc-multisection paddingbot-40">
		<div class="container">
			<div class="row">
                <a id="top"></a>

    <div class="pcMainContent" id="OrderPreviewCtrl" data-ng-controller="OrderPreviewCtrl">
        
        <div data-ng-show="Evaluate(shoppingcart.IsEditOrder)">
            <div class="pcInfoMessage">
                <%= dictLanguage.Item(Session("language")&"_SB_34A") %>
                <%= Session("SBEditOrderID") %>
                <%= dictLanguage.Item(Session("language")&"_SB_34B") %>
            </div>
        </div>

        <% If request("msg")<>"" And IsNumeric(request("msg")) Then %>
            <div class="pcErrorMessage">
                <%
                Select case Clng(request("msg"))
                    Case 1: response.write dictLanguage.Item(Session("language")&"_opc_msg1")
                End Select
                %>
            </div>
        <% End If %>
        
        <div data-ng-show="Evaluate(shoppingcart.ExpressCheckoutInUse) && !Evaluate(shoppingcart.PayWithAmazonInUse)">
            <div class="pcInfoMessage">
                <%=dictLanguage.Item(Session("language")&"_opc_55")%>

                <span data-ng-show="Evaluate(shoppingcart.newCustomer) && shoppingcart.guestCheckoutStatus == '2'"><br /><br /><%=dictLanguage.Item(Session("language")&"_opc_55b")%></span>
            </div>
        </div>
		
		<div data-ng-show="Evaluate(shoppingcart.ExpressCheckoutInUse) && Evaluate(shoppingcart.PayWithAmazonInUse)">
            <div class="pcInfoMessage">
				<div data-ng-show="!Evaluate(shoppingcart.AmazonFirstTime) && Evaluate(shoppingcart.displayShippingAddress)">
                <%=dictLanguage.Item(Session("language")&"_opc_87")%>
				</div>
				<div data-ng-show="!Evaluate(shoppingcart.AmazonFirstTime) && !Evaluate(shoppingcart.displayShippingAddress)">
				<%=dictLanguage.Item(Session("language")&"_opc_92")%>
				</div>
				<div data-ng-show="Evaluate(shoppingcart.AmazonFirstTime) && Evaluate(shoppingcart.displayShippingAddress)">
				<%=dictLanguage.Item(Session("language")&"_opc_88")%>
				</div>
				<div data-ng-show="Evaluate(shoppingcart.AmazonFirstTime) && !Evaluate(shoppingcart.displayShippingAddress)">
				<%=dictLanguage.Item(Session("language")&"_opc_91")%>
				</div>
            </div>
        </div>

        <%
        '/////////////////////////////////////////////////////////////////////////
        '// START:  LOGIN AREA
        '/////////////////////////////////////////////////////////////////////////
        %>

        <div id="LoginOptions" class="pcMainContent" data-ng-show="Evaluate(shoppingcart.displayLogin) && !Evaluate(shoppingcart.IsLoggedIn)">
                               
            <%
            'FB-S            
            IF session("idCustomer")="0" then
                session("pcFBS_TurnOnOff")=0
                query="SELECT pcFBS_TurnOnOff,pcFBS_AppID FROM pcFacebookSettings;"
                set rs=connTemp.execute(query)
                if not rs.eof then
                    session("pcFBS_TurnOnOff")=rs("pcFBS_TurnOnOff")
                    if IsNull(session("pcFBS_TurnOnOff")) OR session("pcFBS_TurnOnOff")="" then
                        session("pcFBS_TurnOnOff")="0"
                    end if
                    session("pcFBS_AppID")=rs("pcFBS_AppID")
                end if
                set rs=nothing
                                       
                IF session("Facebook")="1" OR session("pcFBS_TurnOnOff")="1" THEN 
                    %>  
                    <form name="LoginForm1" method="post" action="checkout.asp" class="pcForms">                
                    <input type="hidden" name="cmode" value="0">
                    <input type="hidden" id="fe" name="fe" value="">
                    <input type="hidden" id="fl" name="fl" value="">
                    <input type="hidden" id="ffn" name="ffn" value="">
                    <input type="hidden" id="fln" name="fln" value="">
                    <input type="hidden" id="fb" name="fb" value="1">
                    </form>
                    <%
                END IF
            END IF
            'FB-E 
            %>

            <form name="loginForm" id="loginForm" class="form">
           
                <div id="opcLoginTable" class="container-fluid">
                
                    <div class="row">

                        <div id="pcShowLoginFields2" class="col-md-6 multsection-col mmc-product">
							<div class="multi-submenu daOpcLoginBox wow fadeInUp animated" data-wow-delay="0.1s" style="visibility: visible; animation-delay: 0.1s; animation-name: fadeInUp;">
                            	<div class="row"><div class="col-sm-12">
                            <h1>
                                <span data-ng-show="shoppingcart.guestCheckoutStatus == 0">
                                    <%=dictLanguage.Item(Session("language")&"_opc_4")%>
                                </span>
                                <span data-ng-show="shoppingcart.guestCheckoutStatus == 1">
                                    <%=dictLanguage.Item(Session("language")&"_opc_4d")%>
                                </span>
                                <span data-ng-show="shoppingcart.guestCheckoutStatus == 2">
                                    <%=dictLanguage.Item(Session("language")&"_opc_4f")%>
                                </span>
                            </h1>
							<h2>Guest Checkout</h2>
							<p>Checkout quickly and easily without needing to sign up for an account:</p>
                            <div id="pcOtherCheckoutMethods">  
                                

                                <div class="pcSpacer"></div>
                                <div class="pcFormButtons">
                                    <a class="pcButton pcButtonContinue btn btn-skin btn-wc semi pcnw-btn margintop-20 marginbot-20" 
                                            name="GuestSubmit" 
                                            id="GuestSubmit" 
                                            data-ng-click="startGuestCheckout(<%=scUseImgsGC%>)"
                                            >
                                            Continue to Guest Checkout
                                            <i class="fa fa-angle-right"></i>
        
                                    </a>
                                </div>

                                <div data-ng-show="Evaluate(shoppingcart.hasExpressCheckout)">
                                
                                    <h4><%=dictLanguage.Item(Session("language")&"_opc_86")%></h4>
                                    
                                    <!--#include file="pcPay_PayPal.asp"-->
																	
                                    <%
                                    If hasPayPalButtons And pcAmazonTurnOn = "1" Then
                                        %><div class="pcAltCheckoutSeparator">OR</div><%
                                    End If
                                    %>

                                    <!--#include file="inc_AmazonButton.asp"-->

                                </div>
							<p style="font-size:80%; font-style:italic;">You can save a full account during the checkout process or on the final order confirmation page if you would like to however this is not needed for a sale.</p>
                            </div>
                            </div></div>
                            </div>
                        </div>
                        <div id="pcShowLoginFields" class="col-md-6 multsection-col mmc-product">
                        	<div class="multi-submenu daOpcLoginBox wow fadeInUp animated" data-wow-delay="0.1s" style="visibility: visible; animation-delay: 0.1s; animation-name: fadeInUp;">
                            	<div class="row"><div class="col-sm-12">
                            <h1>Existing Customers</h1>
								<h2>Login To Your Account</h2>

                            <div id="opcLoginFormFields">
                            
                                <div id="opcLoginEmail" class="form-group">
                                    <div class="row specb-row">
                                <div class="col-sm-2 daOpcLoginTitle">
								<p>Email:</p>
                                </div>
                                <div class="col-sm-10 specb-field">
                                    <input type="email" class="form-control" name="email" id="email" size="25">
                                </div>
                                </div>
                            	</div>
                                <div id="opcLoginPassword" class="form-group">
                                <div class="row specb-row">
                                <div class="col-sm-2 daOpcLoginTitle">
								<p>Password:</p>
                                </div>
                                <div class="col-sm-10 specb-field">
                                    
                                    <input type="password" class="form-control" name="password" id="password" size="25">
                               </div>
                                </div>
                            	</div>

                                <% 
                                '// If Advanced Security is turned on 
                                call advancedSecurity()
                                %>
    
                                <div id="LoginMessageBox" style="display: none"></div>
    
                                <div class="pcFormButtons">
                                      
                                        <button class="pcButton pcButtonContinue btn btn-skin btn-wc semi pcnw-btn margintop-10 marginbot-20" 
                                                data-ng-click="login()"
                                                name="LoginSubmit" 
                                                id="LoginSubmit"
                                                type="button"
                                                >Login To Your Account <i class="fa fa-angle-right"></i>
                                        </button>
                                        
                                        <%
                                        'FB-S
                                        IF session("Facebook")="1" OR session("pcFBS_TurnOnOff")="1" THEN 
                                            'Show FB button below Continue
                                            %>  
                                            <div class="pcSpacer"></div>
                                            <div id="fb-root"></div>
                                            <fb:login-button show-faces="false" size="medium" width="200" max-rows="1" scope="email" onlogin="checkLoginState();">Login with Facebook</fb:login-button>
                                            <%
                                        END IF
                                        'FB-E 
                                        %>
    
                                        <div class="pcSpacer"></div>

                                           <p>Forgot your password? <a href="checkout.asp?cmode=2&fmode=<%=pcPageMode%>&orderReview=no">
                                                <%= dictLanguage.Item(Session("language")&"_Custva_8")%>
                                            </a></p>

                                    
                                </div>
                                
                                
                            </div>
                            
                            <div class="pcSpacer"></div>
                            </div></div>
                            </div>
                        </div>


                    </div>
                    
                </div>
            </form>
        </div>
        <div class="pcSpacer"></div>

        <div id="LoginOptions" data-ng-show="!Evaluate(shoppingcart.displayLogin) && Evaluate(shoppingcart.IsLoggedIn)">

            <div class="pcFormItem" data-ng-show="!Evaluate(shoppingcart.guestCustomer)">
                <span>
                    <%=dictLanguage.Item(Session("language")&"_opc_7")%>
                    <a href="custPref.asp">{{shoppingcart.billingAddress.FirstName}} {{shoppingcart.billingAddress.LastName}}</a>
                </span>
                <div class="pcClear"></div>
            </div>

        </div>        
        <%
        '/////////////////////////////////////////////////////////////////////////
        '// END:  LOGIN AREA
        '/////////////////////////////////////////////////////////////////////////
        %>



        <%
        '/////////////////////////////////////////////////////////////////////////
        '// START:  ACCORDION
        '/////////////////////////////////////////////////////////////////////////
        %>
        <div id="acc1" class="panel-group" data-ng-show="Evaluate(shoppingcart.IsLoggedIn) || Evaluate(guestSession)">

            <div id="opcBillingPanel" class="panel panel-default">
            
                <% '// START: BILLING ADDRESS HEADING %>
                <div id="opcLoginAnchor" class="panel-heading">
    
                    <span class="pcCheckoutTitle panel-title"><%=dictLanguage.Item(Session("language")&"_opc_8")%></span>
                    <div class="StatusIndicators" data-ng-show="showBillingEditArea()">
                        <a id="btnEditCO" 
                            class="pcButton secondary daOPCEditLink"
                            href="javascript:;" 
                            data-ng-click="switchPanel('billing')"
                            ><%=dictLanguage.Item(Session("language")&"_opc_53") %></a>
                    </div>
                    <div id="BillingAddress" data-ng-show="showBillingEditArea()">
                        <div class="editbox daOPCEditBox" style="margin-top: 6px;">
                            {{shoppingcart.billingAddress.FirstName}} {{shoppingcart.billingAddress.LastName}} <br />
                            <span data-ng-show="!IsEmpty(shoppingcart.billingAddress.company)">
                                {{shoppingcart.billingAddress.company}} <br />
                            </span>
                            {{shoppingcart.billingAddress.address}} <br />
                            <span data-ng-show="!IsEmpty(shoppingcart.billingAddress.address2)">
                                {{shoppingcart.billingAddress.address2}} <br />
                            </span>
                            {{shoppingcart.billingAddress.city}}&nbsp;{{shoppingcart.billingAddress.state}}{{shoppingcart.billingAddress.province}}&nbsp; {{shoppingcart.billingAddress.postalCode}}
                        </div>
                    </div>
    
                </div>
                <% '// END: BILLING ADDRESS HEADING %>
    
    
                <% '// START: BILLING ADDRESS BODY %>
                <div id="pcBillingPanelContent" class="panel-collapse collapse in">
                    <div class="panel-body">
    
                        <div id="BillingArea" style="display: none">
        
                            <form name="BillingForm" id="BillingForm" class="form" method="post">
                            	
								<% If (ptaxAvalara = 1 AND ptaxAvalaraEnabled = 1 AND ptaxAvalaraAddressValidation = 1) OR USPS_AddressValidation = 1 Then %>
                            	<input type="hidden" id="IsBillingAddressValidated" name="IsBillingAddressValidated" value="" />
                                <input type="hidden" id="billingAddressToken" name="billingAddressToken" value="{{shoppingcart.billingAddress.token}}" data-ng-model="shoppingcart.billingAddress.token" />
                                <% End If %>
                                
                                <!-- First billing column -->
                                <div class="col-md-6">

                                <% 'Billing First Name %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>First Name: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="billfname" id="billfname" data-ng-model="shoppingcart.billingAddress.FirstName">
                                </div>
                                </div>


                                <% 'Billing Last Name %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Last Name: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="billlname" id="billlname" data-ng-model="shoppingcart.billingAddress.LastName">
                                    <input type="hidden" name="billemail2" id="billemail2" value="{{shoppingcart.billingAddress.email}}" />
                                </div>
                                </div>
                                
                                
                                <% 'Billing Email %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Email Address: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="email" class="form-control" name="billemail" id="billemail" data-ng-model="shoppingcart.billingAddress.email">
                                </div>
                                </div>
                                
                                <% '// Billing Guest Fields %>

                                <% 'Billing Password %>
                                <div id="opcBillingPassword" class="form-group" data-ng-show="((shoppingcart.guestCheckoutStatus) == 2 && (Evaluate(shoppingcart.displayGuestFields)))">
                                    <label for="billpass"><%= dictLanguage.Item(Session("language")&"_opc_6") %><img class="pcRequiredIcon" src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>"></label>
                                    <input type="password" class="form-control" name="billpass" id="billpass" autocomplete="off">
                                </div>
        
                                <% 'Billing Password Confirmation %>
                                <div id="opcBillingPasswordConfirm" class="form-group" data-ng-show="((shoppingcart.guestCheckoutStatus) == 2 && (Evaluate(shoppingcart.displayGuestFields)))">
                                    <label for="billrepass"><%= dictLanguage.Item(Session("language")&"_opc_38") %><img class="pcRequiredIcon" src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>"></label>
                                    <input type="password" class="form-control" name="billrepass" id="billrepass" autocomplete="off">
                                </div>       
                                                             
                                <% 'Billing Phone %>
                                 <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Telephone: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="billphone" id="billphone" data-ng-model="shoppingcart.billingAddress.phone">
                                    <input type="hidden" class="form-control" name="billfax" id="billfax" data-ng-model="shoppingcart.billingAddress.fax">
                                </div>
                                </div>

                                <% 'Billing Company %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Company Name: </p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="billcompany" id="billcompany" data-ng-model="shoppingcart.billingAddress.company">
                                </div>
                                </div> 
                                
                                                               
                                <% 'VAT %>
                                <div id="opcBillingVATID" class="row specb-row" data-ng-show="Evaluate(ShowVatId)">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>VAT Number:</p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="billVATID" id="billVATID" data-ng-model="shoppingcart.billingAddress.VATID">
                                    <p><a class="daOPCVATInfoLink" data-toggle="lightbox" data-title="VAT / Tax Rules" href="/pop-pages/opc-vat-rules.htm">Learn about VAT charges and rules</a></p>
                                </div>
                                </div>
                                
                                </div><!-- End first billing column -->
                                <!-- Second billing column -->
                                <div class="col-md-6">
        

                                <% 'SSN %>
                                <div id="opcBillingSSN" class="form-group" data-ng-show="Evaluate(ShowSSN)">
                                    <label for="billSSN"><%=dictLanguage.Item(Session("language")&"_Custmoda_24")%></label>
                                    <input type="text" class="form-control" name="billSSN" id="billSSN" data-ng-model="shoppingcart.billingAddress.SSN">
                                </div>


                                <% 'Billing Street Address %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Address Line 1: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="billaddr" id="billaddr" data-ng-model="shoppingcart.billingAddress.address">
                                </div>
                                </div>


                                <% 'Billing Street Address 2 %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Address Line 2:</p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="billaddr2" id="billaddr2" data-ng-model="shoppingcart.billingAddress.address2">
                                </div>
                                </div>


                                <% 'Billing City %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>City: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="billcity" id="billcity" data-ng-model="shoppingcart.billingAddress.city">
                                </div>
                                </div>


                                <% 'Billing Zip %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Post Code: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="billzip" id="billzip" data-ng-model="shoppingcart.billingAddress.postalCode">
                                </div>
                                </div>


                                <% 'Load Billing State/Country Boxes %>
                                <%
                                pcv_strTargetForm 	= "BillingForm" '// Name of Form
                                pcv_strCountryBox 	= "billcountry" '// Name of Country Dropdown
                                pcv_strTargetBox	= "billstate" '// Name of State Dropdown
                                pcv_strProvinceBox 	= "billprovince" '// Name of Province Field
                                tmp_CountryBoxFunc	= "switchZipName1(this.value);"
                                pcv_strContextName	= "Billing"
                                
                                If Session("idCustomer")>0 Then
                                    pcv_strCountryModel = "shoppingcart.billingAddress.country"
                                    pcv_strStateModel = "shoppingcart.billingAddress.state"
                                    pcv_strProvinceModel = "shoppingcart.billingAddress.province"
                                End If
                                
                                '// Set local Country to Session
                                If Session(pcv_strSessionPrefix & pcv_strCountryBox) = "" Then
                                    Session(pcv_strSessionPrefix & pcv_strCountryBox) = pcStrBillingCountryCode
                                End If
                                
                                '// Set local State to Session
                                If Session(pcv_strSessionPrefix & pcv_strTargetBox) = "" Then
                                    Session(pcv_strSessionPrefix & pcv_strTargetBox) = pcStrBillingStateCode
                                End If
                                
                                '// Set local Province to Session
                                If Session(pcv_strSessionPrefix & pcv_strProvinceBox) = "" Then
                                    Session(pcv_strSessionPrefix & pcv_strProvinceBox) = pcStrBillingProvince
                                End If
                                %>
                                <!--#include file="../includes/javascripts/opc_pcStateAndProvince.asp"-->
        
                                <% 'Billing State/Province %>
                                <% 
                                pcs_StateProvince
                                %> 
                                
                                <% 'Billing Country %>
                                <%
                                pcs_CountryDropdown
                                %>

                                <% 'Billing Address Type (Residential/Commercial) %>
                                <div id="billAddrTypeArea" class="form-group" data-ng-show="billingAddressType()">
                                    <label for="pcAddressType"><%=dictLanguage.Item(Session("language")&"_opc_23")%></label>
                                    <input type="radio" name="pcAddressType" value="1" checked>&nbsp;<span class="pcSmallText"><%=dictLanguage.Item(Session("language")&"_opc_24")%></span>&nbsp;
                                    <input type="radio" name="pcAddressType" value="0">&nbsp;<span class="pcSmallText"><%=dictLanguage.Item(Session("language")&"_opc_25")%></span>
                                    <!-- TO DO: Move default value to backend processing script... <input type="hidden" name="pcAddressType" value="<%=pcComResShipAddress%>"> -->
                                </div>
                                <p>&nbsp;</p>
                                </div><!-- Second billing column -->
        
                                <% 'Special Customer Fields %>
                                <% call specialCustomerFields() %>
        
                                <% 'Referrer Field %>
                                <% call referrerFields() %>
        
                                <% 'Newsletter %>
                                <% call newsletterFields() %>
        
                                <% 'Terms Area %>
                                <% call termsAndConditions() %>
        
                                <div data-ng-class="getLayoutClass()" class="opcRow">
                                    <div id="BillingMessageBox" class="pcAttention" style="display: none"></div>
                                </div>
                                
                                <div class="pcSpacer"></div>
        
                                <% 'Billing Submit %>
                                <div class="pcFormButtons">
        
                                    <button type="button" class="pcButton pcButtonContinue btn btn-skin btn-wc updateBtn" data-ng-click="updateBilling()" name="BillingSubmit" id="BillingSubmit">
                                        <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_update")%></span>
                                    </button>
        
                                    <span data-ng-show="Evaluate(displayLogin)">
                                    
                                        &nbsp;
                                        <button type="button" class="pcButton pcButtonBack" name="BillingCancel" id="BillingCancel">
                                            <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back")%></span>
                                        </button>
                                        
                                    </span>
                                    
                                </div>
        
                            </form>
        
                        </div>

                    </div>
                </div>  
                <% '// END: BILLING ADDRESS BODY %>

            </div>


            <div id="opcShippingPanel" class="panel panel-default" data-ng-show="!Evaluate(shoppingcart.IsHideShippingPanel)">
            
                <% '// START: SHIPPING ADDRESS HEADING %>
                <div id="opcShippingAddressAnchor" class="panel-heading" data-ng-show="Evaluate(shoppingcart.displayShippingAddress)">
    
                    <span class="pcCheckoutTitle panel-title">Shipping Information<%'=dictLanguage.Item(Session("language")&"_opc_2")%><%'=dictLanguage.Item(Session("language")&"_opc_27")%></span>
                   
                    <div class="StatusIndicators" data-ng-show="showShippingEditArea()">
                        <a id="btnEditCO" 
                            class="pcButton secondary daOPCEditLink"
                            href="javascript:;" 
                            data-ng-click="switchPanel('shipping')"
                            ><%=dictLanguage.Item(Session("language")&"_opc_53") %></a>
                    </div>
                    
                    <div id="ShippingAddress" data-ng-show="showShippingEditArea()">                   
                        <div class="editbox daOPCEditBox" style="margin-top: 6px;">
                            {{shoppingcart.shippingAddress.FirstName}} {{shoppingcart.shippingAddress.LastName}} <br />
                            <span data-ng-show="!IsEmpty(shoppingcart.shippingAddress.company)">
                                {{shoppingcart.shippingAddress.company}} <br />
                            </span>
                            {{shoppingcart.shippingAddress.address}} <br />
                            <span data-ng-show="!IsEmpty(shoppingcart.shippingAddress.address2)">
                                {{shoppingcart.shippingAddress.address2}} <br />
                            </span>
                            {{shoppingcart.shippingAddress.city}}&nbsp;{{shoppingcart.shippingAddress.state}}{{shoppingcart.shippingAddress.province}}&nbsp; {{shoppingcart.shippingAddress.postalCode}}
                        </div>
                    </div>
    
                </div>
                <% '// END: SHIPPING ADDRESS HEADING %>
    
    
                <% '// START: SHIPPING ADDRESS BODY %>
                <div id="pcShippingPanelContent" class="panel-collapse collapse">
                    <div class="panel-body">
                        
                        
                    <div id="ShippingArea">
						<%if session("PayWithAmazon")="YES" then%>
						<div data-ng-show="Evaluate(shoppingcart.PayWithAmazonInUse)">
							<div id="addressBookWidgetDiv">
							</div> 
							<script type=text/javascript>
							var AmzShippingSelected=0;
							var AmazonOrderReferenceId="";
							new OffAmazonPayments.Widgets.AddressBook({
								sellerId: '<%=pcAMZSellerID%>',
								onOrderReferenceCreate: function(orderReference) {
								AmazonOrderReferenceId=orderReference.getAmazonOrderReferenceId();
							},
							onAddressSelect: function(orderReference) {
							 	AmzShippingSelected=1;
							},
							design: {
								size : {width:'400px', height:'260px'}
							},
							onError: function(error) {
								alert("<%=dictLanguage.Item(Session("language")&"_AmazonPay_6")%>");
								document.location="viewcart.asp";
							}
							}).bind("addressBookWidgetDiv");
							</script>
							
							<div class="pcFormButtons">
                            
                                <button type="button" class="pcButton pcButtonLogin" data-ng-click="updateAmazonShipping()" name="ShippingSubmit" id="ShippingSubmit">
                                    <img src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" alt="Submit" />
                                    <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_update")%></span>
                                </button>
                                
                            </div>
						</div>
						<%end if%>
						
						<div data-ng-show="!Evaluate(shoppingcart.PayWithAmazonInUse)">
    
                        <form name="ShippingForm" id="ShippingForm" method="post">
                        	
							<% If (ptaxAvalara = 1 AND ptaxAvalaraEnabled = 1 AND ptaxAvalaraAddressValidation = 1) OR USPS_AddressValidation = 1 Then %>
                            <input type="hidden" id="IsShippingAddressValidated" name="IsShippingAddressValidated" value="" />
                            <input type="hidden" id="shippingAddressToken" name="shippingAddressToken" value="{{shoppingcart.shippingAddress.token}}" data-ng-model="shoppingcart.shippingAddress.token" />
                            <% End If %>
                            
                            <div id="opcShippingRadios" data-toggle="buttons" class="btn-group">
                                
                                <% if (session("Cust_IDEvent")="") OR (session("Cust_IDEvent")<>"" AND gDelivery=0) then %>
                                    <label class="btn btn-default" for="rad_0">
                                        <input radio-with-change-handler id="rad_0" type="radio" name="ShipArrOpts" value="-1" data-ng-model="ShipArrOpts" /> <span id="rad_0" class="pcCheckBox"></span><%=dictLanguage.Item(Session("language")&"_opc_20")%>
                                    </label>
                                <% end if %>
    
                                <% if session("Cust_IDEvent")<>"" then %>
                                    <label class="btn btn-default" for="rad_1">
                                        <input radio-with-change-handler id="rad_1" type="radio" name="ShipArrOpts" value="-2" data-ng-model="ShipArrOpts" /> <span id="rad_1" class="pcCheckBox"></span><%=dictLanguage.Item(Session("language")&"_opc_21")%>
                                    </label>
                                <% end if %>
    
                            </div>
    
                            <div class="pcShowContent" id="shippingAddressArea" style="display: none;">
                            
                                <div id="copyfromBillingLink">
                                    <a href="javascript:copyfromBillAddr();"><%=dictLanguage.Item(Session("language")&"_opc_msg2")%></a>
                                    <div class="pcSpacer"></div>                                
                                </div>
    
   								<!-- First shipping column -->
                                <div class="col-md-6">
    
                                <% 'Shipping Nickname %>

                                    <input type="hidden" class="form-control" name="shipnickname" id="shipnickname">

    
                                <% 'Shipping Name %>
                                <div id="shipnameArea">
                                
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>First Name: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                        <input type="text" class="form-control" name="shipfname" id="shipfname">
                                    </div></div>
                                    
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Last Name: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                        <input type="text" class="form-control" name="shiplname" id="shiplname">
                                    </div></div>

                                </div>
    
                                <% 'Shipping Email %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Email Address: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="shipemail" id="shipemail">
                                </div></div>
                                
                                <% 'Shipping Phone %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Telephone: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="shipphone" id="shipphone">
                                </div></div>
                                

                                <% 'Shipping Company %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Company Name: </p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="shipcompany" id="shipcompany">
                                </div></div>
                                
                                </div><!-- End first shipping column -->

								<!-- Second shipping column -->
                                <div class="col-md-6">

                                <% 'Shipping Address %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Address Line 1: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="shipaddr" id="shipaddr">
                                </div></div>

    
                                <% 'Shipping Address 2 %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Address Line 2: </p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="shipaddr2" id="shipaddr2">
                                </div></div>

    
                                <% 'Shipping City %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>City: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="shipcity" id="shipcity">
                                </div></div>

    
                                <% 'Shipping Zip %>
                                <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Post Code: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
                                    <input type="text" class="form-control" name="shipzip" id="shipzip">
                                </div></div>

    
                                <% 'Load Shipping State/Country %>
                                <%
                                pcv_strTargetForm = "ShippingForm" '// Name of Form
                                pcv_strCountryBox = "shipcountry" '// Name of Country Dropdown
                                pcv_strTargetBox = "shipstate" '// Name of State Dropdown
                                pcv_strProvinceBox =  "shipprovince" '// Name of Province Field
                                tmp_CountryBoxFunc = "switchZipName2(this.value);"
                                pcv_strContextName	= "Shipping"
                                
                                If Session("idCustomer")>0 Then
                                    pcv_strCountryModel = "shoppingcart.shippingAddress.country"
                                    pcv_strStateModel = "shoppingcart.shippingAddress.state"
                                    pcv_strProvinceModel = "shoppingcart.shippingAddress.province"
                                End If
                      
                                '// Set local Country to Session
                                if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
                                    Session(pcv_strSessionPrefix&pcv_strCountryBox) = pcStrShippingCountryCode
                                end if
                    
                                '// Set local State to Session
                                if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
                                    Session(pcv_strSessionPrefix&pcv_strTargetBox) = pcStrShippingStateCode
                                end if
                        
                                '// Set local Province to Session
                                if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
                                    Session(pcv_strSessionPrefix&pcv_strProvinceBox) =  pcStrShippingProvince
                                end if
                                %>
    
                                <% 'Shipping Country %>
                                <%
                                pcs_CountryDropdown
                                %>
    
                                <% 'Shipping State/Province %>
                                <% 
                                pcs_StateProvince
                                %>
    


    
                                <% 'Shipping Fax %>

                                    <input type="hidden" class="form-control" name="shipfax" id="shipfax">
                                
                                </div><!-- End second shipping column -->
                                
                            </div>
                            
                            <% 'Shipping Address Type %>
                            <div class="opcFormRow" id="shipAddrTypeArea" data-ng-show="shippingAddressType()">
                            
                                <div id="shipnicknameArea" class="form-group">
                                    <label for="pcAddressType"><%=dictLanguage.Item(Session("language")&"_opc_23")%></label>
                                    <input type="radio" name="pcAddressType" value="1" checked>&nbsp;<span class="pcSmallText"><%=dictLanguage.Item(Session("language")&"_opc_24")%></span>&nbsp;
                                    <input type="radio" name="pcAddressType" value="0">&nbsp;<span class="pcSmallText"><%=dictLanguage.Item(Session("language")&"_opc_25")%></span>
                                </div>

                            </div>
    
                            <% 'Shipping Delivery Area %>
                            <div class="opcFormRow" id="shipDeliveryArea" data-ng-show="displayDeliveryArea()">
        
                                <% 'Delivery Date Field %>
                                <span data-ng-show="Evaluate(shoppingcart.displayDateField)">
    
                                    <div class="form-group">
                                        <label for="DF1">{{shoppingcart.dateFieldLabel}}:<div class="pcRequiredIcon" data-ng-show="Evaluate(shoppingcart.dateFieldRequired)"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div></label>
                                        <input class="form-control datepicker" type="text" id="DF1" name="DF1" data-ng-model="shoppingcart.dateFieldValue" data-ng-required="Evaluate(shoppingcart.dateFieldRequired)">

                                        <span class="help-text" data-ng-show="Evaluate(shoppingcart.displayBlackOutDates)">
                                            <a href="javascript:openbrowser('blackoutDates.asp')"><%= dictLanguage.Item(Session("language")&"_catering_20")%></a><%= dictLanguage.Item(Session("language")&"_catering_21")%>
                                        </span>
                                    </div>
             
                                </span>
    
                                <% 'TF Field? %>
                                <div class="form-group" data-ng-show="Evaluate(shoppingcart.displayTimeField)">
                                    <label for="TF1">{{shoppingcart.timeFieldLabel}}:<div class="pcRequiredIcon" data-ng-show="Evaluate(shoppingcart.timeFieldRequired)"><img src="<%=pcf_getImagePath("",pcv_strRequiredIcon)%>" alt="Required"></div></label>
                                    <select class="form-control" id="TF1" name="TF1" data-ng-model="shoppingcart.timeFieldValue" data-ng-required="Evaluate(shoppingcart.timeFieldRequired)">
                                        <% If pcSFTF1="" Then %>
                                        <option value=""><%= dictLanguage.Item(Session("language")&"_viewCatOrder_6")%></option>
                                        <% End If %>
                                        <% If scDateFrmt="DD/MM/YY" Then %>
                                            <option value="7:00">7:00</option>
                                            <option value="7:30">7:30</option>
                                            <option value="8:00">8:00</option>
                                            <option value="8:30">8:30</option>
                                            <option value="9:00">9:00</option>
                                            <option value="9:30">9:30</option>
                                            <option value="10:00">10:00</option>
                                            <option value="10:30">10:30</option>
                                            <option value="11:00">11:00</option>
                                            <option value="11:30">11:30</option>
                                            <option value="12:00">12:00</option>
                                            <option value="12:30">12:30</option>
                                            <option value="13:00">13:00</option>
                                            <option value="13:30">13:30</option>
                                            <option value="14:00">14:00</option>
                                            <option value="14:30">14:30</option>
                                            <option value="15:00">15:00</option>
                                            <option value="15:30">15:30</option>
                                            <option value="16:00">16:00</option>
                                            <option value="16:30">16:30</option>
                                            <option value="17:00">17:00</option>
                                            <option value="17:30">17:30</option>
                                            <option value="18:00">18:00</option>
                                            <option value="18:30">18:30</option>
                                            <option value="19:00">19:00</option>
                                            <option value="19:30">19:30</option>
                                            <option value="20:00">20:00</option>
                                            <option value="20:30">20:30</option>
                                            <option value="21:00">21:00</option>
                                        <% Else %>
                                            <option value="7:00 AM">7:00 AM</option>
                                            <option value="7:30 AM">7:30 AM</option>
                                            <option value="8:00 AM">8:00 AM</option>
                                            <option value="8:30 AM">8:30 AM</option>
                                            <option value="9:00 AM">9:00 AM</option>
                                            <option value="9:30 AM">9:30 AM</option>
                                            <option value="10:00 AM">10:00 AM</option>
                                            <option value="10:30 AM">10:30 AM</option>
                                            <option value="11:00 AM">11:00 AM</option>
                                            <option value="11:30 AM">11:30 AM</option>
                                            <option value="12:00 PM">12:00 PM</option>
                                            <option value="12:30 PM">12:30 PM</option>
                                            <option value="1:00 PM">1:00 PM</option>
                                            <option value="1:30 PM">1:30 PM</option>
                                            <option value="2:00 PM">2:00 PM</option>
                                            <option value="2:30 PM">2:30 PM</option>
                                            <option value="3:00 PM">3:00 PM</option>
                                            <option value="3:30 PM">3:30 PM</option>
                                            <option value="4:00 PM">4:00 PM</option>
                                            <option value="4:30 PM">4:30 PM</option>
                                            <option value="5:00 PM">5:00 PM</option>
                                            <option value="5:30 PM">5:30 PM</option>
                                            <option value="6:00 PM">6:00 PM</option>
                                            <option value="7:00 PM">7:00 PM</option>
                                            <option value="7:30 PM">7:30 PM</option>
                                            <option value="8:00 PM">8:00 PM</option>
                                            <option value="8:30 PM">8:30 PM</option>
                                            <option value="9:00 PM">9:00 PM</option>
                                        <% End If %>
                                    </select>
                                    
                                </div>
   
                                <% 'Delivery Date Message %>
                                <span class="help-text" data-ng-show="Evaluate(shoppingcart.displayDeliveryDateMessage)">
                                    <i><%= dictLanguage.Item(Session("language")&"_catering_6")%></i>
                                </span>
      
                            </div>
    
                            <div data-ng-class="getLayoutClass()" class="opcRow">
                                <div id="ShippingMessageBox" class="pcAttention" style="display: none"></div>
                                <div id="OPRArea" class="pcAttention" style="display: none"></div>
                            </div>
                            
                            <div class="pcSpacer"></div>
    
                            <% 'Shipping Submit %>
                            <div class="pcFormButtons">
                            
                                <button type="button" class="pcButton pcButtonLogin btn btn-skin btn-wc updateBtn" data-ng-click="updateShipping()" name="ShippingSubmit" id="ShippingSubmit">
                                    <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_update")%></span>
                                </button>
                                
                            </div>
                            
                            <div class="pcClear"></div>
                            
                        </form>
						
						</div>
    
                    </div>
    
                    </div>
                </div>
                <% '// END: SHIPPING ADDRESS BODY %>

            </div>


            <div id="opcRatesPanel" class="panel panel-default" data-ng-show="!Evaluate(shoppingcart.IsHideDeliveryPanel)">
            
                <% '// START: RATES HEADING %>            
                <div id="opcShippingAddressAnchor" class="panel-heading" data-ng-show="Evaluate(shoppingcart.displayRatesArea)">
    
                    <span class="pcCheckoutTitle panel-title">Delivery Information<%'=dictLanguage.Item(Session("language")&"_opc_2")%><%'=dictLanguage.Item(Session("language")&"_opc_27")%></span>
                   
                    <div class="StatusIndicators" data-ng-show="showRatesEditArea()">
                        <a id="btnEditCO" 
                            class="pcButton secondary daOPCEditLink"
                            href="javascript:;" 
                            data-ng-click="switchPanel('rates')"
                            ><%=dictLanguage.Item(Session("language")&"_opc_53") %></a>
                    </div>
    
                    <div id="ShippingMethod" data-ng-show="showRatesEditArea()">
                        <div class="editbox daOPCEditBox" style="margin-top: 6px;">
                            {{shoppingcart.shippingMethod}}
                        </div>
                    </div>           
    
                </div>
                <% '// END: RATES HEADING %>
    
    
                <% '// START: RATES BODY %>
                <div id="pcRatesPanelContent" class="panel-collapse collapse">
                    <div class="panel-body">
                    
                        <div id="ShipChargeLoadContentMsg" style="display: none;"></div>
                        <div id="ShippingChargeArea" style="display: none;"></div>
                    
                    </div>
                </div>
    
                <% '// END: RATES BODY %> 
                
            </div>


            <div id="opcPaymentPanel" class="panel panel-default">
            
                <% '// START: PAYMENT INFORMATION HEADING %>
                <div id="opcPaymentAnchor" class="panel-heading">
                    
                    <div class="StatusIndicators">
                        <a id="btnEditPay" href="javascript:;" onClick="javascript: OPCopenPanel('opcShipping') " title="<%=dictLanguage.Item(Session("language")&"_opc_45")%>" style="display: none;">Edit</a>
                    </div>
                    
                    <span class="pcCheckoutTitle panel-title"><%=dictLanguage.Item(Session("language")&"_opc_28")%></span>
                
                </div>
                <% '// END: PAYMENT INFORMATION HEADING %>
    
    
                <% '// START: PAYMENT INFORMATION BODY %>
                <div id="pcPaymentPanelContent" class="panel-collapse collapse">
                    <div class="panel-body">
                
                        <div id="TaxLoadContentMsg" style="display: none;"></div>
                        <div id="TaxContentArea" style="display: none;"></div>
                        <div id="PaymentContentArea">
        
                            <%
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            '// Password - Start
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
                            %>                    
                            <div data-ng-show="Evaluate(shoppingcart.allowPassword) && Evaluate(shoppingcart.DisplayOptionalPassword)">
        
                                    <div id="PwdArea" class="pcShowContent" style="display: none;">
                                
                                        <form id="PwdForm" name="PwdForm">
                                            
                                            <div class="daOPCAccSaveMsg"><span class="daOPCAccSaveHeader">Save an Account:</span> (Optional) <%=dictLanguage.Item(Session("language")&"_opc_common_3")%></div>
        
                                            <div data-ng-class="getLayoutClass()" id="opcAddPassword">
                                                <div class="pcFormItem">
                                                    <div class="pcFormLabel"><label for="newPass1"><%=dictLanguage.Item(Session("language")&"_opc_6")%><label></div>
                                                    <div class="pcFormField">
                                                        <input class="form-control input-sm" type="password" name="newPass1" id="newPass1" size="20" autocomplete="off" />
                                                    </div>
                                                </div>
                                            </div>
                
                                            <div data-ng-class="getLayoutClass()" id="opcAddPasswordConfirm">
                                                <div class="pcFormItem">
                                                    <div class="pcFormLabel"><label for="newPass2"><%=dictLanguage.Item(Session("language")&"_opc_38")%><label></div>
                                                    <div class="pcFormField">
                                                        <input class="form-control input-sm" type="password" name="newPass2" id="newPass2" size="20" autocomplete="off" />
                                                    </div>
                                                </div>
                                            </div>

                                            <div id="PwdLoader"></div>
        
                                            <div class="pcFormButtons">
                                            
                                                <button class="pcButton pcButtonSavePassword btn btn-skin opcSaveBtn" data-ng-click="savePassword()" name="PwdSubmit" id="PwdSubmit">                                                    
                                                    <%= dictLanguage.Item(Session("language")&"_opc_90")%>
                                                </button>
        
                                                <div id="PwdNoThanks" data-ng-show="shoppingcart.guestCheckoutStatus==0 || shoppingcart.guestCheckoutStatus==''">
                                                    <a href="javascript:;" onClick="$pc('#PwdLoader').hide(); $pc('#PwdWarning').hide(); $pc('#PwdArea').hide();"><%=dictLanguage.Item(Session("language")&"_opc_51")%></a>
                                                </div>
                                            
                                            </div>
                                        
                                        </form>
                                
                                        <div class="pcClear"></div>
                            
                                    </div>
                            </div>
                            <%
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            '// Password - End
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
        
        
        
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            '// Other - Start
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
                            %>
                            <div id="OtherArea">
                            	<div class="daOPCAccSaveMsg"><span class="daOPCAccSaveHeader">Add Order Comments:</span> (Optional) Let us know of any specific comments or notes about this order.</div>
                                <form id="OtherForm" name="OtherForm">
                                    
                                    
        
                                        <% '// Nickname - S %>
                                        <li id="opcOrderName" data-ng-show="Evaluate(shoppingcart.AllowNicknames)">
                                            
                                            <a href="javascript:;" onClick="togglediv('nickname');"><img src="<%=pcf_getImagePath("images","edit.gif")%>" alt="Edit"></a>
                                            <a href="javascript:;" onClick="togglediv('nickname');"><%= dictLanguage.Item(Session("language")&"_catering_13b")%></a>
        
                                            <div id="nickname" class="pcBorderDiv" style="display: none;">
        
                                                <ul class="pcListLayout">
                                                    <li><%= dictLanguage.Item(Session("language")&"_catering_1")%></li>
                                                    <li>
                                                        <label for="OrderNickName"><%= dictLanguage.Item(Session("language")&"_catering_12")%></label>
                                                        <input class="form-control input-sm" type="text" id="OrderNickName" name="OrderNickName" value="{{shoppingcart.orderNickName}}" size="20">
                                                    </li>
                                                </ul>
                                                
                                            </div>
                                        </li>
                                        <% '// Nickname - E %>
        
        
                                        <% '// Comments - S %>
                                        
        
                                            <a href="javascript:;" onClick="togglediv('comments');"><img src="<%=pcf_getImagePath("images","edit.gif")%>" alt="Edit"></a>
                                            <a href="javascript:;" onClick="togglediv('comments');">Add comments or notes to your order</a>
        
                                            <div id="comments" class="pcBorderDiv" style="display: none">
                                                <ul class="pcListLayout">
                                                    <li><textarea class="form-control input-sm" name="OrderComments" cols="50" rows="3">{{shoppingcart.orderComments}}</textarea></li>
                                                </ul>
                                            </div>
                                        
                                        <% '// Comments - E %>
                                        
                                    
        
        
        
                                    <% '// Gift Certs - S %>
                                    <div data-ng-show="Evaluate(shoppingcart.HaveGcs)">
                                    
                                        <div class="pcSpacer"></div>
                                        <div id="opcGiftCert">
                                        
                                            <div class="pcSectionTitle">
                                                <img src="<%=pcf_getImagePath("images","pc4_notify.png")%>" alt="" style="margin-right: 4px;">
                                                <%= dictLanguage.Item(Session("language")&"_NotifyRe_1")%>
                                            </div>
            
                                            <ul class="pcListLayout">
                                                <li><%= dictLanguage.Item(Session("language")&"_NotifyRe_2")%></li>
            
                                                <% 'Recipient Name %>
                                                <li>
                                                    <label for="GcReName"><%= dictLanguage.Item(Session("language")&"_NotifyRe_3")%></label>
                                                    <input class="form-control input-sm" type="text" size="20" id="GcReName" name="GcReName" value="">
                                                </li>
            
                                                <% 'Recipient Email %>
                                                <li>
                                                    <label for="GcReEmail"><%= dictLanguage.Item(Session("language")&"_NotifyRe_4")%></label>
                                                    <input class="form-control input-sm" type="text" size="20" id="GcReEmail" name="GcReEmail" value="">
                                                </li>
            
                                                <% 'Recipient Message %>
                                                <li>
                                                    <label for="GcReMsg"><%= dictLanguage.Item(Session("language")&"_NotifyRe_5")%></label>
                                                    <textarea class="form-control input-sm" cols="50" rows="3" id="GcReMsg" name="GcReMsg"></textarea>
                                                </li>
                                            </ul>
                                        
                                        </div>
                                    
                                    </div>
                                    <% '// Gift Certs - E %>
        
        
                                    <div class="pcSpacer"></div>
        
                                    <div id="OtherLoader"></div>
                                
                                    <input type="image" name="OtherSubmit" id="OtherSubmit" src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" align="absmiddle" border="0" style="display:none">
                                </form>
                            </div>
                            <%
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            '// Other - End
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
                            %>
        
        
                            <%
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            '// Gift Wrapping - Start
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
                            %>
                            <div id="GiftArea">
        
                                <div data-ng-show="Evaluate(shoppingcart.ShowGiftWrapOptions)" class="pcTable">
        
                                    <div data-ng-show="Evaluate(shoppingcart.showGiftWrapProductList)">
                                        
                                        <form id="GWForm" name="GWForm">
        
                                            <div class="pcSectionTitle"><%=dictLanguage.Item(Session("language")&"_opc_32")%>:  </div>
        
                                            <div class="pcTableRow opcGiftWrappingDetails" data-ng-show="!IsEmpty(shoppingcart.giftWrappingDetails)">                             
                                                <div class="pcSingleColumn" data-ng-bind-html="shoppingcart.giftWrappingDetails|unsafe"></div>
                                            </div>
                                            
                                            <div class="pcSpacer"></div>
                                            
                                            <div class="pcTableRow opcGiftWrappingOptionsHeader">
                                                <div class="pcLeftColumn"><span class="subTitle"><%=dictLanguage.Item(Session("language")&"_opc_33")%></span></div>
                                                <div class="pcRightColumn"><%'=dictLanguage.Item(Session("language")&"_opc_34")%></div>
                                            </div> 
                                            
                                            <div class="pcTableRow opcGiftWrappingOptionsBody" data-ng-show="shoppingcartitem.PrdCanGWchecks == 1" data-ng-repeat="shoppingcartitem in shoppingcart.shoppingcartrow">
                                                <div class="pcLeftColumn"><span class="subTitle" data-ng-bind-html="shoppingcartitem.description|unsafe">{{shoppingcartitem.description}}</span></div>
                                                <div class="pcRightColumn">                                            
                                                    <div id="GWMarker{{shoppingcartitem.id}}">
                                                    
                                                        <a data-ng-show="Evaluate(shoppingcartitem.IsGiftWrapped)" href="javascript:;" data-ng-click="GWAdd(shoppingcartitem.id, shoppingcartitem.row);">
                                                            <%=dictLanguage.Item(Session("language")&"_opc_giftWrap_1") %>
                                                        </a>
                                                        <a data-ng-show="!Evaluate(shoppingcartitem.IsGiftWrapped)" href="javascript:;" data-ng-click="GWAdd(shoppingcartitem.id, shoppingcartitem.row);">
                                                            <%=dictLanguage.Item(Session("language")&"_opc_giftWrap_2") %>
                                                        </a>
                                                        
                                                    </div>
                                                </div>
                                            </div> 
        
                                            <input type="image" data-ng-click="GWSubmit()" name="GWSubmit" id="GWSubmit" src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" alt="Submit" style="display: none">
                                            
                                        </form>
                                
                                    </div>
                                    
                                    <div data-ng-show="!Evaluate(shoppingcart.showGiftWrapProductList)">
                                    
                                        <input type="image" data-ng-click="GWSubmit()" name="GWSubmit" id="GWSubmit" src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" alt="Submit" style="display: none">
                                        
                                    </div>
                                
                                </div>
                                
                                <div data-ng-show="!Evaluate(shoppingcart.ShowGiftWrapOptions)">
        
                                    <input type="image" data-ng-click="GWSubmit()" name="GWSubmit" id="GWSubmit" src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" alt="Submit" style="display: none">  
                                                          
                                </div>
                                
                                <div class="pcSpacer"></div>
                                
                            </div>
                            <%
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            '// Gift Wrapping - End
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~													
        
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            '// Discounts - Start
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
                            %>
                            <% 
                            DiscPanelHaveInfo=0
                            HavePrdsOnSale=0
							
                            If (TurnOffDiscountCodesWhenHasSale="1") And (Not scHideDiscField="1") Then

                                Dim tmpPrdList
                                tmpPrdList=""
                                
                                For f = 1 To ppcCartIndex
                                    If pcCartArray(f,10) = 0 Then
                                        If tmpPrdList <> "" Then
                                            tmpPrdList = tmpPrdList & ","
                                        End If
                                        tmpPrdList = tmpPrdList & pcCartArray(f,0)
                                    End If
                                Next
                                        
                                If tmpPrdList = "" Then
                                    tmpPrdList = "0"
                                End If
                                
                                tmpPrdList = "(" & tmpPrdList & ")"
                                query = "SELECT idProduct FROM Products WHERE idProduct IN " & tmpPrdList & " AND pcSC_ID>0;"
                                Set rsQ = connTemp.execute(query)
                                If Not rsQ.Eof Then
                                    HavePrdsOnSale=1
                                End If
                                Set rsQ = Nothing
                                
                            End If
							
  
                            If (Not scHideDiscField="1") And (HavePrdsOnSale=0) Then
                                
                                displayDiscountCode=session("pcSFCust_discountcode")
                                
                                If pcv_strPayPanel = "1" Then
                                
                                    pdiscountDetails = ""
                                    
                                    query="SELECT pcCustSession_discountcode FROM pcCustomerSessions WHERE pcCustomerSessions.idDbSession="&session("pcSFIdDbSession")&" AND pcCustomerSessions.randomKey="&session("pcSFRandomKey")&" AND pcCustomerSessions.idCustomer="&session("idCustomer")&";"
                                    Set rs = server.CreateObject("ADODB.RecordSet")
                                    Set rs = conntemp.execute(query)
                                    If Not rs.eof Then
                                        pdiscountDetails = rs("pcCustSession_discountcode")
                                    End If									
                                    Set rs = Nothing
                                    
                                    displayDiscountCode = pdiscountDetails
                                
                                Else
                                
                                    query="SELECT discountcode FROM discounts WHERE pcDisc_Auto=1 AND active=-1 ORDER BY percentagetodiscount DESC,pricetodiscount DESC;"
                                    Set rs = server.CreateObject("ADODB.RecordSet")
                                    Set rs = conntemp.execute(query)
                                    If Not rs.Eof Then
                                        pcStrAutoDiscCode=""
                                        Do Until rs.Eof
                                            pcIntADCnt = pcIntADCnt+1
                                            If pcStrAutoDiscCode <> "" Then
                                                pcStrAutoDiscCode = pcStrAutoDiscCode & ","
                                            End If
                                            pcStrAutoDiscCode = pcStrAutoDiscCode & rs("discountcode")
                                            rs.Movenext
                                        Loop
                                    End If
                                    displayDiscountCode = displayDiscountCode & pcStrAutoDiscCode
                                    Set rs = Nothing
                                    
                                End If

                                If DiscPanelHaveInfo = 0 Then
                                    %>
                                    <div id="DiscArea">
                                        <form id="DiscForm" name="DiscForm">
                                            <%
                                            DiscPanelHaveInfo=1
                                End If
                                %>
        
                                    <% '// Discount or Gift Certificate Code %>
                                    <div class="daOPCAccSaveMsg"><span class="daOPCAccSaveHeader">Discount Code:</span> (Optional) Enter a discount code.</div>
                                        <div class="form-group">
                                            
                                            <div class="input-group">
                                                <input type="text" class="form-control daOPCDiscBox" id="DiscountCode" name="DiscountCode" value="<%=displayDiscountCode%>" size="30">
                                                <span class="input-group-btn">
                                                    <button class="btn btn-default btn-skin" type="button" data-ng-click="calculateDiscounts()" name="DiscRecal" id="DiscRecal"><%= dictLanguage.Item(Session("language")&"_css_recalculate")%></button>
                                                </span>
                                            </div>
                                            <span class="help-block"><%=dictLanguage.Item(Session("language")&"_orderverify_41")%></span>
                                        </div>
                                    
        
                                    <%End If%>
        
                                    <%
                                    If RewardsActive=1 Then
                                    
                                        If session("idCustomer")>"0" And (session("CustomerGuest")="0" Or session("CustomerGuest")="2") Then
                                        
                                            '// Add visual separator
                                            If DiscPanelHaveInfo = 1 Then 
                                                %>
                                                <!--
                                                <li>
                                                    <hr />
                                                </li>
                                                -->
                                                <%
                                            End IF
                                        
                                            'query="SELECT iRewardPointsAccrued, iRewardPointsUsed FROM Customers WHERE idcustomer=" & session("idCustomer") & " AND pcCust_Guest=0;"
                                            query="SELECT iRewardPointsAccrued, iRewardPointsUsed FROM Customers WHERE idcustomer=" & session("idCustomer") & ";"
                                            set rs=connTemp.execute(query)
                                            pcIntRewardPointsAccrued = 0
                                            pcIntRewardPointsUsed = 0
                                            if not rs.eof then
                                                pcIntRewardPointsAccrued = rs("iRewardPointsAccrued")
                                                pcIntRewardPointsUsed = rs("iRewardPointsUsed")
                                            end if
                                            set rs=nothing
                                        
                                        If RewardsActive = 1 Then
                                            opcSFIntBalance = 0
                                            If IsNull(pcIntRewardPointsAccrued) or pcIntRewardPointsAccrued="" Then 
                                                pcIntRewardPointsAccrued = 0
                                            End if
                                            If IsNull(pcIntRewardPointsUsed) or pcIntRewardPointsUsed="" Then 
                                                pcIntRewardPointsUsed = 0
                                            End if
                                            pcIntBalance = pcIntRewardPointsAccrued - pcIntRewardPointsUsed
                                            pcIntDollarValue = pcIntBalance * (RewardsPercent / 100)
                                            opcSFIntBalance = pcIntBalance
                                        End If
                                        
                                        'if customer has reward points - show total here 
                                        if opcSFIntBalance > 0 AND ((pcIntDollarValue > 0 AND session("customerType")<>"1") OR (session("customerType")="1" AND RewardsIncludeWholesale=1)) then
                                        if DiscPanelHaveInfo=0 then%>
                                    <div id="DiscArea">
                                        <form id="DiscForm" name="DiscForm">
                                            <ul class="pcListLayout">
                                                <%DiscPanelHaveInfo=1
                                        end if%>
                                                <li>
                                                    <i><span class="help-block"><%= ship_dictLanguage.Item(Session("language")&"_login_e")%> <%=opcSFIntBalance%>&nbsp;<%=RewardsLabel%> <%= ship_dictLanguage.Item(Session("language")&"_login_f")%> <%= scCurSign & money(pcIntDollarValue)%> <%= ship_dictLanguage.Item(Session("language")&"_login_g")%></span></i>
                                                </li>
                                                <% end if %>
        
                                                <% If pcIntDollarValue>0 AND ((session("customerType")<>"1") OR (session("customerType")="1" AND RewardsIncludeWholesale=1)) Then
                        if DiscPanelHaveInfo=0 then%>
                                                <div id="DiscArea">
                                                    <form id="DiscForm" name="DiscForm">
                                                        <ul class="pcListLayout">
                                                            <%DiscPanelHaveInfo=1
                                        end if%>
                                                            <li>
                                                                <div class="form-group">
                                                                    <label for="UseRewards"><%= dictRewardsLanguage.Item(Session("language")&"_order_AA")%></label>
                                                                    <div class="input-group">
                                                                        <input type="text" class="form-control" id="UseRewards" name="UseRewards" size="30" maxlength="10" value="0">
                                                                        <span class="input-group-btn">
                                                                            <button class="btn btn-default" type="button" data-ng-click="calculateRewards()" name="RewardsRecal" id="RewardsRecal"><%= dictLanguage.Item(Session("language")&"_css_recalculate")%></button>
                                                                        </span>
                                                                    </div>
                                                                </div>
                                                                <button type="button" class="pcButton pcButtonRecalculate" name="RewardsSubmit" id="RewardsSubmit" style="display: none">
                                                                    <img src="<%=pcf_getImagePath("",RSlayout("recalculate"))%>" alt="Submit" />
                                                                    <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_recalculate")%></span>
                                                                </button>     
                                                            </li>
                                                            <%end if 'if customer has reward points %>
                                                            <%end if 'Customer Logged in
                                    END IF%>
        
                                                            <%if DiscPanelHaveInfo=1 then
                      session("NoNeedStep5")="0"%>
                                                            
                                                                <div id="PaymentMessageBox" class="pcAttention" style="display: none"></div>
                                                            
                                                            
                                                            
                                                            <%else
                      session("NoNeedStep5")="1"
                      end if%>
        


								<div id="DiscountMessageBox" class="pcSuccessMessage daOPCDiscMsgOk" style="display: none"></div>

                                <% If DiscPanelHaveInfo=1 Then %>
                                        </form>
                                    </div>
                                <% End If %>
                            
                            
                            
                            <%
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            '// Discounts - End
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
        					%>
                            <div id="OPCPaySummary">
                            	<p class="daOPCPaySummaryTxt">Final Payment Summary:</p>
                            	<div class="row">
                           			<div class="col-md-12">
                            			<div class="col-sm-2">Product Subtotal:</div>
                             			<div class="col-sm-2 OPCPaySummaryRight">{{shoppingcart.dasubTotalOPCSummary}}</div>
                                        <div class="col-sm-8"></div>
                            		</div>
                           		</div>
                            	<div data-ng-if="Evaluate(shoppingcart.daOPCDiscountApplied);" class="row">
                            		<div class="col-md-12">
                            			<div class="col-sm-2">Discount:</div>
                             			<div class="col-sm-2 OPCPaySummaryRight">{{shoppingcart.daTotalPromotions}}</div>
                                        <div class="col-sm-8"></div>
                            		</div>
                            	</div>
                            	<div class="row">
                           			<div class="col-md-12">
                            			<div class="col-sm-2">Delivery:</div>
                             			<div class="col-sm-2 OPCPaySummaryRight">{{shoppingcart.dashipmentTotal}}</div>
                                        <div class="col-sm-8"></div>
                            		</div>
                            	</div>
                            	<div class="row">
                            		<div class="col-md-12">
                            			<div class="col-sm-2">VAT:</div>
                             			<div class="col-sm-2 OPCPaySummaryRight">{{shoppingcart.daOPCVAT}}</div>
                                        <div class="col-sm-8"></div>
                            		</div>
                            	</div>
                            	<div class="row">
                            		<div class="col-md-12">
                            			<div class="col-sm-2"><strong>Total Payment Due:</strong></div>
                             			<div class="col-sm-2 OPCPaySummaryRight"><strong>{{shoppingcart.total}}</strong></div>
                                        <div class="col-sm-8"></div>
                            		</div>
                            	</div>
                            </div>
                            
							<%
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            '// Payment - Start
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
                            %>
                            <div id="PayArea">
                                
                                <div id="PayNoNeed" style="display: none">
                                    <div class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_opc_40")%></div>
                                </div>
								
                                <div id="PayAreaSub">
                                    <script type=text/javascript>
                                        var NeedToUpdatePay=0;
                                        var PaymentSelected="0";
                                        var CustomPayment=0;
										var PayWithAmazon=0;
										var ShowAmzPay=0;
										<%if session("PayWithAmazon")="YES" then%>
										PayWithAmazon=1;
										ShowAmzPay=1;
										<%end if%>
                                    </script>
                                    <form id="PayForm" name="PayForm">
                                        <div class="pcSectionTitle daOPCPayTypeHeader"><% if session("PayWithAmazon") = "YES" then %><strong><%=dictLanguage.Item(Session("language")&"_opc_89")%></strong><%else%><%=dictLanguage.Item(Session("language")&"_orderverify_12")%><%end if%></div>
											
        
                                        <% if session("ExpressCheckoutPayment") = "YES" then %>
        
                                                <%
												tmpQ=""
												if session("PayWithAmazon")="YES" then
													tmpQ=" AND (payTypes.gwCode = 88)"
												else
													tmpQ=" AND (payTypes.pcPayTypes_ppab <> 1) AND (payTypes.gwCode = 999999 OR payTypes.gwCode = 46 OR payTypes.gwCode = 53 OR payTypes.gwCode = 80 OR payTypes.gwCode = 99)"
												end if
                                                query="SELECT payTypes.idPayment, payTypes.paymentDesc, payTypes.priceToAdd, payTypes.percentageToAdd, payTypes.Type, payTypes.gwCode, payTypes.paymentNickName, payTypes.sslUrl FROM payTypes WHERE (payTypes.active = - 1) " & tmpQ
                                                set rs=server.CreateObject("Adodb.recordset")
                                                set rs=conntemp.execute(query)
                                                if not rs.eof then 
                                                    tempidPayment=rs("idPayment")
													AmzidPayment=tempidPayment
                                                    temppaymentDesc=rs("paymentDesc")
                                                    temppriceToAdd=rs("priceToAdd")
                                                    temppercentageToAdd=rs("percentageToAdd")
                                                    tempType=rs("Type")
                                                    tempgwCode=rs("gwCode")
                                                    tempPaymentNickName=rs("paymentNickName")
                                                    if isNull(temppriceToAdd) OR temppriceToAdd="" then
                                                        temppriceToAdd=0
                                                    end if
                                                    if isNull(temppercentageToAdd) OR temppercentageToAdd="" then
                                                        temppercentageToAdd=0
                                                    end if
                                                    if ccur(temppriceToAdd)<>0 then
                                                        HowMuch=temppriceToAdd 
                                                        HowMuch=roundTo(HowMuch,.01)           
                                                    else
                                                        HowMuch=""
                                                    end if
                                                    if ccur(temppercentageToAdd)<>0 then
                                                        HowMuch1=temppercentageToAdd & "% of Order Total"         
                                                    else
                                                        HowMuch1=""
                                                    end if
                                                    If len(Session("DefaultIdPayment"))=0 Then
                                                        Session("DefaultIdPayment") = tempidPayment
                                                    End If
                                                   	tmpStrPay=tmpStrPay & "if (tmpid==" & tempidPayment & ") {CustomPayment=1; }" & vbcrlf
                                                    %>
                
                                                    <div class="pcShowContent">
                                                        <ul id="PayList">
                                                            <li>
																<%if session("PayWithAmazon")<>"YES" then%>
                                                                    <input type="radio" id="chkPayment" name="chkPayment" class="chkPay required" value="<%=tempidPayment%>" onClick="CheckPayType('<%=tempidPayment%>',1);" autocomplete="off">
                                                                    <%=dictLanguage.Item(Session("language")&"_opc_56")%>
                                                                    <script type="text/javascript">
                                                                        $pc(document).ready(function() { $pc('#chkPayment').click(); });			
                                                                    </script>
                                                                <%else%>
                                                                    <input type="radio" id="chkPayment" name="chkPayment" class="chkPay" value="<%=tempidPayment%>" onClick="CheckPayType('<%=tempidPayment%>',1);" checked style="display: none">
																<%end if%>
                                                            </li>
                                                        </ul>
                                                    </div>
													
													<%if session("PayWithAmazon")="YES" then%>
															<div id="walletWidgetDiv">
															</div>
															<script type=text/javascript>
							                                var AmazonBillingAgreementId="";
															function buildAmzPayment()
															{
                                                                var AmazonOptions = {
																    sellerId: '<%=pcAMZSellerID%>',
																	    design: {
																	    size : {width:'400px', height:'260px'}
																    },
                                                                    onPaymentSelect: function(orderReference) {
                                                                        // do something?
																    },
																    onError: function(error) {
																	    alert("<%=dictLanguage.Item(Session("language")&"_AmazonPay_6")%>");
																	    document.location="viewcart.asp";
																    }
																};

                                                                if (!displayShippingAddress) {
                                                                    AmazonOptions.onReady = function(billingAgreement) {
                                                                        AmazonBillingAgreementId = billingAgreement.getAmazonBillingAgreementId();
                                                                        updateAmazonBillingAgreement();
                                                                    };
                                                                    AmazonOptions.agreementType = 'BillingAgreement';
                                                                }

																new OffAmazonPayments.Widgets.Wallet(AmazonOptions).bind("walletWidgetDiv");
															}
															</script>
													<%end if%>
													
                                                <% end if %>
        
                                        <% else %>
                                        
                                            <div class="pcShowContent">
                                                
                                                <ul id="PayList">
        
                                                <%  'Get available paytypes
                                                    'If customer session
                                                    If session("customerCategory")<>0 AND session("customerCategory")<>"" then
                                                        'SB S
                                                        strAndSub = ""
                                                    
                                                        if pcIsSubscription = True Then
                                                            strAndSub = " AND pcPayTypes_Subscription = 1 ORDER BY payTypes.paymentPriority"
                                                        else
                                                            strAndSub = " ORDER BY payTypes.paymentPriority"
                                                        End if 
                                                        'SB E	
                              
                                                        query="SELECT payTypes.idPayment, payTypes.paymentDesc, payTypes.priceToAdd, payTypes.percentageToAdd, payTypes.Type, payTypes.gwCode, payTypes.paymentNickName, payTypes.sslUrl, CustCategoryPayTypes.idCustomerCategory FROM payTypes INNER JOIN CustCategoryPayTypes ON payTypes.idPayment = CustCategoryPayTypes.idPayment WHERE (payTypes.active = - 1) AND (payTypes.pcPayTypes_ppab <> 1) AND (payTypes.gwCode <> 50) AND (payTypes.gwCode <> 999999) AND (payTypes.gwCode <> 88) AND (CustCategoryPayTypes.idCustomerCategory = "&session("customerCategory")&")" & strAndSub
                                                    
                                                        set rs=server.CreateObject("Adodb.recordset")
                                                        set rs=conntemp.execute(query)
                                                    
                                                        if not rs.eof then
                                                            tmpStrPay=""
                                                        
                                                            while not rs.eof
                                                                tempidPayment=rs("idPayment")
                                                                temppaymentDesc=rs("paymentDesc")
                                                                temppriceToAdd=rs("priceToAdd")
                                                                temppercentageToAdd=rs("percentageToAdd")
                                                                tempType=rs("Type")
                                                                tempgwCode=rs("gwCode")
                                                                tempPaymentNickName=rs("paymentNickName")
                                                                if isNull(temppriceToAdd) OR temppriceToAdd="" then
                                                                        temppriceToAdd=0
                                                                end if
                                                                if isNull(temppercentageToAdd) OR temppercentageToAdd="" then
                                                                        temppercentageToAdd=0
                                                                end if
                                                                if ccur(temppriceToAdd)<>0 then
                                                                        HowMuch=temppriceToAdd 
                                                                        HowMuch=roundTo(HowMuch,.01)           
                                                                else
                                                                        HowMuch=""
                                                                end if
                                                                if ccur(temppercentageToAdd)<>0 then
                                                                        HowMuch1=temppercentageToAdd & "% of Order Total"
                                                                        'HowMuch=temppriceToAdd + (temppercentageToAdd*intCalPaymnt/100)
                                                                        'HowMuch=roundTo(HowMuch,.01)           
                                                                else
                                                                        HowMuch1=""
                                                                end if
                                                                CustomPayType=0
                                                                payURL=rs("sslURL")
                                                                if payURL<>"" then
                                                                        if Instr(UCase(payURL),UCASE("paymnta_"))=1 then
                                                                                CustomPayType=1
                                                                                payURL="opc_" & payURL
                                                                        end if
                                                                end if
                                                                If len(Session("DefaultIdPayment"))=0 Then
                                                                    Session("DefaultIdPayment") = tempidPayment
                                                                End If
                                                                if CustomPayType=1 then
                                                                        tmpStrPay=tmpStrPay & "if (tmpid==" & tempidPayment & ") {CustomPayment=1; getPayDetails(" & tempidPayment & ",'" & payURL & "'); tmpOK=1;}" & vbcrlf
                                                                end if
                                                                %>
        
                                                                <li>
                                                                    <input type="radio" name="chkPayment" class="chkPay required" value="<%=tempidPayment%>" onClick="CheckPayType('<%=tempidPayment%>',0);" autocomplete="off">
        
                                                                    <% 
                                                                    if tempgwCode="999999" OR tempgwCode="3" or tempgwCode="80" then
                                                                        
                                                                        if tempgwCode="3" or tempgwCode="80" then  %>
                                                                        
                                                                            Credit/Debit Card or PayPal
                                                                        
                                                                        <% else %>
                                                    
                                                                            <img src="<%=pcf_getImagePath("images","PayPal_mark_50x34.gif")%>" width="50" height="34">
        
                                                                            <span style="font-size: smaller;"><a href="https://www.paypal.com/us/cgibin/webscr?cmd=xpt/Marketing/popup/OLCWhatIsPayPal-outside" target="_blank">What is PayPal?</a></span>
                                                                        
                                                                        <% end if
                                                                        
                                                                    else '// if tempgwCode="999999" OR tempgwCode="3" or tempgwCode="80" then
                                                                                              
                                                                        if tempgwCode<>"7" then
                                                                            if tempType="CU" then
                                                                                response.write temppaymentDesc
                                                                            else
                                                                                response.write tempPaymentNickName
                                                                            end if
                                                                        else
                                                                            response.write temppaymentDesc
                                                                        end if 
                                                                    end if 
                                                    
                                                                    ' check to see if it's an interac online method
                                                                    if tempgwCode ="66" then
                                                                        strInteractrade = "<i>&reg; Trade-mark of Interac Inc. Used under licence.<i/>"
                                                                    end if 
                                                                    
                                                                    if HowMuch<>"" then								
                                                                        response.write " - "&scCurSign&money(HowMuch)
                                                                    else
                                                                        if HowMuch1<>"" then
                                                                            response.write " - "& HowMuch1
                                                                        end if
                                                                    end if 
                                                                    %>
                                                                    
                                                                </li>
        
                                                                <% 
                                                                rs.movenext
                                                            wend
                                                        set rs=nothing
                                                    end if
                                                    
                                                End if %>
                                                <% 'end if customer session %>
        
        
                                                <% 'SB S %>
                                                <%
                                                    strAndSub = ""
                                                    if pcIsSubscription = True Then
                                                     strAndSub = " AND pcPayTypes_Subscription = 1 ORDER by pcPayTypes_Subscription, paymentPriority"
                                                    else
                                                     strAndSub = " ORDER by paymentPriority"
                                                    End if
                                                %>
                                                <% 'SB E %>
        
                                                <%
                                                    if session("customerType")=1 then
                                                        query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd, type, gwcode, paymentNickName,sslURL FROM paytypes WHERE active=-1 AND Cbtob<>2 AND (payTypes.pcPayTypes_ppab <> 1) AND (gwcode<>50 AND gwcode<>999999 AND gwCode <> 88)" & strAndSub
                                                    else
                                                        query="SELECT idPayment,paymentDesc,priceToAdd,percentageToAdd, type, gwcode, paymentNickName,sslURL FROM paytypes WHERE active=-1 AND Cbtob=0 AND (payTypes.pcPayTypes_ppab <> 1) AND (gwcode<>50 AND gwcode<>999999 AND gwCode <> 88)" & strAndSub
                                                    end if
                                                    set rs=server.CreateObject("Adodb.recordset")
                                                    set rs=conntemp.execute(query)
            
                                                    if err.number<>0 then
                                                        call LogErrorToDatabase()
                                                        set rs=nothing
                                                        call closedb()
                                                        response.redirect "techErr.asp?err="&pcStrCustRefID
                                                    end if
        
                              if not rs.eof then
                                                        while not rs.eof
                                  tempidPayment=rs("idPayment")
                                  temppaymentDesc=rs("paymentDesc")
                                  temppriceToAdd=rs("priceToAdd")
                                  temppercentageToAdd=rs("percentageToAdd")
                                  tempType=rs("Type")
                                  tempgwCode=rs("gwCode")
                                  tempPaymentNickName=rs("paymentNickName")
                                  if isNull(temppriceToAdd) OR temppriceToAdd="" then
                                    temppriceToAdd=0
                                  end if
                                  if isNull(temppercentageToAdd) OR temppercentageToAdd="" then
                                    temppercentageToAdd=0
                                  end if
                                  if ccur(temppriceToAdd)<>0 then
                                    HowMuch=temppriceToAdd 
                                    HowMuch=roundTo(HowMuch,.01)           
                                  else
                                    HowMuch=""
                                  end if
                                  if ccur(temppercentageToAdd)<>0 then
                                    HowMuch1=temppercentageToAdd & "% of Order Total"
                                    'HowMuch=temppriceToAdd + (temppercentageToAdd*intCalPaymnt/100)
                                    'HowMuch=roundTo(HowMuch,.01)           
                                  else
                                      HowMuch1=""
                                  end if
                                  CustomPayType=0
                                  payURL=rs("sslURL")
                                  if payURL<>"" then
                                    if Instr(UCase(payURL),UCASE("paymnta_"))=1 then
                                        CustomPayType=1
                                      payURL="opc_" & payURL
                                    end if
                                  end if
                                  If len(Session("DefaultIdPayment"))=0 Then
                                    Session("DefaultIdPayment") = tempidPayment
                                  End If
                                  if CustomPayType=1 then
                                    tmpStrPay=tmpStrPay & "if (tmpid==" & tempidPayment & ") {CustomPayment=1; getPayDetails(" & tempidPayment & ",'" & payURL & "'); tmpOK=1;}" & vbcrlf
                                  end if
                                                %>
                                                <li>
                                                    <input type="radio" name="chkPayment" class="chkPay required" value="<%=tempidPayment%>" onClick="CheckPayType('<%=tempidPayment%>',0);" autocomplete="off">
                                                    <% 
                                                                    if tempgwCode="999999" OR tempgwCode="3" or tempgwCode="80" then
                                                                        if tempgwCode="3" or tempgwCode="80" then 
                                                    %>
                                                                            Credit/Debit Card or PayPal
                                                                        <% else %>
                                                    <img src="<%=pcf_getImagePath("images","PayPal_mark_50x34.gif")%>" width="50" height="34">
        
                                                    <span style="font-size: smaller;"><a href="https://www.paypal.com/us/cgibin/webscr?cmd=xpt/Marketing/popup/OLCWhatIsPayPal-outside" target="_blank">What is PayPal?</a></span>
                                                    <% end if 
                                                                    else		
                                                                        if tempgwCode<>"7" then
                                                                            if tempType="CU" then
                                                                                response.write temppaymentDesc
                                                                            else
                                                                                response.write tempPaymentNickName
                                                                            end if
                                                                        else
                                                                            response.write temppaymentDesc
                                                                        end if 
                                                                    end if
                                                
                                                                    ' chedck to see if it's an interac online method
                                                                    if tempgwCode ="66" then
                                                                        strInteractrade = "<i>&reg; Trade-mark of Interac Inc. Used under licence.<i/>"
                                                                    end if 
                                                                    if HowMuch<>"" then								
                                                                        response.write " - "&scCurSign&money(HowMuch)
                                                                    else
                                                                        if HowMuch1<>"" then
                                                                            response.write " - "& HowMuch1
                                                                        end if
                                                                    end if
                                                                    
                                                                    rs.movenext
                                                    %>    
                                                </li>
                                                <%
                                wend
                                set rs=nothing
                              end if 
                                                %>
                                            </ul>
                                        </div>
										<span class="help-block">By placing an order you agree to our <a target="_new" href="/pages/terms/">terms and conditions</a> of sale.</span>									
                                        <%
                                                end if
                                        %>
        
                                        <div id="PayFormArea" style="display: none"></div>
                                    </form>
                                    
                                    
                                    <script type=text/javascript>
                                        
                                        function CheckPayType(tmpid,ctype)
                                        {
                                            PaymentSelected=tmpid;
                                            CustomPayment=0;
                                            var tmpOK=0;
                                            $pc('.chkPay').prop('disabled', true);
                                            $pc('#PayFormArea').html(); $pc('#PayFormArea').hide();
                                            <%=tmpStrPay%>
                                            if (tmpOK==0) {NeedToUpdatePay=0;}                                            
                                            if (ctype==0) {
                                                recalculate(tmpid, '#PayLoader1', ctype, '')
                                            } else {
                                                $pc('.chkPay').prop('disabled', false);
                                            }
                                        }
                          
                                        function PreSelectPayType(tmpid)
                                        {
                                            if (tmpid!="")
                                            {
                                                var totalradio=document.getElementsByName("chkPayment").length;
                                                for (var i=0;i<totalradio;i++)
                                                {
                                                    if (document.getElementsByName("chkPayment")[i].value+""==tmpid+"")
                                                    {
                                                        document.getElementsByName("chkPayment")[i].checked=true;                                                        
                                                        CheckPayType(tmpid,1);
                                                        $pc('#PayFormArea').hide();
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                        
                                        <% If len(Session("DefaultIdPayment"))>0 Then %>
                                            var defaultPaymentSelection = <%=Session("DefaultIdPayment")%>;
                                        <% End If %>
										
                                    </script>
                                </div>
                                <div id="PayLoader" style="display: none"></div>
                                <div id="PayLoader1" style="display: none"></div>
                            </div>
                        </div>
                    <% '// END: PAYMENT INFORMATION BODY %>
        
        
        
        
        
        
        
        
        
        
                            <%
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            '// Payment Button - Start
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
                            %>
                            <div id="ButtonArea" style="display: none;">
        						
                                <div id="PlaceOrderButton" style="display: none;">
                                    <a class="pcButton pcButtonPlaceOrder btn btn-skin btn-wc updateBtn" href="javascript:;" data-ng-click="ValidateGroup1();">
                                        <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_placeorder") %></span>
                                    </a>
                                </div>
                                
                                <div data-ng-show="Evaluate(shoppingcart.IsEditOrder)">
        						<span class="help-block">By clicking Continue you agree to our <a target="_new" href="/pages/terms/">terms and conditions</a> of sale.</span>
                                    <div class="ContinueButton" style="display:none;">
                                        <a class="pcButton pcButtonUpdateCC" href="javascript:;" data-ng-click="ValidateGroup2();">
                                            <img src="<%=pcf_getImagePath("images","pc_buttons_update_cc.png")%>" alt="Update SB">
                                            <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_update_cc") %></span>
                                        </a>
                                    </div>
                                    
                                    <div id="SBSKipButton" style="display:none;">
                                        <a class="pcButton pcButtonSave" href="javascript:;" data-ng-click="SBSkip=1; ValidateGroup2();">
                                            <img src="<%=pcf_getImagePath("images","pc_buttons_save.png")%>" alt="Save SB">
                                            <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_save") %></span>
                                        </a>
                                    </div>
                                
                                </div>
        
                                <div data-ng-show="!Evaluate(shoppingcart.IsEditOrder)">
        
                                    <div class="ContinueButton" style="display:none;">
                                        <a class="pcButton pcButtonSubmit btn btn-skin btn-wc updateBtn" href="javascript:;" data-ng-click="ValidateGroup2();">
                                            <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
                                        </a>
                                    </div>
                                    
                                </div>
        
                                <div id="PlaceOrderTips" style="display: none;"></div>
                                
                            </div>
                            <%
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                            '// Payment Button - End
                            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~												
                            %>
    
    
                        </div>
                        
                  
                </div>
                <% '// START: PAYMENT INFORMATION BODY %>
                
            </div>

        </div>
        <%
        '/////////////////////////////////////////////////////////////////////////
        '// END:  ACCORDION
        '/////////////////////////////////////////////////////////////////////////
        %>

    </div>

    <script type=text/javascript>    
        var runfirst = "billing";
        if (goToPaymentPanel == true) {
            runfirst = "payment";
        }
    </script>

        <input type="image" data-ng-click="LoadPaymentPanel()" name="LoadPaymentPanel" id="LoadPaymentPanel" src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" alt="Update" style="display: none">

    <%
    Session("SF_DiscountTotal") = ""
    Session("SF_RewardPointTotal") = ""
    %>
    
</div>

<% 'FB-S            
IF session("idCustomer") = "0" then

    session("pcFBS_TurnOnOff") = 0
    query="SELECT pcFBS_TurnOnOff,pcFBS_AppID FROM pcFacebookSettings;"
    Set rs = connTemp.execute(query)
    If Not rs.Eof Then
        session("pcFBS_TurnOnOff") = rs("pcFBS_TurnOnOff")
        If IsNull(session("pcFBS_TurnOnOff")) Or session("pcFBS_TurnOnOff") = "" Then
            session("pcFBS_TurnOnOff")="0"
        End If
        session("pcFBS_AppID") = rs("pcFBS_AppID")
    End If
    Set rs = Nothing

    If session("Facebook") = "1" Or session("pcFBS_TurnOnOff") = "1" Then %>

        <script type=text/javascript>
			function statusChangeCallback(response) {
			    if (response.status === 'connected') {
			        // Logged into your app and Facebook.
			        loginPCAPI();
			    } else if (response.status === 'not_authorized') {
			    } else {
			    }
			}			
			function checkLoginState() {
			    FB.getLoginStatus(function(response) {
			    statusChangeCallback(response);
			    });
			}			
            try {
                window.fbAsyncInit = function() {
                    FB.init({
                        appId      : '<%=session("pcFBS_AppID")%>', // App ID
                        cookie     : true, // enable cookies
                        xfbml      : true,  // parse XFBML
                        version    : 'v2.5' // use version 2.5
                    });                      
                    FB.getLoginStatus(function(response) {
                        statusChangeCallback(response);
                    });

                };
                (function(d, s, id) {
                    var js, fjs = d.getElementsByTagName(s)[0];
                    if (d.getElementById(id)) return;
                    js = d.createElement(s); js.id = id;
                    js.src = "//connect.facebook.net/en_US/sdk.js";
                    fjs.parentNode.insertBefore(js, fjs);
                    }(document, 'script', 'facebook-jssdk'));
    
                    function loginPCAPI() {
                        FB.api('/me', {fields: 'email,name,first_name,last_name,id'},function(response) {
                        document.getElementById("fe").value=response.email;
                        document.getElementById("fl").value=response.id;
                        document.getElementById("ffn").value=response.first_name;
                        document.getElementById("fln").value=response.last_name;
                        document.LoginForm1.submit();
                    });
                }
            } catch (err) { }
        </script>
    <%
    End If
End If
'FB-E %>
			</div>
		</div>
	</section>
    <!-- /Section: Welcome -->
<!--#include file="footer_wrapper.asp" -->