app.directive('radioWithChangeHandler', [
    function checkboxWithChangeHandler() {
        return {
            replace: false,
            require: 'ngModel',
            scope: false,
            link: function (scope, element, attr, ngModelCtrl) {
                $pc(element).change(function () {
                    if (element[0].checked) {
                        scope.ShipArrOpts = attr.value;
                        scope.FillShipForm(0);
                    }
                });
            }
        };
    }
]);

var updateAmazonBillingAgreement = null;


// START: CONTROLLER
app.controller('OrderPreviewCtrl', ['$rootScope', '$scope', '$http', '$timeout', '$log', '$compile', 'httpRequest', 'pcService', function ($rootScope, $scope, $http, $timeout, $log, $compile, httpRequest, pcService) {

    $scope.shoppingcart = [];
    
    $scope.displayEmail = displayEmail;
    $scope.ShowVatId = ShowVatId;
    $scope.ShowSSN = ShowSSN;
    $scope.guestSession = false;


    $scope.$on('handleBroadcast', function(event, data){
        $scope.shoppingcart = data;
    });
    $scope.refresh = function () {
        pcService.getShoppingCart($scope.shoppingcartIsSaved, false);
    };
    function init() {
		if (scDispDiscCart=="1") {
            $scope.copyDiscountsCart(''); 
        }
    };    
    

    // START: Error Handling
    $scope.breakPoint = function (msg) {
        if (debuggingEnabled) {
            $log.info(msg);
        }
    }
    // START: Error Handling
    

    // START: Layout Options
    $scope.getLayoutClass = function () {
        return 'opcFormField';
    };
    // END: Layout Options



    // START: Check Expired Cart
    $scope.checkSessionExpired = function (formdata) {

        httpRequest.loadAsync('opc_cartcheck.asp', formdata).then(function (data) {

            if (data !== "OK") {
                window.location = "viewcart.asp";
            }

        });

    };
    // END: Check Expired Cart



    // START: Recalculate Cart when logging in first time...
    if ($scope.shoppingcart.IsLoggedIn) {
        $scope.recalculate("", "#Login", 1, '');
    }
    // END: Recalculate Cart when logging in first time...



    // START: Check Expired Cart
    $scope.startGuestCheckout = function (formdata) {

        if(formdata == 1) {
		
			httpRequest.loadAsync('opc_checklogin.asp', "guestcheckout=1&securityCode=" + $pc('#securityCode').val() + "&CAPTCHA_Postback=" + $pc('#CAPTCHA_Postback').val() + "&g-recaptcha-response=" + $pc('#g-recaptcha-response').val()).then(function (data) {
				
				if (data.substr(data.length-2) == "OK") {
	
					$scope.guestSession = true;
					
					$pc('#LoginOptions').hide();
					$pc('#acc1').show();
					$pc('#BillingArea').show();
			
					$scope.switchPanel('billing');
			
					document.BillingForm.billemail.focus();
	
				} else {
                    
                    $scope.guestSession = false;
	
					$scope.unblockUI('#LoginOptions');
	
					// Login Error                    
					$scope.displayMessage("#LoginMessageBox", '<img src="images/pc_icon_error_small.png" align="absmiddle"> ' + data);
	
					var CaptchaTest = $pc("#securityCode").val();
					if (CaptchaTest != null) {
						reloadCAPTCHA();
					}
					if (document.getElementById("gcaptcha")!=null) {grecaptcha.reset(widgetId);}
				}
			});
			
		} else {
            
            $scope.guestSession = true;
			
            $pc('#LoginOptions').hide();
            $pc('#acc1').show();
            $pc('#BillingArea').show();
    
            $scope.switchPanel('billing');
    
            document.BillingForm.billemail.focus();

		}
		
    };
    // END: Check Expired Cart




    // START: Secure Login
    $scope.login = function (formdata) {

        if ($pc('#loginForm').validate().form()) {

            // Block UI
            //$scope.blockUI('#LoginOptions', '<div class="pcCheckoutSubTitle"><img src="images/ajax-loader1.gif" /> Validating Login</div>');

            httpRequest.loadAsync('opc_checklogin.asp', "email=" + $pc('#email').val() + "&password=" + $pc('#password').val() + "&securityCode=" + $pc('#securityCode').val() + "&CAPTCHA_Postback=" + $pc('#CAPTCHA_Postback').val() + "&g-recaptcha-response=" + $pc('#g-recaptcha-response').val()).then(function (data) {

                if (data == "OK") {

                    //$scope.blockUI('#LoginOptions', '<div class="pcCheckoutSubTitle"><img src="images/ajax-loader1.gif" /> Loading Checkout</div>');
                    //$scope.recalculate("", "#Login", 1, '');
                    window.location = "onepagecheckout.asp";
                    return (false);

                } else {

                    $scope.unblockUI('#LoginOptions');

                    // Login Error                    
                    $scope.displayMessage("#LoginMessageBox", '<img src="images/pc_icon_error_small.png" align="absmiddle"> ' + data);

                    var CaptchaTest = $pc("#securityCode").val();
                    if (CaptchaTest != null) {
                        reloadCAPTCHA();
                    }
					if (document.getElementById("gcaptcha")!=null) {grecaptcha.reset(widgetId);}
                    //validator0.resetForm();
                }

            });
            return (false);
        }
        return (false);

    };
    // END: Secure Login



    // START: Panel Management
    $scope.panel = "billing";
    $scope.billingSubmit = false;
    $scope.shippingSubmit = false;
    $scope.ratingSubmit = false;
    $scope.shoppingcartIsSaved = false;


    $scope.showBillingEditArea = function (val) {
        if ($scope.panel !== 'billing') {
            return true;
        } else {
            return false;
        }
    };
    $scope.showShippingEditArea = function (val) {
        if (($scope.panel !== 'shipping') && ($scope.billingSubmit == true) && ($scope.shippingSubmit == true)) {
            return true;
        } else {
            return false;
        }
    };
    $scope.showRatesEditArea = function (val) {
        if (($scope.panel !== 'rates') && ($scope.billingSubmit == true) && ($scope.shippingSubmit == true) && ($scope.ratingSubmit == true)) {
            return true;
        } else {
            return false;
        }
    };
    $scope.$watch('panel', function () {

        // Unblock UI
        $scope.unblockUI('#pcShippingPanelContent');
        $scope.unblockUI('#pcBillingPanelContent');
        $scope.unblockUI('#pcRatesPanelContent');
        $scope.unblockUI('#pcPaymentPanelContent');

        // HIDE ALL: 

        // -> Hide Messages
        $scope.clearAllErrorMsg();

        // -> Hide Results
        //$pc("#ShipChargeLoadContentMsg").hide();

            // -> Hide Areas
            $pc("#BillingArea").hide();
            $pc('#ShippingArea').hide();
            $pc('#TaxContentArea').hide();
            $pc('#ShippingChargeArea').hide();


            $pc('#pcBillingPanelContent').collapse('hide')
            $pc('#pcShippingPanelContent').collapse('hide')
            $pc('#pcRatesPanelContent').collapse('hide')
            $pc('#pcPaymentPanelContent').collapse('hide')

            $pc('#opcBillingPanel').removeClass('active');
            $pc('#opcShippingPanel').removeClass('active');
            $pc('#opcRatesPanel').removeClass('active');
            $pc('#opcPaymentPanel').removeClass('active');

            switch ($scope.panel) {
                case 'billing':
                    $pc('#pcBillingPanelContent').collapse('show')
                    $pc('#opcBillingPanel').addClass('active');
                    $pc("#BillingArea").show();
                    break;
                case 'shipping':
                    $pc('#pcShippingPanelContent').collapse('show')
                    $pc('#opcShippingPanel').addClass('active');
                    $pc("#ShippingArea").show();
                    break;
                case 'rates':
                    $pc('#pcRatesPanelContent').collapse('show')
                    $pc('#opcRatesPanel').addClass('active');
                    $pc('#ShippingChargeArea').show();
                    break;
                case 'payment':
                    $pc('#pcPaymentPanelContent').collapse('show');
                    $pc('#opcPaymentPanel').addClass('active');
                    if ((PayWithAmazon == 1) && (ShowAmzPay == 1)) { buildAmzPayment(); ShowAmzPay = 0; }
					if ($pc('#TaxContentArea').length > 0) { $pc('#TaxContentArea').show(); }
                    if (defaultPaymentSelection) {
                        $("input[name=chkPayment][value=" + defaultPaymentSelection + "]").trigger("click");
                    }
                    break;
            }

        });

    $scope.switchPanel = function (panel) {
        $timeout(function () {
            $scope.panel = panel;
			$scope.refresh();
        }, 1500);
    }
	
    if (goToPaymentPanel==true) {
        $scope.billingSubmit = true;
        $scope.shippingSubmit = true;
        $scope.ratingSubmit = true;
        $scope.shoppingcartIsSaved = true; 
        if (defaultPaymentSelection) {
            PreSelectPayType(defaultPaymentSelection); 
        }
    }

    
	$scope.switchPanel(runfirst);
    // END: Panel Management 



    // START: Customer Messages
    $scope.blockUI = function (div, msg) {
        $pc.blockUI.defaults.css = {};
        $pc(div).block({
            message: '<div id="pcMain">' + msg + "</div>",
            overlayCSS: {
                backgroundColor: '#FFFFFF',
                cursor: 'wait'
            }
        });
    };
    $scope.unblockUI = function (div) {
        $pc(div).unblock();
    };
    $scope.displayMessage = function (div, msg) {
        $pc(div).html(msg);
        $pc(div).show();
    };
    $scope.clearCustErrorMsg = function (div, msg) {
        $pc(div).hide();
    };
    $scope.clearAllErrorMsg = function () {

        $pc('#LoginMessageBox').html('');
        $pc('#BillingMessageBox').html('');
        $pc('#ShippingMessageBox').html('');
        $pc('#RatingMessageBox').html('');
        $pc('#PaymentMessageBox').html('');

        $pc('#LoginMessageBox').hide();
        $pc('#BillingMessageBox').hide();
        $pc('#ShippingMessageBox').hide();
        $pc('#RatingMessageBox').hide();
        $pc('#PaymentMessageBox').hide();

    };
    // END: Customer Messages



    // START: Update Billing Address
    $scope.updateBilling = function (formdata) {

        var displayTermsWarning = false;
        
        // 'SB S
        if (pcCustomerRegAgreed == '0') {
            if ($pc("#sb_AgreeTerms").is(':checked')) {
                $scope.updTerms();
            } else {
                displayTermsWarning = true;
            }
        }
        // 'SB E

        if (pcCustomerTermsAgreed == '0') {
            if ($pc("#AgreeTerms").is(':checked')) {
                $scope.updTerms();
            } else {
                displayTermsWarning = true;
            }
        }
        
        // Terms Warning
        if (displayTermsWarning == true) {
            $scope.displayMessage("#BillingMessageBox", login_5);
            return (false);
        } else {
            $scope.clearAllErrorMsg();
        }

        // Block UI
        $scope.blockUI('#pcBillingPanelContent', '<div class="pcCheckoutSubTitle"><img src="images/ajax-loader1.gif" /> Saving Billing Information</div>');

        if ($pc('#BillingForm').validate().form()) {

            $scope.checkSessionExpired();

            httpRequest.loadAsync('opc_UpdBillAddr.asp', $pc('#BillingForm').formSerialize()).then(function (data) {

                $scope.processBillingAddress(data);

            });

        } else {

            // UnBlock UI
            $scope.unblockUI('#pcBillingPanelContent');

        }
        return (false);

    };
    // START: Update Billing Address
		
		
		
	// START: Process Billing Address
	$scope.processBillingAddress = function (data) {
		
		if (data == "SECURITY") {
			window.location = "msg.asp?message=1";
		
		} else if ((data == "ZIPLENGTH")) {
			$scope.unblockUI('#pcBillingPanelContent');
			validator1.showErrors({
				"billzip": opc_js_74
			});
			return (false);
			
		} else if (data == "ERROR") {
			$scope.unblockUI('#pcBillingPanelContent');
			$scope.displayMessage("#BillingMessageBox", opc_icon_error + opc_57);
			return (false);
			
		// BEGIN: Display Address Validation Confirmation Modal
		} else if (data == "CHECK_ADDRESS") {
			$scope.unblockUI('#pcBillingPanelContent');
			$pc('#QuickViewDialog').appendTo("body").modal({
				show: true,
				remote: 'opc_addressValidation.asp'
			});
			return (false);
		// END: Display Address Validation Confirmation Modal
		
		} else if ( (((data.indexOf("OK") >= 0) || (data.indexOf("NEW") >= 0))) || (((data == "OK") || (data == "NEW"))) ) {
			
			$scope.billingSubmit = true;
			
			// SHIPPING
			if ($scope.shoppingcart.displayShippingAddress == true) {
				
				// SHIPPING SELECT ADDRESS:  Check the options, open shipping panel
				if ($scope.shoppingcart.NeedLoadShipContent == 1) {
					$scope.switchPanel('shipping');
					$scope.getShipContents("");
											
				} else {
					$scope.switchPanel('shipping');
				}
				
			} else {
				
				// Copy billing address to shipping address when shipping address panel is not displayed in frontend.
				httpRequest.loadAsync('opc_UpdShipAddr.asp', "ShipArrOpts=-1");
								
				// SHIPPING RATES OR NONE:
				//  - Do tax, save order, and goto payment panel
				$scope.checkShippingRates("");
				
			}
			
			// SHOW PASSWORD AREA
			if ($scope.shoppingcart.allowPassword) {
				
				if (( ((data.indexOf("NEW") >= 0) || ($scope.shoppingcart.GuestCheckout == '1'))) || ( ((data == "NEW") || ($scope.shoppingcart.GuestCheckout == '1')))) {
					$pc("#PwdArea").show(); // Display Optional password selection (in payment area)
				} else {
					$pc("#PwdArea").hide(); // No optional password selection
				}
				
			}

			// CHANGE PANELS
			// $scope.switchPanel('payment');
			
		} else {
			$scope.displayMessage("#BillingMessageBox", opc_icon_error + data);
			
		}
	}
	// END: Process Billing Address
	
	
    // START: Discounts
    $scope.calculateDiscounts = function (formdata) {

        $scope.checkSessionExpired();

        httpRequest.loadAsync('opc_calculate.asp', $pc('#DiscForm').formSerialize() + '&rtype=1').then(function (data) {

            if (data == "SECURITY") {
                // Session Expired
                window.location = "msg.asp?message=1";
            } else {
                if (data.indexOf("|***|OK|***|") >= 0) {

                    var tmpArr = data.split("|***|")

                    $scope.clearAllErrorMsg();

                    $pc("#DiscountCode").val(tmpArr[2]);
                    $pc("#UseRewards").val(tmpArr[4]);
                    $pc("#OPRWarning").hide();
                    $pc('#PayArea').show();
                    OPCReady = tmpArr[8];
                    tmpchkFree = tmpArr[9];
                    if (tmpArr[8] == "NO") {
                        $pc("#OPRWarning").hide();
                        $pc("#OPRArea").show();
                    }
                    // Free Order - Start
                    if (tmpArr[8] == "YES") {
                        if (tmpArr[5] == "FREE") {
                            CustomPayment = 1;
                            NeedToUpdatePay = 0;
                            $pc("#PayNoNeed").show();
                            $pc("#PayAreaSub").hide();
                            OPCFree = 1;
                        } else {
                            $pc("#PayNoNeed").hide();
                            $pc("#PayAreaSub").show();
                            OPCFree = 0;
                        }
                        $scope.displayPlaceOrderButton();
                    }
                    // Free Order - End
                    
                } else {
                    
                    var tmpArr = data.split("|***|")

                    $scope.displayMessage("#PaymentMessageBox", '<img src="images/pc_icon_error_small.png" align="absmiddle"> ' + tmpArr[3]);

                    $pc("#DiscountCode").val(tmpArr[2]);
                    $pc("#UseRewards").val(tmpArr[4]);

                    $pc("#OPRArea").html(tmpArr[0]);
                    // Free Order - Start
                    if (tmpArr[8] == "YES") {
                        if (tmpArr[5] == "FREE") {
                            CustomPayment = 1;
                            NeedToUpdatePay = 0;
                            $pc("#PayNoNeed").show();
                            $pc("#PayAreaSub").hide();
                            OPCFree = 1;
                        } else {
                            $pc("#PayNoNeed").hide();
                            $pc("#PayAreaSub").show();
                            OPCFree = 0;
                        }
                        $scope.displayPlaceOrderButton();
                    }
                    // Free Order - End
                    $pc("#OPRWarning").hide();
                    $pc("#OPRArea").show();
                }
                $scope.GenOrderPreview(data, '', 0, '');

                $pc("#DiscountMessageBox").html('' + opc_js_42);
                $pc("#DiscountMessageBox").show();

                // $pc.growlUI('', opc_js_42); 

            }


            $scope.refresh();

        });
    }
    $scope.copyDiscountsCart = function (rtype) {

        httpRequest.loadAsync('opc_calculate.asp', 'DiscountCode=' + $pc('#DiscountCode').val() + '&fromcart=1&rtype='+rtype).then(function (data) {            
			if (data == "SECURITY") {

            } else {
                if (data.indexOf("|***|OK|***|") >= 0) {
                    var tmpArr = data.split("|***|")
                    $pc("#DiscountCode").val(tmpArr[2]); 
                }
            }
        });
    }
    // END: Discounts


    // START: Discounts [Note: this should be consolidated with above and removed]
    $scope.calculateDiscountsOnClick = function (formdata) {

        $scope.checkSessionExpired();

        httpRequest.loadAsync('opc_calculate.asp', $pc('#DiscForm').formSerialize() + '&rtype=1').then(function (data) {

            if (data == "SECURITY") {
                // Session Expired
                window.location = "msg.asp?message=1";
            } else {
                if (data.indexOf("|***|OK|***|") >= 0) {
                    var tmpArr = data.split("|***|")

                    $scope.displayMessage("#PaymentMessageBox", '<img src="images/pc_icon_success_small.png" align="absmiddle"> ' + opc_js_42);

                    $pc("#DiscountCode").val(tmpArr[2]);
                    $pc("#UseRewards").val(tmpArr[4]);
                    $pc("#OPRWarning").hide();
                    $pc('#PayArea').show();
                    OPCReady = tmpArr[8];
                    tmpchkFree = tmpArr[9];
                    if (tmpArr[8] == "NO") {

                        $pc("#OPRWarning").hide();
                        $pc("#OPRArea").show();
                    }
                    // Free Order - Start
                    if (tmpArr[8] == "YES") {
                        if (tmpArr[5] == "FREE") {
                            CustomPayment = 1;
                            NeedToUpdatePay = 0;
                            $pc("#PayNoNeed").show();
                            $pc("#PayAreaSub").hide();
                            OPCFree = 1;
                        } else {
                            $pc("#PayNoNeed").hide();
                            $pc("#PayAreaSub").show();
                            OPCFree = 0;
                        }
                    }
                    // Free Order - End
                    $scope.ValidateGroup3();
                    return;
                } else {
                    var tmpArr = data.split("|***|")

                    $scope.displayMessage("#PaymentMessageBox", '<img src="images/pc_icon_error_small.png" align="absmiddle"> ' + tmpArr[3]);

                    $pc("#DiscountCode").val(tmpArr[2]);
                    $pc("#UseRewards").val(tmpArr[4]);
                    $pc("#OPRArea").html(tmpArr[0]);
                    // Free Order - Start
                    if (tmpArr[8] == "YES") {
                        if (tmpArr[5] == "FREE") {
                            CustomPayment = 1;
                            NeedToUpdatePay = 0;
                            $pc("#PayNoNeed").show();
                            $pc("#PayAreaSub").hide();
                            OPCFree = 1;
                        } else {
                            $pc("#PayNoNeed").hide();
                            $pc("#PayAreaSub").show();
                            OPCFree = 0;
                        }
                    }
                    // Free Order - End
                    $pc("#OPRWarning").hide();
                    $pc("#OPRArea").show();
                }
            }

            $scope.unblockUI('#pcPaymentPanelContent');

            $scope.refresh();

        });
    }
    // END: Discounts [Note: this should be consolidated with above and removed]



    // Update Order Preview [This can also be consolidated with discount form]
    $scope.calculateRewards = function (data, dloader, ctype, tmpsaveOrd) {

        $scope.checkSessionExpired();

        httpRequest.loadAsync('opc_calculate.asp', $pc('#DiscForm').formSerialize() + '&rtype=1').then(function (data) {

            if (data == "SECURITY") {
                // Session Expired
                window.location = "msg.asp?message=1";
            } else {
                if (data.indexOf("|***|OK|***|") >= 0) {
                    var tmpArr = data.split("|***|")

                    $scope.clearAllErrorMsg();

                    $pc("#DiscountCode").val(tmpArr[2]);
                    $pc("#UseRewards").val(tmpArr[4]);
                    //$pc("#pcOPCtotalAmount").text(tmpArr[7]);
                    $pc("#OPRWarning").hide();
                    $pc('#PayArea').show();
                    OPCReady = tmpArr[8];
                    tmpchkFree = tmpArr[9];
                    if (tmpArr[8] == "NO") {

                        $pc("#OPRWarning").hide();
                        $pc("#OPRArea").show();
                    }
                } else {
                    var tmpArr = data.split("|***|")

                    $scope.displayMessage("#PaymentMessageBox", '<img src="images/pc_icon_error_small.png" align="absmiddle"> ' + tmpArr[3]);

                    $pc("#DiscountCode").val(tmpArr[2]);
                    $pc("#UseRewards").val(tmpArr[4]);
                    $pc("#OPRArea").html(tmpArr[0]);
                    $pc("#OPRWarning").hide();
                    $pc("#OPRArea").show();
                }

                $scope.GenOrderPreview(data, 0);

                $pc("#DiscountMessageBox").html('' + opc_js_42);
                $pc("#DiscountMessageBox").show();

            }

            $scope.refresh();

        });
    }

    // Load the Payment Panel (when coming back from other pages)
    $scope.loadPaymentPanel = function () {

        $scope.blockUI('#pcPaymentPanelContent', '<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"> ' + opc_js_79);

        if ($scope.shoppingcart.displayShippingAddress == true) {
            getShipMethod();
        }
        if (NeedLoadShipChargeContent == 1) {
            $scope.getShipContents("");
            $scope.checkShippingRates("");
        } else {
            $scope.switchPanel("payment");
            $scope.unblockUI('#pcPaymentPanelContent');	
        }

        $scope.recalculate("", "#TaxLoadContentMsg", 1, ''); // GetOrderInfo		

    }



    //*Submit Gift Wrapping Link
    $scope.GWAdd = function (pid, index) {

        $scope.checkSessionExpired();

        var tmpdata = pid;

        httpRequest.loadAsync('opc_getGiftWrap.asp?list=' + tmpdata, '{}').then(function (data) {

            if (data == "LOAD") {

                $pc('#QuickViewDialog').appendTo("body").modal({
                    show: true,
                    remote: 'opc_giftwrap.asp?list=' + tmpdata + '&index=' + index
                });
                
            } else {
                parent.tmpCheckedList = "";
            }

            $scope.refresh();

        });
    }



    //*Submit Password
    $scope.savePassword = function () {

        $scope.checkSessionExpired();

        if ($scope.shoppingcart.displayGuestFields == true) {
            if ($scope.shoppingcart.guestCheckoutStatus == 1) {
                if ((AskEnterPass == 0) && (AcceptEnterPass == 0) && ($pc("#newPass1").val() == "")) {
                    $timeout($scope.openPasswordDialog, 0);
                    return (false);
                }
            }
        }

        if ($pc('#PwdForm').validate().form()) {

            httpRequest.loadAsync('opc_createacc.asp', $pc('#PwdForm').formSerialize() + "&action=create").then(function (data) {

                if (data == "SECURITY") {
                    // Session Expired
                    window.location = "msg.asp?message=1";
                } else {
                    
                    if ((data == "OK") || (data == "REG") || (data == "OKA") || (data == "REGA")) {

                        $pc("#PwdLoader").hide();
                        $pc("#PwdWarning").hide();
                        $pc("#PwdArea").hide();

                    } else {
                        
                        $pc("#PwdLoader").html('<div class="pcErrorMessage pcClear">' + data + '</div>');
                        $pc("#PwdLoader").show();
                        validator5.resetForm();
                    }
                }

                $scope.refresh();

            });
        }
        return (false);

    }


    // Submit Gift Wrapping Area
    $scope.calculateGiftWrapping = function () {

        $scope.checkSessionExpired();

        var tmpdata = getGWPrdList();

        httpRequest.loadAsync('opc_getGiftWrap.asp?action=setup&list=' + tmpdata, "{}").then(function (data) {

            if (data == "LOAD") {
                
                $pc('#QuickViewDialog').appendTo("body").modal({
                    show: true,
                    remote: 'opc_giftwrap.asp?action=setup&list=' + tmpdata
                });
                
            } else {
                parent.tmpCheckedList = "";
                parent.updGWPrdList(); // ?
            }

            $scope.refresh();

        });
    }


    // Submit Other Order Information Form
    $scope.calculateOthers = function () {

        $scope.checkSessionExpired();

        httpRequest.loadAsync('opc_updotherinfo.asp', $pc('#OtherForm').formSerialize() + '&rtype=1').then(function (data) {

            if (data == "SECURITY") {
                // Session Expired
                window.location = "msg.asp?message=1";
            } else {
                if (data == "OK") {
                    $pc("#OtherLoader").hide();
                    $scope.ValidateGroup4();
                    return;
                } else {
                    $pc("#OtherLoader").html('<img src="images/pc_icon_error_small.png" align="absmiddle"> ' + data);
                    $pc("#OtherLoader").show();

                    validator3.resetForm();
                    $scope.unblockUI('#pcPaymentPanelContent');
                }
            }
            $scope.unblockUI('#pcPaymentPanelContent');

            $scope.refresh();

        });
    }



    // Update Order Preview
    $scope.GenOrderPreview = function (data, dloader, ctype, tmpsaveOrd) {

        if (data != "") {

            var tmpArr = data.split("|***|")
            //$pc(dloader).hide();
            $pc("#DiscountCode").val(tmpArr[2]);
            $pc("#UseRewards").val(tmpArr[4]);
            //$pc("#pcOPCtotalAmount").text(tmpArr[7]);
            $pc("#OPRArea").html(tmpArr[0]);
            $pc("#OPRWarning").hide();
            $pc("#OPRArea").hide();

            if (ctype == 1) {
                if (tmpArr[8] == "YES") {
                    if (tmpArr[5] == "FREE") {
                        CustomPayment = 1;
                        NeedToUpdatePay = 0;
                        $pc("#PayNoNeed").show();
                        $pc("#PayAreaSub").hide();
                        OPCFree = 1;
                    } else {
                        PreSelectPayType(tmpArr[6]);
                        $pc("#PayNoNeed").hide();
                        $pc("#PayAreaSub").show();
                        OPCFree = 0;
                    }
                }
            }

            OPCReady = tmpArr[8];
            tmpchkFree = tmpArr[9];

            // Order does not allow FREE Shipping
            if (OPCReady == "NO") {
                //$pc("#OPRWarning").hide();
                $pc("#OPRArea").show();
               // $scope.checkShippingRates(tmpArr[9]);
                $scope.switchPanel('shipping');
            }

            // Free Order - Start
            if (OPCReady == "YES") {
                if (tmpArr[5] == "FREE") {
                    CustomPayment = 1;
                    NeedToUpdatePay = 0;
                    $pc("#PayNoNeed").show();
                    $pc("#PayAreaSub").hide();
                    OPCFree = 1;
                } else {
                    $pc("#PayNoNeed").hide();
                    $pc("#PayAreaSub").show();
                    OPCFree = 0;
                }
            }
            // Free Order - End	

            if (OPCReady != "NO") {

                if (tmpsaveOrd == 'Y') {                 
                    $scope.saveIncompleteOrder(); // First time save of the order.  This happens in the shipping panel.
                }
            }

            $scope.displayPlaceOrderButton();
        }

    }




    // START: Update Terms
    $scope.updTerms = function (formdata) {

        $scope.checkSessionExpired();

        httpRequest.loadAsync('opc_updAgree.asp', '{}').then(function (data) {

            if (data == "OK") {
                // pass
            } else {
                $scope.displayMessage("#BillingMessageBox", '<img src="images/pc_icon_error_small.png" align="absmiddle">' + opc_js_53);
            }

        });
    }
    // END: Update Terms




    // Update Shipping Information
    $scope.updateShipping = function (formdata) {

        if ($pc('#ShippingForm').validate().form()) {
            var qtyok = true;

            $scope.checkSessionExpired();

            // Block UI
            $scope.blockUI('#pcShippingPanelContent', '<div class="pcCheckoutSubTitle"><img src="images/ajax-loader1.gif" /> Preparing Delivery Information</div>');

            httpRequest.loadAsync('opc_UpdShipAddr.asp', $pc('#ShippingForm').formSerialize()).then(function (data) {

                $scope.processShippingAddress(data);

            });

        }
        return (false);

    };
		
		
		
	// START: Process Shipping Address
	$scope.processShippingAddress = function (data) {
		if (data == "SECURITY") {
			// Session Expired
			window.location="msg.asp?message=1";
			
		} else if ((data == "ZIPLENGTH")) {
			$scope.unblockUI('#pcShippingPanelContent');
			$scope.displayMessage("#ShippingMessageBox", '<img src="images/pc_icon_error_small.png" align="absmiddle"> ' + opc_js_74);
			return (false);
			
		} else if (data == "ERROR") {
			$scope.unblockUI('#pcShippingPanelContent');
			$scope.displayMessage("#ShippingMessageBox", '<img src="images/pc_icon_error_small.png" align="absmiddle"> ' + opc_57);
			return (false);
			
		// BEGIN: Address Validation
		} else if (data == "CHECK_ADDRESS") {
			$scope.unblockUI('#pcShippingPanelContent');
			$pc('#QuickViewDialog').appendTo("body").modal({
				show: true,
				remote: 'opc_addressValidation.asp'
			});
			return (false);
		// END: Address Validation
		
		} else {

			$scope.shippingSubmit = true;
			
			if (data.indexOf("OK") >= 0) {
				
				$scope.shippingSubmit = true;
				$scope.checkShippingRates(tmpchkFree); // getShipChargeContents(tmpchkFree);
				
			} else {
				$scope.unblockUI('#pcShippingPanelContent');
				$scope.displayMessage("#ShippingMessageBox", '<img src="images/pc_icon_error_small.png" align="absmiddle"> ' + data);
			}
			
			$scope.refresh();
			
		}
	}
	// END: Process Shipping Address
	
	//Amazon Shipping
        $scope.updateAmazonShipping = function () {

            // Block UI
            $scope.blockUI('#pcShippingPanelContent', '<div class="pcCheckoutSubTitle"><img src="images/ajax-loader1.gif" /> Preparing Delivery Information</div>');


            if ((AmazonOrderReferenceId != "") && (AmzShippingSelected == 1)) {
                httpRequest.loadAsync('opc_AmzUpdShipAddr.asp', 'id=' + AmazonOrderReferenceId).then(function (data) {
                    if (data.indexOf("OK") >= 0) {
                        AmzShippingSelected = 1;
                        $scope.billingSubmit = true;
                        $scope.shippingSubmit = true;
                        $scope.checkShippingRates(tmpchkFree);
                    }
                    else {
                        $scope.unblockUI('#pcShippingPanelContent');
                        $scope.displayMessage("#ShippingMessageBox", '<img src="images/pc_icon_error_small.png" align="absmiddle"> ' + data);
                    }
                    $scope.refresh();
                });

            } else {
                $scope.unblockUI('#pcShippingPanelContent');
            }
        }

        $scope.updateAmazonBillingAgreement = function () {
            if (AmazonBillingAgreementId != "") {
                httpRequest.loadAsync('opc_AmzUpdBillAgreement.asp', 'id=' + AmazonBillingAgreementId);
            }
        };

        updateAmazonBillingAgreement = $scope.updateAmazonBillingAgreement;
	

    // Check for Shipping Methods and Rates
    $scope.checkShippingRates = function (formdata) {

        $pc("#ShipChargeLoadContentMsg").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle">' + opc_js_54); // Show the loading area
        $pc("#ShipChargeLoadContentMsg").show();
        $pc('#ShippingChargeArea').hide(); // Hide the results panel

        var tmpdata = "";
        if (tmpchkFree != "") {
            tmpdata = "pSubTotalCheckFreeShipping=" + tmpchkFree;
        }

        httpRequest.loadAsync('opc_chooseShpmnt.asp', formdata).then(function (data) {

            if (data == "SECURITY") {
                // Session Expired
                window.location = "msg.asp?message=1";
            } else {

                // Hide Msg after results are res
                var tmpArr = data.split("|*|")
                if (tmpArr[0] == "STOP") {
                    $scope.switchPanel('rates');
                    $pc("#ShipChargeLoadContentMsg").html(tmpArr[1]);
                    $pc("#ShipChargeLoadContentMsg").show();
                    return;
                }
                if (tmpArr[0] == "OK") {

                    //  NO SHIPPING:
                    //  - Shipping charges are not required.  
                    //  - Calculate Tax and redirect to Payment panel                
                    $pc('#ShippingChargeArea').html(tmpArr[1]);
                    $scope.checkTaxRates();

                } else {

                    //  SHIPPING:
                    //  - Shipping charges are required.  
                    //  - Open the Shipping panel and display the selections
                    $pc('#ShippingChargeArea').html($compile(data)($scope));
                    $pc('#ShippingChargeArea').show();
                    $scope.switchPanel('rates');
                    $scope.shippingSubmit = true;

                    var shippingPanel = "#TabbedPanelsShipping";
                    $pc(shippingPanel).tabCollapse('show');
                    $pc(shippingPanel).on("shown-accordion.bs.tabcollapse", function () {
                        $pc(shippingPanel + "-accordion .panel").first().addClass("active");
                    });

                    $pc(document).on("click", shippingPanel + "-accordion .panel-title a", function () {
                        $pc(shippingPanel + "-accordion .panel").removeClass("active");
                        $pc(this).parent().parent().parent().addClass("active");
                    });
                }

                //$pc('#ShippingChargeArea').show();
                $pc("#ShipChargeLoadContentMsg").hide();
            }

            $scope.refresh();

        });

    };




    //////////////////////////////////////////////////////
    // START: CALCULATE TAX
    //   This function is called when shipping methods are loaded 
    //	 and when shipping methods are submit (opc_chooseShpmnt.asp).
    //	 It calculates taxes and creates a new order preview.
    //////////////////////////////////////////////////////
    $scope.checkTaxRates = function (formdata) {

        $pc('#TaxContentArea').html("");
        $pc('#TaxContentArea').hide();

        httpRequest.loadAsync('opc_tax.asp', '{}').then(function (data) {

            if (data == "SECURITY") {
                // Session Expired
                window.location = "msg.asp?message=1";
            } else if ((data == "ZIPLENGTH")) {
                openDialog(button_closeHTML, '<div id="ValidationErrorMsg" class="pcErrorMessage">' + opc_js_74 + '</div>', title_validation, false);
            } else {
                if ((data == "OK") || (data == "OKA")) {

                    // TAX APPLIED
                    //   Generate new Order Preview
                    //	 Open Payment Panel
                    $scope.recalculate("", "#TaxLoadContentMsg", 1, 'Y');

                } else {

                    // TAX NOT APPLIED
                    //   Mulptiple zip codes. Must select an option
                    $pc('#TaxContentArea').html(data);
                    $pc('#TaxContentArea').show();
                    $pc("#PaymentContentArea").hide();
                    $pc("#TaxLoadContentMsg").hide();
                    $pc("#ShippingChargeArea").hide();

                    $scope.switchPanel('payment');
                    return;

                }

            }

            $scope.refresh();

        });

    };


    // Update Order with Selected Shipping Method
    $scope.updateShippingMethod = function (formdata) {

        formdata = $pc('#ShipChargeForm').formSerialize() + "&ShippingChargeSubmit=yes";

        // Block UI
        $scope.blockUI('#pcRatesPanelContent', '<div class="pcCheckoutSubTitle"><img src="images/ajax-loader1.gif" /> Preparing Payment Information</div>');

        httpRequest.loadAsync('opc_chooseShpmnt.asp', formdata).then(function (data) {

            if (data == "SECURITY") {
                // Session Expired
                window.location = "msg.asp?message=1";
                
            } else {

                if (data == "OK") {

                    $scope.clearAllErrorMsg();

                    $scope.ratingSubmit = true;

                    $scope.checkTaxRates();

                } else {
                    $scope.unblockUI('#pcRatesPanelContent');
                    $scope.displayCustErrorMsg("#RatingMessageBox", '<img src="images/pc_icon_error_small.png" align="absmiddle"> ' + data);
                }

                $scope.refresh();

            }

        });

    };

    // Recalculate Order Total

    $scope.recalculate = function (tmpid, dloader, ctype, tmpsaveOrd) {
        var tmpdata = "";
        if (tmpid != "") tmpdata = "idpayment=" + tmpid;
        if (dloader == "#TaxLoadContentMsg") {
            if ($pc("#DiscountCode") != "") {
                if (tmpdata != "") tmpdata = tmpdata + "&";
                tmpdata = tmpdata + "DiscountCode=" + $pc("#DiscountCode").val() + "&rtype=1";
            }
        }

        httpRequest.loadAsync('opc_calculate.asp', tmpdata + '&{}').then(function (data) {

            if (dloader == "#PayLoader1" || dloader == "#TaxLoadContentMsg") {
                $pc('.chkPay').prop("disabled", false);
            }

            if (data == "SECURITY") {
                // Session Expired
                window.location = "msg.asp?message=1";
            } else {
                if (data != "") {

                    $scope.GenOrderPreview(data, dloader, ctype, tmpsaveOrd);

                }
            }

            $scope.refresh();

        });

    };



    // Save Incomplete Order
    $scope.saveIncompleteOrder = function () {

        httpRequest.loadAsync('SaveOrd.asp?opc=true', '{}').then(function (data) {

            if (data == "SECURITY") {
                // Session Expired
                window.location = "msg.asp?message=1";
            } else {
                $scope.shoppingcartIsSaved = true; // Make note cart is saved for reference...
                $scope.switchPanel('payment');
            }

        });
    };




    // Check to show or hide Place Order buttons
    $scope.displayPlaceOrderButton = function () {

        OPCFinal = 0;
        if ((OPCReady == "") || (OPCReady == "NO")) {
            //$pc("#ButtonArea").hide();
            //$pc("#PlaceOrderButton").hide();
            //$pc("#ContinueButton").hide();
        } else {
            //SB-S
            if ($scope.shoppingcart.IsEditOrder == true) {
                $pc("#SBSKipButton").show();
            }
            //SB-E

            if ((CustomPayment == 1) && (NeedToUpdatePay == 1)) {

                if (CustomPayment == 1) {
                    $pc("#PlaceOrderButton").show();
                    $pc(".ContinueButton").hide();
                } else {
                    $pc("#PlaceOrderButton").hide();
                    $pc(".ContinueButton").show();
                }
                OPCFinal = 0;
                $pc("#ButtonArea").show();
            } else {

                if (CustomPayment == 1) {
                    $pc("#PlaceOrderButton").show();
                    $pc(".ContinueButton").hide();
                } else {
                    $pc("#PlaceOrderButton").hide();
                    $pc(".ContinueButton").show();
                }
                OPCFinal = 1;
                $pc("#ButtonArea").show();

            }
        }
    };




    // Load Shipping Address Drop Menu
    $scope.getShipContents = function (tmpvalue) {

        $scope.checkSessionExpired();

        $pc('#ShippingArea').hide();

        httpRequest.loadAsync('opc_genshipselect.asp', '{}').then(function (data) {

            if (data == "SECURITY") {
                // Session Expired
                window.location = "msg.asp?message=1";
            } else if (data == "ERROR") {
                $pc('#ShippingArea').hide();
            } else {
                ShipContents = data;
                $scope.generateShipDrop(tmpvalue);
            }

        });

    };


    // Generate Shipping Address Selection
    $scope.generateShipDrop = function (tmpvalue) {
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Start: Radio Buttons
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        count = 0
        $scope.ShipArrOpts = "-1";

        var strRadioBtns = '';
        strRadioBtns = strRadioBtns + '<label class="btn btn-default" for="rad_' + count + '">' + '<input radio-with-change-handler id="rad_' + count + '" type="radio" name="ShipArrOpts" value="-1" ng-model="ShipArrOpts" />' + '<span id="rad_' + count + 'g" class="pcCheckBox"></span>' + opc_20 + '</label>';

        if (HaveGRAddress == 1) {
            count = count + 1
            strRadioBtns = strRadioBtns + '<label class="btn btn-default" for="rad_' + count + '">' + '<input radio-with-change-handler id="rad_' + count + '" type="radio" name="ShipArrOpts" value="-2" ng-model="ShipArrOpts" />' + '<span id="rad_' + count + 's" class="pcCheckBox"></span>' + opc_20 + '</label>';
        }

        var tmpHaveShipAddr = 0;
        var SelectedShip = "";
        var SelectedC = "";

        if (ShipContents != "") {
            var tmpShipList = ShipContents.split("|$|");
            if (tmpShipList.length > 0) {
                SelectedShip = tmpShipList[0];
                for (var i = 1; i < tmpShipList.length; i++) {                    
                    if (tmpShipList[i] != "") {
                        tmpShipRe = tmpShipList[i].split("|*|")
                        count = count + 1
                        SelectedShip = tmpShipRe[0];
                        strRadioBtns = strRadioBtns + '<label class="btn btn-default" for="rad_' + count + '">' + '<input radio-with-change-handler id="rad_' + count + '" type="radio" name="ShipArrOpts" value="' + tmpShipRe[0] + '" ng-model="ShipArrOpts" />' + '<span id="rad_' + count + 'g" class="pcCheckBox"></span>' + tmpShipRe[1] + '</label>';
                        tmpHaveShipAddr = tmpHaveShipAddr + 1;
                        if (SelectedShip == tmpShipRe[0]) SelectedC = count;
                    }
                }
            }
        }
        if ((tmpHaveShipAddr < 2) || ($scope.shoppingcart.newCustomer == 0))
            if ($scope.shoppingcart.CanCreateNewShip == 1) {
                count = count + 1
                strRadioBtns = strRadioBtns + '<label class="btn btn-default" for="rad_' + count + '">' + '<input radio-with-change-handler id="rad_' + count + '" type="radio" name="ShipArrOpts" value="ADD" ng-model="ShipArrOpts" />' + '<span id="rad_' + count + 'g" class="pcCheckBox"></span>' + opc_js_38 + '</label>';
                if (SelectedShip == "ADD") SelectedC = count;
            }

        injector = $compile(strRadioBtns)($scope);
        $pc("#opcShippingRadios").html(injector);

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // End: Radio Buttons
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        if (PPSAID) {
            $scope.ShipArrOpts = SelectedShip;
            $scope.checkRadio('#rad_1');            
            $scope.FillShipForm(0);            
        } else {
            if (SelectedShip != "") {
                if (SelectedShip == "-1") SelectedC = 0;
                if (SelectedShip == "-2") SelectedC = 1;
                $scope.ShipArrOpts = SelectedShip;
                $scope.checkRadio('#rad_' + SelectedC);
                $scope.FillShipForm(1);
            } else {
                $scope.ShipArrOpts = "-1";
                $scope.checkRadio('#rad_0');
                $scope.FillShipForm(1);
            }
        }

        $pc('#ShippingArea').show();

    };


    // Check Default ShipForm Values
    $scope.checkRadio = function (val) {        
        $pc(val).attr('checked', 'checked');
        $pc(val).prop('checked', true);
        $pc(val).parent().addClass('active');
        $pc('.pcCheckBox').removeClass("glyphicon glyphicon-ok");
        $pc('#opcShippingRadios :input:checked').parent().find('span').addClass("glyphicon glyphicon-ok");
    };

    // Fill information to Shipping Form
    $scope.FillShipForm = function (cback) {

        tmpvalue = $scope.ShipArrOpts

        if ((tmpvalue == "-1") || (tmpvalue == "") || (tmpvalue == "-2") || (tmpvalue == -1) || (tmpvalue == -2)) {

            $pc("#shippingAddressArea").hide();
            if (tmpvalue == "-2") {
                if (HaveShipTypeArea == 1 || HaveShipTypeAreaBilling == 1) {
                    $pc("#shipAddrTypeArea").hide();
                    $pc("#billAddrTypeArea").hide();
                }
            }
            
        } else {

            if (tmpvalue == "ADD") {
                
                $pc("#shipnickname").val("");
                $pc("#shipfname").val("");
                $pc("#shiplname").val("");
                $pc("#shipcompany").val("");
                $pc("#shipaddr").val("");
                $pc("#shipaddr2").val("");
                $pc("#shipcity").val("");
                $pc("#shipzip").val("");
                $pc("#shipprovince").val("");
                $pc("#shipstate").val("");
                $pc("#shipcountry").val($pc("#billcountry").val());
                $pc("#shipphone").val("");
                $pc("#shipfax").val("");
                $pc("#shipemail").val("");
                SwitchStates('ShippingForm',document.ShippingForm.shipcountry.options.selectedIndex, 'shipcountry', 'shipstate', 'shipprovince', $pc("#billstate").val(), '');
                if (HaveShipTypeArea == 1) {
                    document.ShippingForm.pcAddressType[0].checked = true
                };
                $pc("#shipnicknameArea").show();
                $pc("#shipnameArea").show("");
                $pc("#shippingAddressArea").show();
            
            } else {

                if (ShipContents != "") {
                    var tmpShipList = ShipContents.split("|$|");
                    if (tmpShipList.length > 0) {
                        for (var i = 0; i < tmpShipList.length; i++) {
                            if (tmpShipList[i] != "") {
                                tmpShipRe = tmpShipList[i].split("|*|");
                                if (tmpShipRe[0] == tmpvalue) {

                                    $pc("#shipnickname").val(tmpShipRe[1]);
                                    $pc("#shipfname").val(tmpShipRe[2]);
                                    $pc("#shiplname").val(tmpShipRe[3]);
                                    $pc("#shipemail").val(tmpShipRe[4]);
                                    $pc("#shipphone").val(tmpShipRe[5]);
                                    $pc("#shipfax").val(tmpShipRe[6]);
                                    $pc("#shipcompany").val(tmpShipRe[7]);
                                    $pc("#shipaddr").val(tmpShipRe[8]);
                                    $pc("#shipaddr2").val(tmpShipRe[9]);
                                    $pc("#shipcity").val(tmpShipRe[10]);
                                    $pc("#shipprovince").val(tmpShipRe[11]);
                                    $pc("#shipstate").val(tmpShipRe[12]);
                                    $pc("#shipzip").val(tmpShipRe[13]);
                                    $pc("#shipcountry").val(tmpShipRe[14]);

                                    SwitchStates('ShippingForm', document.ShippingForm.shipcountry.options.selectedIndex, 'shipcountry', 'shipstate', 'shipprovince', tmpShipRe[12], '');

                                    if (tmpShipRe[15] == "") {
                                        tmpShipRe[15] = "1";
                                    }

                                    if (HaveShipTypeArea == 1) {
                                        if (tmpShipRe[15] == "0") {
                                            document.ShippingForm.pcAddressType[1].checked = true
                                        } else {
                                            document.ShippingForm.pcAddressType[0].checked = true
                                        }
                                    }

                                    if (tmpShipRe[0] == 0) {
                                        $pc("#shipnicknameArea").hide();
                                        $pc("#shipnameArea").hide("");
                                    } else {
                                        $pc("#shipnicknameArea").show();
                                        $pc("#shipnameArea").show("");
                                    }

                                    $pc("#shippingAddressArea").show();
                                }
                            }
                        }
                    }
                }
            }
        }
        if (tmpvalue == "") {
            if (HaveShipTypeArea == 1 || HaveShipTypeAreaBilling == 1) {
                $pc("#shipAddrTypeArea").hide();
                $pc("#billAddrTypeArea").hide();
            }
            if ($scope.shoppingcart.HaveDeliveryArea == 1) $pc("#shipDeliveryArea").hide();
        } else {
            if ((tmpvalue == "-1")) {
                if (cback != 1)
                    if (HaveShipTypeArea == 1) {
                        document.ShippingForm.pcAddressType[0].checked = true
                    };
            }
            if (tmpvalue != "-2") {
                if (HaveShipTypeArea == 1) {
                    $pc("#shipAddrTypeArea").show();
                }
                if (HaveShipTypeAreaBilling == 1) {
                    $pc("#billAddrTypeArea").show();
                }
            }
            if ($scope.shoppingcart.HaveDeliveryArea == 1) $pc("#shipDeliveryArea").show();
        }
    };


    $scope.billingAddressType = function () {
        if ($scope.shoppingcart.billingAddressTypeArea == 1) {
            HaveShipTypeAreaBilling = 1;
            return true;
        } else {
            HaveShipTypeAreaBilling = 0;
            return false;
        }
    };


    $scope.shippingAddressType = function () {
        if ($scope.shoppingcart.shippingAddressTypeArea == 1) {
            HaveShipTypeArea = 1;
            return true;
        } else {
            HaveShipTypeArea = 0;
            return false;
        }
    };
    
     $scope.displayDeliveryArea = function () {
        if ($scope.shoppingcart.HaveDeliveryArea == 1) {
            return true;
        } else {
            return false;
        }
    };

    $scope.Evaluate = function (val) {
        if (val == 'true' || val == true) {
            return true;
        } else {
            return false;
        }
    };

    $scope.IsEmpty = function (val) {
        if (val) {
            return false;
        } else {
            return true;
        }
    };


    $scope.CheckQuantityMins = function (a, b, c, d, e, f) {
        checkproqtyNew(a, b, c, d, e, f);
    };




    //* Submit All Payment Features
    $scope.ValidateGroup1 = function (tmpID) {

        if ($scope.shoppingcart.ExpressCheckoutInUse == false) {

            if ($pc("#PaySubmit").length > 0) {
                
                $scope.breakPoint(OPCFree);
                
                if (OPCFree == 0) {
                    try {
                        $pc('#PaySubmit').click();
                    } catch (err) {
                        $scope.ValidateGroup2();
                    }
                } else {
                    $scope.ValidateGroup2();
                }
            } else {
                $scope.ValidateGroup2();
            }

        } else {
            $scope.ValidateGroup2();
        }

    }


    $scope.ValidateGroup2 = function (tmpID) {

        $scope.breakPoint('ValidateGroup2');
        
        // validate tax
        if ($pc("#TaxChoice").length > 0) {
            if ($pc('#TaxForm').validate().form()) {               
                if ($pc('input[name=chkPayment]:checked').length) {                
                } else { 
                    if (OPCFree == 0) {      
                        //$pc("#TaxLoader").html('<div class="pcErrorMessage pcClear"> ' + 'Please select a county and click "Continue" below your selection.' + '</div>');
                        //$pc("#TaxLoader").show();
                        //return false;
                    }
                }  
            } else {               
               //return false; 
            }
        } 
        
        // validate radio buttons
        $pc("#PayLoader").hide();
        if ($pc('input[name=chkPayment]:checked').length) {
        
        } else {
            
            if (OPCFree == 0) { 
                $pc("#PayLoader").html('<div class="pcErrorMessage pcClear"> ' + 'You must select a payment method.' + '</div>');
                $pc("#PayLoader").show();
                return false; // stop whatever action would normally happen
            }
        }

        if ($scope.shoppingcart.displayGuestFields = true) {
            if ($scope.shoppingcart.guestCheckoutStatus == 1 && $scope.shoppingcart.DisplayOptionalPassword == true) {
                if ((AskEnterPass == 0) && (AcceptEnterPass == 0) && ($pc("#newPass1").val() == "")) {
                    $timeout($scope.openPasswordDialog, 0);
                    return (false);
                }
            }
        }

        $timeout($scope.openSavingDialog, 0);

        if ($pc("#DiscSubmit").length > 0) {
            $scope.calculateDiscountsOnClick() // Calculate discounts and then proceed to ValidateGroup3
        } else {
            $scope.ValidateGroup3();
        }

    }


    $scope.ValidateGroup3 = function (tmpID) {

        $scope.breakPoint('ValidateGroup3');

        if ($pc("#OtherSubmit").length > 0) {
            $scope.calculateOthers() // Calculate others and then proceed to ValidateGroup4
        } else {
            $scope.ValidateGroup4();
        }
    }


    $scope.ValidateGroup4 = function (tmpID) {

        $scope.breakPoint('ValidateGroup4');

        $scope.displayPlaceOrderButton();

        if (OPCFinal == 1) {
            setTimeout('window.location="SaveOrd.asp?sbskip="+SBSkip;', 2000);
        } else {
            $scope.unblockUI('#pcPaymentPanelContent');
        }
    }


    $scope.openSavingDialog = function () {
        $scope.blockUI('#pcPaymentPanelContent', '<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"> ' + opc_js_81);
    }

    $scope.openPasswordDialog = function () {
        AcceptEnterPass = 0;        
        openDialog(button_askPasswordHTML, message_askPassword, title_askPassword, false); 
        AskEnterPass = 1;
    }
    
    // Initialize Pageload
    init();
  
}]);
// END: CONTROLLER


$pc(document).ready(function () {
    $pc(".pcButtonSet").button(); 
});


// Recalculate from Modal Window
function recalculate(tmpid, dloader, ctype, tmpsaveOrd) {
    var doc = angular.element($pc("#OrderPreviewCtrl")).scope();
    doc.recalculate(tmpid, dloader, ctype, tmpsaveOrd);
};

// Validate from Modal Window
function ValidateGroup2(val) {
    var doc = angular.element($pc("#OrderPreviewCtrl")).scope();
    doc.ValidateGroup2();
};


// 'SB S - Agreement Functions
function closeRegDialog() {
    $pc("#RegTermsDialog").dialog("close");
}
// 'SB E


function updGWPrdList() {
    if (tmpCheckedList != "") {
        var tmpArr = tmpCheckedList.split(",")
        var i = 0;
        for (var i = 1; i <= $scope.shoppingcart.totalGiftWrapProducts; i++) {
            eval("document.GWForm.PrdGW" + i).checked = false;
            for (var j = 0; j < tmpArr.length; j++) {
                if (tmpArr[j] != "") {
                    if (eval("document.GWForm.PrdGW" + i).value == tmpArr[j]) {
                        eval("document.GWForm.PrdGW" + i).checked = true;
                    }
                }
            }
        }

    } else {
        var i = 0;
        for (var i = 1; i <= $scope.shoppingcart.totalGiftWrapProducts; i++) {
            eval("document.GWForm.PrdGW" + i).checked = false;
        }
    }
}

//*Prepare Contents of Pay Details Form
function getPayDetails(tmpid, tmpURL) {

    var doc = angular.element($pc("#OrderPreviewCtrl")).scope();
    doc.checkSessionExpired();

    $pc('#PayFormArea').html("");
    $pc('#PayFormArea').hide();
    $pc.ajax({
        type: "POST",
        url: tmpURL,
        data: "idpayment=" + tmpid,
        timeout: 45000,
        success: function (data, textStatus) {
            if (data == "SECURITY") {
                // Session Expired
                window.location = "msg.asp?message=1";
            } else {
                $pc('#PayFormArea').html(data);
                $pc('#PayFormArea').show();
                $pc("#PayLoader").hide();
                $pc('.chkPay').prop('disabled', false);
            }
        }
    });
}




function pcf_LoadPaymentPanel() {
    $pc('#LoadPaymentPanel').click();
}

function finishFunc(elmentObj, FadeObj) {
    elmentObj.style.display = "none";
}


function getGWPrdList() {
    var tmpdata = "";
    var i = 0;
    for (var i = 1; i <= $scope.shoppingcart.totalGiftWrapProducts; i++) {
        if (eval("document.GWForm.PrdGW" + i).checked == true) {
            if (tmpdata != "") tmpdata = tmpdata + ",";
            tmpdata = tmpdata + eval("document.GWForm.PrdGW" + i).value;
        }
    }
    return (tmpdata);
}

function optwin2(fileName) {
    myFloater = window.open('', 'myWindow', 'scrollbars=yes,status=no,width=400,height=300');
    myFloater.location.href = fileName;
}

function copyfromBillAddr() {
    //$pc("#shipnickname").val("");
    $pc("#shipfname").val($pc("#billfname").val());
    $pc("#shiplname").val($pc("#billlname").val());
    if ($pc("#billemail").length) {
        $pc("#shipemail").val($pc("#billemail").val());
    } else {
        $pc("#shipemail").val($pc("#billemail2").val());
    }
    $pc("#shipphone").val($pc("#billphone").val());
    $pc("#shipfax").val($pc("#billfax").val());
    $pc("#shipcompany").val($pc("#billcompany").val());
    $pc("#shipaddr").val($pc("#billaddr").val());
    $pc("#shipaddr2").val($pc("#billaddr2").val());
    $pc("#shipcity").val($pc("#billcity").val());
    $pc("#shipprovince").val($pc("#billprovince").val());
    $pc("#shipstate").val($pc("#billstate").val());
    $pc("#shipzip").val($pc("#billzip").val());
    $pc("#shipcountry").val($pc("#billcountry").val());
    SwitchStates('ShippingForm', document.ShippingForm.shipcountry.options.selectedIndex, 'shipcountry', 'shipstate', 'shipprovince', $pc("#billstate").val(), '');
}

function switchZipName1(tmpValue) {
    if (tmpValue == "CA") {
        $pc("#billzipname").html(opc_16a);
    } else {
        $pc("#billzipname").html(opc_16);
    }
}

function switchZipName2(tmpValue) {
    if (tmpValue == "CA") {
        $pc("#shipzipname").html(opc_16a);
    } else {
        $pc("#shipzipname").html(opc_16);
    }
}

function newWindow(file, window) {
    msgWindow = open(file, window, 'resizable=no,width=530,height=150');
    if (msgWindow.opener == null) msgWindow.opener = self;
}

function togglediv(id) {
    $pc("#" + id).slideToggle('fast');
}

function win(fileName) {
    myFloater = window.open('', 'myWindow', 'scrollbars=yes,status=no,width=300,height=250')
    myFloater.location.href = fileName;
}

function noPass() {						
    AcceptEnterPass=0;
    $pc('#PwdLoader').hide();
    $pc('#PwdWarning').hide();
    $pc('#PwdArea').hide();
}
function addPass() {						
    AcceptEnterPass=1;
    $pc('#newPass1').focus();
}	

$pc('.btn-group').change(function () {
    $pc('.pcCheckBox').removeClass("glyphicon glyphicon-ok");
    $pc('#opcShippingRadios :input:checked').parent().find('span').addClass("glyphicon glyphicon-ok");
});