var validator0
var validator1
var validator2
var validator3
var validator4
var validator5

$pc(document).ready(function () {

    jQuery.validator.setDefaults({
        success: function (element) {
            // $pc(element).parent("div").addClass("has-success")
        }
    });

    jQuery.validator.addMethod("credit_card", function (value, element) {
        return this.optional(element) || /^[0-9\ ]+$/i.test(value);
    }, opc_js_1);



    ///////////////////////////////////////////////////////////////////////
    // START: LOGIN VALIDATION
    ///////////////////////////////////////////////////////////////////////
    validator0 = $pc("#loginForm").validate({
        rules: {
            email: {
                required: true,
                email: true
            },
            password: {
                required: true
            }
        },
        messages: {
            email: {
                required: opc_js_2,
                email: opc_js_3
            },
            password: {
                required: opc_js_4
            }
        }
    })
    ///////////////////////////////////////////////////////////////////////
    // END: LOGIN VALIDATION
    ///////////////////////////////////////////////////////////////////////




    ///////////////////////////////////////////////////////////////////////
    // START: BILLING PANEL VALIDATION
    ///////////////////////////////////////////////////////////////////////
    validator1 = $pc("#BillingForm").validate({
        errorElement: "span",
        errorClass: "help-block",
        highlight: function (element, errorClass, validClass) {
            $pc(element).closest('.form-group').addClass('has-error');
        },
        unhighlight: function (element, errorClass, validClass) {
            $pc(element).closest('.form-group').removeClass('has-error');
        },
        errorPlacement: function (error, element) {
            if (element.parent('.input-group').length || element.prop('type') === 'checkbox' || element.prop('type') === 'radio') {
                error.insertAfter(element.parent());
            } else {
                error.insertAfter(element);
            }
        }, 
        rules: {
            billfname: "required",
            billlname: "required",
            billaddr: "required",
            billcity: "required",
            billcountry: "required",
            billphone: {
                required: function (element) {
                    return IsBillingPhoneReq;
                }
            },
            billVATID: {
                required: function (element) {
                    return IsVatIdRequired;
                }
            },
            billSSN: {
                required: function (element) {
                    return IsSSNRequired;
                }
            },
            billzip: {
                required: function (element) {
                    var str = document.BillingForm.billcountry.value;
                    if (CountriesRequireZipCode.indexOf(str + ",") >= 0) {
                        return (true)
                    } else {
                        return (false)
                    }
                }
            },
            billstate: {
                required: function (element) {
                    var str = document.BillingForm.billcountry.value;
                    if (CountriesRequireStateCode.indexOf(str + ",") >= 0) {
                        return (true)
                    } else {
                        return (false)
                    }
                }
            },

            // Guest Checkout
            billemail: {
                required: validateGuestInfo,
                email: true,
                remote: {
                    url: "opc_checkEmail.asp",
                    type: "POST"
                }
            },
            billpass: {
            	required: function (element) {
            		if (scGuestCheckoutOpt == '2') {
            			return (true)
            		} else {
            			return (false)
            		}
            	}
            },
            billrepass: {
                required: function (element) {
                    if (scGuestCheckoutOpt == '2') {
                        return (true)
                    } else {
                        return (false)
                    }
                },
                equalTo: "#billpass"
            },


            billprovince: {
                required: function (element) {
                    var str = document.BillingForm.billcountry.value;
                    if (CountriesRequireProvince.indexOf(str + ",") >= 0) {
                        return (true)
                    } else {
                        return (false)
                    }
                }
            }

        },
        messages: {
            billfname: {
                required: opc_js_8
            },
            billlname: {
                required: opc_js_9
            },
            billaddr: {
                required: opc_js_10
            },
            billcity: {
                required: opc_js_11
            },
            billcountry: {
                required: opc_js_12
            },
            billphone: {
                required: opc_js_16
            },
            billVATID: {
                required: Custmoda_27
            },
            billSSN: {
                required: Custmoda_25
            },
            billzip: {
                required: opc_js_13
            },
            billstate: {
                required: opc_js_15
            },
            billprovince: {
                required: opc_js_63
            },
            billemail: {
                required: opc_js_2,
                email: opc_js_3,
                remote: opc_js_5a
            },
            billpass: {
                required: opc_js_4
            },
            billrepass: {
                required: opc_js_47,
                equalTo: opc_js_48
            }

        }
    });


    $pc('#BillingCancel').click(function () {

        $scope.clearAllErrorMsg();

        $pc("#BillingArea").hide();
        $pc('#LoginOptions').show();
        $pc('#acc1').hide();

    });


    $pc('#ViewTerms').click(function () {
        $pc('#TermsDialog').appendTo("body").modal('show');
        return;
    });
    $pc('#sb_ViewTerms').click(function () {
        $pc('#sb_TermsDialog').appendTo("body").modal('show');
        return;
    });
    ///////////////////////////////////////////////////////////////////////
    // END: BILLING PANEL VALIDATION
    ///////////////////////////////////////////////////////////////////////




    ///////////////////////////////////////////////////////////////////////
    // START: SHIPPING PANEL VALIDATION
    ///////////////////////////////////////////////////////////////////////
    var validator2 = $pc("#ShippingForm").validate({
        errorElement: "span",
        errorClass: "help-block",
        highlight: function (element, errorClass, validClass) {
            $pc(element).closest('.form-group').addClass('has-error');
        },
        unhighlight: function (element, errorClass, validClass) {
            $pc(element).closest('.form-group').removeClass('has-error');
        },
        errorPlacement: function (error, element) {
            if (element.parent('.input-group').length || element.prop('type') === 'checkbox' || element.prop('type') === 'radio') {
                error.insertAfter(element.parent());
            } else {
                error.insertAfter(element);
            }
        }, 
        rules: {
            shipfname: {
                required: function (element) {
                    return (($pc("input[name='ShipArrOpts']:checked").val() != "-1") && ($pc("input[name='ShipArrOpts']:checked").val() != "-2"))
                }
            },
            shiplname: {
                required: function (element) {
                    return (($pc("input[name='ShipArrOpts']:checked").val() != "-1") && ($pc("input[name='ShipArrOpts']:checked").val() != "-2"))
                }
            },
            shipaddr: {
                required: function (element) {
                    return (($pc("input[name='ShipArrOpts']:checked").val() != "-1") && ($pc("input[name='ShipArrOpts']:checked").val() != "-2"))
                }
            },
            shipcity: {
                required: function (element) {
                    return (($pc("input[name='ShipArrOpts']:checked").val() != "-1") && ($pc("input[name='ShipArrOpts']:checked").val() != "-2"))
                }
            },
            shipcountry: {
                required: function (element) {
                    return (($pc("input[name='ShipArrOpts']:checked").val() != "-1") && ($pc("input[name='ShipArrOpts']:checked").val() != "-2"))
                }
            },
            shipzip: {
                required: function (element) {
                    if (($pc("input[name='ShipArrOpts']:checked").val() != "-1") && ($pc("input[name='ShipArrOpts']:checked").val() != "-2")) {
                        var str = document.ShippingForm.shipcountry.value;
                        if (CountriesRequireZipCode.indexOf(str + ",") >= 0) {
                            return (true)
                        } else {
                            return (false)
                        }
                    } else {
                        return (false)
                    }
                }
            },
            shipstate: {
                required: function (element) {
                    if (($pc("input[name='ShipArrOpts']:checked").val() != "-1") && ($pc("input[name='ShipArrOpts']:checked").val() != "-2")) {
                        var str = document.ShippingForm.shipcountry.value;
                        if (CountriesRequireStateCode.indexOf(str + ",") >= 0) {
                            return (true)
                        } else {
                            return (false)
                        }
                    } else {
                        return (false)
                    }
                }
            },
            shipprovince: {
                required: function (element) {
                    if (($pc("input[name='ShipArrOpts']:checked").val() != "-1") && ($pc("input[name='ShipArrOpts']:checked").val() != "-2")) {
                        var str = document.ShippingForm.shipcountry.value;
                        if (CountriesRequireProvince.indexOf(str + ",") >= 0) {
                            return (true)
                        } else {
                            return (false)
                        }
                    } else {
                        return (false)
                    }
                }
            },
			shipphone: {
                required: function (element) {
                    return IsShippingPhoneReq;
                }
            },
            shipemail: {
                required: function (element) {
                    return (($pc("input[name='ShipArrOpts']:checked").val() != "-1") && ($pc("input[name='ShipArrOpts']:checked").val() != "-2"))
                },
                email: true
            }
        },
        messages: {
            shipfname: {
                required: opc_js_25
            },
            shiplname: {
                required: opc_js_26
            },
            shipaddr: {
                required: opc_js_27
            },
            shipcity: {
                required: opc_js_28
            },
            shipcountry: {
                required: opc_js_29
            },
            shipzip: {
                required: opc_js_30
            },
            shipstate: {
                required: opc_js_32
            },
            shipprovince: {
                required: opc_js_64
            },
			shipphone: {
                required: opc_js_16
            },
            shipemail: {
                email: opc_js_3
            },
            DF1: {
                required: opc_js_35
            },
            TF1: {
                required: opc_js_36
            }

        }
    });
    ///////////////////////////////////////////////////////////////////////
    // END: SHIPPING PANEL VALIDATION
    ///////////////////////////////////////////////////////////////////////




    ///////////////////////////////////////////////////////////////////////
    // START: PAYMENT PANEL VALIDATION
    ///////////////////////////////////////////////////////////////////////

    // Other Form
    var validator3 = $pc("#OtherForm").validate({
        rules: {
            shipemail: {
                email: true
            }
        },
        messages: {
            GcReEmail: {
                email: opc_js_3
            }
        }
    });

    // Payment Form
    var validator4 = $pc("#PayForm").validate({
        errorElement: "span",
        errorClass: "help-block",
        highlight: function (element, errorClass, validClass) {
            $pc(element).closest('.form-group').addClass('has-error');
        },
        unhighlight: function (element, errorClass, validClass) {
            $pc(element).closest('.form-group').removeClass('has-error');
        },
        errorPlacement: function (error, element) {
            if (element.parent('.input-group').length || element.prop('type') === 'checkbox' || element.prop('type') === 'radio') {
                error.insertAfter(element.parent());
            } else {
                error.insertAfter(element);
            }
        }, 
        rules: {
            cardNumber: {
                required: true,
                credit_card: true,
                remote: {
                    url: "opc_checkCC.asp",
                    type: "POST",
                    data: {
                        cardType: function () {
                            return $pc("#cardType").val();
                        }
                    }
                }
            }
        },
        messages: {
            cardNumber: {
                required: opc_js_44,
                minlength: opc_js_45,
                remote: opc_js_46
            }
        }
    });

    // Password Form
    var validator5 = $pc("#PwdForm").validate({
        rules: {
            newPass1: {
                required: true,
				remote: {
					type: 'POST',
					url: "checkPass.asp",
					data: {
						passtype: "R",
						pass: function () {
                        	return $pc("#newPass1").val();
                        }
                        },
					dataFilter: function(data) {
 						var myjson = JSON.parse(data);
 						if(myjson.isError == "true") {
 							return "\"" + myjson.errorMessage + "\"";
 						} else {
							return true;
						}
					}
                }
            },
            newPass2: {
                required: true,
                equalTo: "#newPass1"
            }
        },
        messages: {
            newPass1: {
                required: opc_js_4
            },
            newPass2: {
                required: opc_js_47,
                equalTo: opc_js_48
            }
        }
    })

    ///////////////////////////////////////////////////////////////////////
    // START: PAYMENT PANEL VALIDATION
    ///////////////////////////////////////////////////////////////////////  

});