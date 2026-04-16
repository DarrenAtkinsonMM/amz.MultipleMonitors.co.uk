
function mediasize() {
    if ($(window).width() > 1023) {
        $("#bulk-tab").html("Bulk <br /> Add");
    }
    else {
        $("#bulk-tab").html("Bulk Add");
        $("#bulk-panel").css("right", "initial");
    }
}

function showBulkTab(x) {
    if ($(window).width() > 1023) {
        if (x == 1) {
            $("#bulk-panel").css("right", "0px");
            $("#bulk-panel").show();
        }
        else {
            $("#bulk-panel").css("right", "-210px");
        }
    } else
    {
        if (x == 1) {
            $("#bulk-panel").css("bottom", "0px");
            $("#bulk-panel").show();
        }
        else {
            $("#bulk-panel").css("bottom", "-800px");
        }
    }
}

$(document).ready(function () {

    $(window).resize(function () {
        mediasize();
    });

    if ($(window).width() > 1023) {
        $("#bulk-panel").show();
    }
    
    if (readCookie("tabpos") == null) {
        document.cookie = "tabpos=0; path=/"
    }

    if (readCookie("tabpos") == 0) {
        showBulkTab(0);
    }
    else {
        showBulkTab(1);
    }

    $("#bulk-tab").on("click", function () {

        if (readCookie("tabpos") == 1) {
            showBulkTab(0);
            document.cookie = "tabpos=0; path=/"
        }
        else {
            showBulkTab(1);
            document.cookie = "tabpos=1; path=/"
        }
    });
    $("#lnkbulkadd").on("click", function () {
        if ($(window).width() > 1023) {
            showBulkTab(1);
        }
        else {
            mediasize();
            $("#bulk-panel").show();
            $("#bulk-panel").animate({ bottom: "0px" }, 500)
        }
        
    });
        $("#bulkaddtocart").on("click", function () {
            if (getalladds().length > 0) {
                $.ajax({
                    url: "bulkAddToCartdb.asp?addtocart=1", success: function (result) {
                        var $rtnMessageArr = result.split('|');
                        
                        if ($rtnMessageArr[0] == 0) {
                            $.ajax({
                                url: "instprd.asp?" + $rtnMessageArr[1], success: function (result) {
                                    $('#cat-bulk-add .form-control').each(function (index, data) {
                                        $('#sku' + (index)).val("");
                                        $('#sku' + (index) + 'qty').val(1);
                                    });
                                    $.ajax({
                                        url: "bulkAddToCartdb.asp?reset=1", success: function (result) {
                                        }
                                    });  
                                    window.location.replace("viewcart.asp")
                                }
                            });
                        }
                        else
                        {
                            alert("Error: " & result);
                        }
                    }
                });
            }
        });
        $("#bulkreset").on("click", function () {
            $('#cat-bulk-add .form-control').each(function (index, data) {
                $('#sku' + (index)).val(""); 
                $('#sku' + (index)+'qty').val(1); 
            });
            $.ajax({
                url: "bulkAddToCartdb.asp?reset=1", success: function (result) {

                }
            });  
        });

        $("#cat-bulk-add").on("change", '.form-control.sku, .form-control.qty', function (){
            var cursku = $(this).prop("id");
            cursku = cursku.substring(0, 4);
            $cursku = cursku;
            $curskupos = $cursku;
            $curskuval = $('#'+cursku).val();
            $curskuqty = $('#'+cursku+'qty').val();
            $('.alert.' + $curskupos + '.alert.alert-danger').remove();
            if ($curskuval.length > 0) {
                $.ajax({
                    url: "bulkAddToCartdb.asp?curskupos=" + $curskupos + "&curskuval=" + $curskuval + "&curskuqty=" + $curskuqty + "&all=" + getalladds(), success: function (result) {
                        var $rtnMessageArr = result.split('|');
                       
                        if ($rtnMessageArr[0] == 1) {
                            $('#' + $curskupos).closest('.bulksku').append("<div class='alert " + $curskupos + " alert alert-danger'><a name='error'>" + $rtnMessageArr[1] + "</a></div>");
                            $('#' + $curskupos).val("");
                        }
                        else
                        {
                            $('.alert.' + $curskupos + '.alert.alert-danger').remove();
                        }

                    }
                });
            }
            else {
            };
        })

        $("#cat-bulk-add").on("change", '.form-control.qty', function () { //validate inputs
            var $curname = $(this).attr('name');
            var $clean = this.value.replace(/[^0-9.]/g, '');
            if ($clean != "") {
                $clean = parseInt($clean);
                $clean = $clean.toFixed(0);
            }
            if ($clean == 0) {
                $clean = 1;
            }
            this.value = $clean;
        });
    });
    function getalladds() {
        var $alladded = "";
        $('#cat-bulk-add .form-control').each(function (index, data) {
            if ($('#sku' + (index)).val()) {
                $alladded = $alladded + $('#sku' + (index )).val() + "||" + $('#sku' + (index) + 'qty').val() + "||"
            }
        });
        return $alladded;
    }


    function readCookie(name) {
        var nameEQ = name + "=";
        var ca = document.cookie.split(';');
        for (var i = 0; i < ca.length; i++) {
            var c = ca[i];
            while (c.charAt(0) == ' ') c = c.substring(1, c.length);
            if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length, c.length);
        }
        return null;
    }


