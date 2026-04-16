// START: ORDER DETAIL CONTROLLER
app.controller('orderDetailCtrl', ['$scope', '$http', '$timeout', '$log', '$compile', 'httpRequest', 'pcService', function($scope, $http, $timeout, $log, $compile, httpRequest, pcService) {     

    $scope.shoppingcart = [];
    
    $scope.blockUI = function (div, msg) {
        $pc.blockUI.defaults.css = {
            top: '5%',
            left: '20%',
            right: '20%'
        };
        $pc(div).block({
            centerY: 0,
            centerX: false,
            message: '<div id="pcMain">' + msg + "</div>",
            overlayCSS: {
                backgroundColor: '#FFFFFF',
                cursor: 'wait',
                padding: '4px'
            }
        });
    };

    $scope.unblockUI = function (div) {
        $pc(div).unblock();
    };

    $scope.$on('handleOrderDetails', function(event, data){
        $scope.shoppingcart = data;
    });
    $scope.refresh = function () {
        pcService.getOrderDetails('', false);
    };
    function init() {
        if ($scope.shoppingcart == '') {
            pcService.getOrderDetails('', false); 
        }
    };
    init();

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

}]);
// END: ORDER DETAIL CONTROLLER