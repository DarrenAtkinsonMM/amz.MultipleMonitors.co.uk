// Product Apps
var fsApp = angular.module('productcart.apps', []);

fsApp.config(function ($compileProvider, $locationProvider) {
    $locationProvider.html5Mode({ enabled: true, requireBase: false });
});

fsApp.directive('htmldiv', ['$compile', '$parse', function($compile, $parse) {
return {
  restrict: 'E',
  link: function(scope, element, attr) {
    scope.$watch(attr.content, function() {      
      if (scope.roundtrip == 1) {
          element.html($parse(attr.content)(scope));
          $compile(element.contents())(scope);
      }      
    }, true);
  }
}
}]);

fsApp.controller('appsCtrl', ['$scope', '$location', '$http', '$timeout', '$log', '$compile', 'httpRequest', 'pcService', '$sce', '$window', function($scope, $location, $http, $timeout, $log, $compile, httpRequest, pcService, $sce, $window) {

        // Defaults
        $scope.myhtml = '';
        $scope.categories= [];
        $scope.button = 'Pay Now';
        $scope.button_install = 'Install';
        $scope.isDisabled = false;
        $scope.currentView = '/';
        $scope.error = '';
        
        $scope.url = decodeURI($location.url());



        ////////////////////////////////////
        // API HANDLERS
        ////////////////////////////////////
        
        // getURL
        $scope.$on('handleGetURL', function(event, data) {
            $timeout(completeGetRequest, 3000);        
        });

        // getMarketView
        $scope.$on('handleMarketView', function(event, data) {
            $scope.roundtrip = 1;
            $scope.myhtml = data;         
        });
        
        // getRequestPayment
        $scope.$on('handleRequestPayment', function(event, data) {
            if (data=='success') {
                // complete the order
                // loadData($scope.currentView);
                $timeout(completePayment, 3000); 
            } else {
                // display error to merchant
                var errorMsg = '<div class="alert alert-danger">' + data + '</div>';
                console.log(data);
                $scope.error = $sce.trustAsHtml(errorMsg);
            }
        });
        
        // getInstallation
        $scope.$on('handleInstallation', function(event, data) {
            if (data=='success') {
                // Goto App Settings
                $timeout(completeInstall, 3000); 
            } else {
                // display error to merchant
                var errorMsg = '<div class="alert alert-warning">' + data + '</div>';
                console.log(data);
                $scope.error = $sce.trustAsHtml(errorMsg);
                $scope.isDisabled = false;
            }
        });
        
        // getUninstallation
        $scope.$on('handleUninstallation', function(event, data) {
            if (data=='success') {
                // Goto App Settings
                $timeout(completeUninstall, 3000); 
            } else {
                // display error to merchant
                var errorMsg = '<div class="alert alert-warning">' + data + '</div>';
                console.log(data);
                $scope.error = $sce.trustAsHtml(errorMsg);
                $scope.isDisabled = false;
            }
        });



        ////////////////////////////////////
        // API REQUESTS
        ////////////////////////////////////

        // Get Request
        function startGetRequest(f) {           
            pcService.getRequest(f, '', false);
        };

        // Load Data
        function loadData(val) {           
            pcService.getMarketView(val, false);
        };
        
        // Start Payment
        function startPayment(val, code) {
            console.log('...starting payment'); 
            $scope.currentView = val;                      
            pcService.getRequestPayment('subscribe', code, false);            
        };
        
        // Start Install        
        function startInstall(val, code) {
            console.log('...starting installation'); 
            $scope.currentView = val;                      
            pcService.getInstallation('install', code, false);            
        };
        
        // Start Uninstall        
        function startUninstall(val, code) {
            console.log('...starting Uninstallation'); 
            $scope.currentView = val;                      
            pcService.getUninstallation('uninstall', code, false);            
        };


        ////////////////////////////////////
        // EVENTS
        ////////////////////////////////////

        // Start: Load Market View
        $scope.ViewDetail = function (val) {            
            resetPage();
            loadData(val);
        };
        // End: Load Market View


        // Start: Pay Now Button Click
        $scope.PayNow = function (val, code) {
            console.log('...payment button clicked'); 
            $scope.isDisabled = true; 
            $scope.button = "Please Wait..";                         
            startPayment(val, code);
        };
        // End: Pay Now Button Click
        
        // Start: Install Button Click
        $scope.Install = function (val, code) {
            $scope.error = '';
            console.log('...install button clicked'); 
            $scope.isDisabled = true;                         
            startInstall(val, code);
        };
        // End: Install Button Click
        
        // Start: Uninstall Button Click
        $scope.Uninstall = function (val, code) {
            $scope.error = '';
            console.log('...uninstall button clicked'); 
            $scope.isDisabled = true;                         
            startUninstall(val, code);
        };
        // End: Uninstall Button Click

        // Start: Execute Local URL
        $scope.GetRequest = function (file) {
            $scope.error = '';
            console.log('button clicked');
            $scope.processing = true;                         
            startGetRequest(file, '', false);
        };
        // End: Execute Local URL


        ////////////////////////////////////
        // METHODS
        ////////////////////////////////////

        // Get App Status
        $scope.isInstalled = function (val) {
            // console.log('...checking if installed: ' + val); 
            if (eval(val)) {
                if (eval(val) ==  true) {
                   return true; 
                } else {
                   return false; 
                } 
            } else {
                return false;
            }           
        };
        
        // Complete Payment
        function completePayment() { 
            resetPage();  
            $window.location.href = 'pcws_MyApps.asp';          
        };
 
        // Complete Get Request
        function completeGetRequest() { 
            $scope.processing = false;
        };
        
        // Complete Install
        function completeInstall() { 
            resetPage(); 
            $window.location.href = 'pcws_MyApps.asp';
        };
        
        // Complete Uninstall
        function completeUninstall() { 
            resetPage(); 
            $window.location.href = 'pcws_MyApps.asp';
        };
        
        // Reset Default Variables
        function resetPage() {           
            $scope.button = 'Pay Now';
            $scope.isDisabled = false;
            $scope.error = '';
            $('#myModal').modal('hide');
        };

        // Refresh Page View
        function refreshView() { 
            $scope.url = $location.absUrl().split('?')[1]  
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

        refreshView();

}]);