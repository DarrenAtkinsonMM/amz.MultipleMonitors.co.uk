// Product Apps
fsApp.controller('apiCtrl', ['$scope', '$location', '$http', '$timeout', '$log', '$compile', 'httpRequest', 'pcService', '$sce', '$window', function($scope, $location, $http, $timeout, $log, $compile, httpRequest, pcService, $sce, $window) {

        ////////////////////////////////////
        // API HANDLERS
        ////////////////////////////////////
        
        // getURL
        $scope.$on('handleGetURL', function(event, data) {
            // $timeout(completeGetRequest, 3000); 
            $scope.data = data;     
        });


        ////////////////////////////////////
        // API REQUESTS
        ////////////////////////////////////

        // Get Request
        function startGetRequest(f, p, b) {           
            pcService.getRequest(f, p, b);             
        };     


        ////////////////////////////////////
        // EVENTS
        ////////////////////////////////////

        // Start: Execute Local URL
        $scope.GetRequest = function (file, payload) {
            $scope.error = '';
            // $scope.processing = true;                         
            startGetRequest(file, payload, false);
        };
        // End: Execute Local URL


        ////////////////////////////////////
        // METHODS
        ////////////////////////////////////

        // Complete Get Request
        //function completeGetRequest() { 
        //    $scope.processing = false;
        //};
        
        $scope.load = function (path) {
            $scope.GetRequest(path, null);            
        };
        $scope.delete = function (path, v, id) {
            var payload = v + '=' + id;
            $scope.GetRequest(path, payload);            
        };
        $scope.create = function (path, f) {
            var payload = $('#' + f).serialize();
            $scope.GetRequest(path, payload);  
            $('#' + f).trigger("reset");          
        };
        
        
        ////////////////////////////////////
        // UTILITIES
        ////////////////////////////////////
        
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