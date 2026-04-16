// START: SYNDICATION CONTROLLER
app.controller('syndicationCtrl', ['$scope', '$http', '$timeout', '$log', '$compile', 'httpRequest', 'pcService', function($scope, $http, $timeout, $log, $compile, httpRequest, pcService) {     

    $scope.syndicationlist = [];
    
    $scope.$on('handleSyndicationItems', function(event, data){
        $scope.syndicationlist = data;
    });      
    function init() {
        pcService.getSyndicationItems('', false);  
    };
    init();
    
    $scope.displayItems = function (val) {
        if ($scope.syndicationlist.totalItems > 0) {
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

    $scope.Evaluate = function (val) {
        if (val == 'true' || val == true) {
            return true;
        } else {
            return false;
        }
    };

}]);
// END: SYNDICATION CONTROLLER