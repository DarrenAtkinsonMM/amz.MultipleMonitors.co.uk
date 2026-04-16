// ProductCart App
var app = angular.module('productcart', ['productcart.apps']);

app.config(function ($compileProvider, $locationProvider) {
    //$compileProvider.urlSanitizationWhitelist(/^\s*((https?|ftp|mailto|tel):)|.*?/);
    //$compileProvider.aHrefSanitizationWhitelist(/^\s*((https?|ftp|mailto|tel):)|.*?/);
});
app.directive('a', function () {
    return {
        restrict: 'E',
        link: function (scope, element, attrs) {
            var relExternal = attrs.rel && attrs.rel.split(/\s+/).indexOf('external') >= 0;
            var relInternal = attrs.rel && attrs.rel.split(/\s+/).indexOf('internal') >= 0;
            if (!('target' in attrs) && !(relExternal || 'external' in attrs) && !(relInternal || 'internal' in attrs)) {
                attrs.$set('target', '_self');
            }
        }
    };
});


// Bootstrap
angular.element(document).ready(function () {
    // angular.bootstrap(document, ['productcart']);
});


// START: SERVICES
app.factory('httpRequest', function ($q, $http) {
    
    var mycache={};
    var key = 'cache';
    
    return {
        loadAsync: function (endpoint, payload, cache) {

            var defer = $q.defer();
            
            if (mycache[key] && cache==true) {
                defer.resolve(mycache[key]);
            } else {
                var httpRequest = $http({
                    method: 'POST',
                    url: endpoint,
                    data: payload,
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
                }).success(function (data, status) {
                    mycache[key] = data;
                    defer.resolve(mycache[key]);
                }).error(function (data, status) {

                    if (status=='404') {
                        mycache[key] = 'We couldn\'t find that app. Did you upload it?';
                    }
                    
                    if (status=='500') {
                        mycache[key] = data;
                    }
                    
                    defer.resolve(mycache[key]);
                });
            }
            
            return defer.promise;

        }
    }
});

app.service('pcService', function ($rootScope, httpRequest, currencyFilter) {

    return {   
    
        getRequest: function (f, p, c) {
            httpRequest.loadAsync(pcRootUrl + f, p, c).then(function (data) {
                $rootScope.$broadcast('handleGetURL', data);
            });  
        },
        getMarketView: function (s, c) {
            httpRequest.loadAsync('service/api/market.asp?params=' + s, '', c).then(function (data) {
                $rootScope.$broadcast('handleMarketView', data);
            });  
        },
        getInstallation: function (e, s, c) {
            var q = '?event=' + e + '&code=' + s + '&name=' + 'na';  
            httpRequest.loadAsync(pcRootUrl + '/includes/apps/' + s + '/install.asp' + q, '', c).then(function (data) {
                $rootScope.$broadcast('handleInstallation', data);
            });  
        },
        getUninstallation: function (e, s, c) {
            var q = '?event=' + e + '&code=' + s + '&name=' + 'na';  
            httpRequest.loadAsync(pcRootUrl + '/includes/apps/' + s + '/uninstall.asp' + q, '', c).then(function (data) {
                $rootScope.$broadcast('handleUninstallation', data);
            });  
        },
        getRequestPayment: function (e, s, c) {
            var q = '?event=' + e + '&code=' + s + '&name=' + 'na';  
            httpRequest.loadAsync('service/api/subscribe.asp' + q, '', c).then(function (data) {
                $rootScope.$broadcast('handleRequestPayment', data);
            });  
        }   
        
    }
    
});
// END: SERVICES


// START: PRIMARY CONTROLLER
app.controller('serviceCtrl', function($scope, $http, $timeout, $log, $compile, httpRequest, pcService) {     
    function init() {
        pcService.getShoppingCart('', false); 
    };
    //init();
});
// END: PRIMARY CONTROLLER