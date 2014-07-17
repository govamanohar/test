
var app = angular.module('myApp', ['ngGrid']);
app.controller('MyCtrl', function($scope) {
    $scope.form={};
    $scope.types=['Male','Female'];
    $scope.myData=[];
    $scope.form.sex=$scope.types[0];                 
    $scope.gridOptions = { data: 'myData' };
    $scope.adddItem=function(){
        $scope.myData.push({firstName:$scope.form.firstName,lastName:$scope.form.lastName,designation:$scope.form.designation,sex:$scope.form.sex});
        $scope.form={};
        $scope.form.sex=$scope.types[0]; 
    }
});
