// <reference path="messageread.js" />
var app = angular.module('edgelegal', ['ngMaterial', "ngRoute"], function () {


});


app.controller('edgelegalctrl', function ($scope, $mdDialog, $mdToast, $log, $location,) {
    
    $scope.ShowMainDiv = false;
    $scope.ShowRefreshBTN = false;
    var filecontent = "";


    Office.onReady(function (info) {

        if (info.host === Office.HostType.Excel) {

            ProgressLinearActive();
            let MatterNumberdialog;
            let logdialog;

            let userInfo = window.localStorage.getItem('userinfo');
            userInfo = JSON.parse(userInfo)
            if (userInfo) {

                if (userInfo) {
                    $scope.ShowMainDiv = true;
                    $scope.userName = userInfo.userName
                    ProgressLinearInActive();
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }


                } else {

                    $scope.ShowMainDiv = false;
                    ProgressLinearInActive();
                    openDialog();

                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                }

            } else {
                $scope.ShowMainDiv = false;
                ProgressLinearInActive();

                openDialog();

            }


            function openDialog() {
                Office.context.ui.displayDialogAsync(`https://localhost:44311/Templates/Login.html`, { height: 50, width: 30 },
                    function (asyncResult) {
                        logdialog = asyncResult.value;
                        logdialog.addEventHandler(Office.EventType.DialogMessageReceived, logprocess);
                    }
                );
            }

           

            function logprocess(arg) {
                logdialog.close();
                ProgressLinearActive();
                console.log(arg)
                let message = JSON.parse(arg.message);
                window.localStorage.setItem('userinfo', JSON.stringify(message));
                //let userdata = message;
                //message = message.login;
                if (message.login === true) {
                    console.log(message)
                    //$scope.userName = userdata.userName
                    loadToast("logged in successfully")
                    $scope.ShowMainDiv = true;
                    $scope.ShowRefreshBTN = false;
                    ProgressLinearInActive();

                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                    //$scope.Message = false;
                } else {
                    dialog.close();
                    $scope.ShowMainDiv = false;
                    $scope.ShowRefreshBTN = true;
                    ProgressLinearInActive();
                    loadToast("Refresh Addin To Login ")

                    //$scope.Message = true;
                }
            }
            $scope.getFilebase64 = function (ev) {

                Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 },
                    function (result) {
                        if (result.status === "succeeded") {
                            const myFile = result.value;
                            myFile.getSliceAsync(0, function (result) {

                                let data = result.value.data

                                let btoadata = btoa(String.fromCharCode.apply(null, new Uint8Array(data)))



                                if (btoadata) {
                                    myFile.closeAsync();
                                    filecontent = btoadata;
                                    Excel.run(function (context) {
                                        var workbook = context.workbook;

                                        workbook.load(["name"]);

                                        return context.sync()
                                            .then(function () {
                                                // Access the workbook name
                                                var workbookName = workbook.name;
                                                $scope.Filename = workbookName;
                                                Office.context.ui.displayDialogAsync(`https://localhost:44311/Templates/MatterNumber.html?workbookName=${workbookName}`, { height: 50, width: 30 },
                                                    function (asyncResult) {
                                                        MatterNumberdialog = asyncResult.value;
                                                        MatterNumberdialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                                                    }
                                                );

                                                //ProgressLinearInActive();
                                                console.log("Active workbook name: " + workbookName);
                                            })
                                            .catch(function (error) {
                                                console.log("Error: " + error);
                                            });
                                    }).catch(function (error) {
                                        console.log("Error: " + error);
                                    });
                                }



                            });

                        } else {
                            myFile.closeAsync()

                            // Handle the error here
                            //loadToast("Upload Error");
                            //ProgressLinearInActive();

                        }

                    }
                );
            }



           
            function processMessage(arg) {
                ProgressLinearActive();

                MatterNumberdialog.close();

                console.log(arg)
                let message = JSON.parse(arg.message);
                if (message.close == true) {

                    MatterNumberdialog.close();
                    ProgressLinearInActive();

                } else {
                    console.log(message)
                    message.filecontent = filecontent;
                    console.log(message);
                    var form = new FormData();
                    form.append("originalName",message.originalName);
                    form.append("matterId", message.matternumber);
                    form.append("", filecontent); 
                    form.append("UserName", "Aamir");

                    var settings = {
                        "url": "https://grazingdelights.com.au/LPDM/RT/WS/uploadMatterAttachment",
                        "method": "POST",
                        "timeout": 0,
                        "headers": {
                        }, 
                        "processData": false,
                        "mimeType": "multipart/form-data",
                        "contentType": false,
                        "data": form
                    };

                    $.ajax(settings).done(function (response) {
                        console.log(response);
                        loadToast("Uploaded Successfuly");
                        ProgressLinearInActive();
                    }).fail(function (error) {

                        console.log(error);
                        loadToast("upload error");
                        ProgressLinearInActive();
                    });
                }
            }


            $scope.Logout = function () {
                window.localStorage.clear('userinfo');
                window.location.reload();
            }

            $scope.Help = function () {
                window.open("https://support.microsoft.com/en-us")
            }


        } else {


            loadtost("word addin is wroking ")
            loadtost("no functioanlities included yet")

        }
      
    });
        function ProgressLinearActive() {
            $("#StartProgressLinear").show(function () {

                $("#ProgressBgDiv").show();
                $scope.ddeterminateValue = 15;
                $scope.showProgressLinear = false;
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            });
        };
        function ProgressLinearInActive() {
            $("#StartProgressLinear").hide(function () {
                setTimeout(function () {
                    $scope.ddeterminateValue = 0;
                    $scope.showProgressLinear = true;
                    $("#ProgressBgDiv").hide();
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                }, 500);
            });
        };
        function loadToast(alertMessage) {
            var el = document.querySelectorAll('#zoom');
            $mdToast.show(
                $mdToast.simple()
                    .textContent(alertMessage)
                    .position('bottom')
                    .hideDelay(4000))
                .then(function () {
                    $log.log('Toast dismissed.');
                }).catch(function () {
                    $log.log('Toast failed or was forced to close early by another toast.');
                });
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        };

        if (!$scope.$$phase) {
            $scope.$apply();
        }

  
      
   
})
