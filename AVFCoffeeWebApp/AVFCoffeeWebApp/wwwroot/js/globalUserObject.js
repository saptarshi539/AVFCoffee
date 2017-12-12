﻿var UserData = {
    input: {},
    output: {},
    userData: {}
};

 
 $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        url: apiURL + "CellSum/getinput",
        success: function (content, status) {
            if (status != 'nocontent') {
                UserData.input = content.loginfo.Inputs;
                UserData.output = content.loginfo.Outputs;
                //find out what page is requesting the info
                var path = window.location.pathname;
                var page = path.split("/").pop();
                console.log(page);

                if (page == "Simulation") {
                    //load users data into simulation page
                    $("#earlyHectares2").val(UserData.input.earlyHectares);
                    $("#peakHectares2").val(UserData.input.peakHectares);
                    $("#oldHectares2").val(UserData.input.oldHectares);
                    $("#conventional2").prop("checked", UserData.input.conventional);
                    $("#organic2").prop("checked", UserData.input.organic);
                    $("#transition2").prop("checked", UserData.input.transition);
                    $("#workerSalarySoles2").val(UserData.input.workerSalarySoles);
                    $("#productionQuintales2").val(UserData.input.productionQuintales);
                    $("#transportCostSoles2").val(UserData.input.transportCostSoles);
                    $("#costPriceSolesPerQuintal2").val(UserData.input.costPriceSolesPerQuintal);
                }
                else if (page == "input") {
                    //load users data into input page
                    $("#earlyHectares").val(UserData.input.earlyHectares);
                    $("#peakHectares").val(UserData.input.peakHectares);
                    $("#oldHectares").val(UserData.input.oldHectares);
                    $("#conventional").prop("checked", UserData.input.conventional);
                    $("#organic").prop("checked", UserData.input.organic);
                    $("#transition").prop("checked", UserData.input.transition);
                    $("#workerSalarySoles").val(UserData.input.workerSalarySoles);
                    $("#productionQuintales").val(UserData.input.productionQuintales);
                    $("#transportCostSoles").val(UserData.input.transportCostSoles);
                    $("#costPriceSolesPerQuintal").val(UserData.input.costPriceSolesPerQuintal);
                }
               
            }
        },
        error: function () {
            console.log('not successful');
        }

});

  
