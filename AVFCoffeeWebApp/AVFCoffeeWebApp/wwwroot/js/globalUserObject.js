var UserData = {
    input: {},
    output: {},
    simulationInput: {},
    simulationOutput: {},
    chartDataObject: {},
    userData: {}
};

 
 var globalDataPromise = $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        url: apiURL + "CellSum/getinput",
        success: function (content, status) {
            if (status != 'nocontent') {
                UserData.input = content.loginfo.Inputs;
                UserData.output = content.loginfo.Outputs;
                UserData.simulationOutput = content.loginfo.Outputs;
                UserData.simulationInput = content.loginfo.Inputs;
                //find out what page is requesting the info
                var path = window.location.pathname;
                var page = path.split("/").pop();

                if (page.toLowerCase() == "simulation") {
                    //load users data into simulation page
                    $("#earlyHectares2").val(UserData.simulationInput.earlyHectares);
                    $("#peakHectares2").val(UserData.simulationInput.peakHectares);
                    $("#oldHectares2").val(UserData.simulationInput.oldHectares);
                    $("#conventional2").prop("checked", UserData.simulationInput.conventional);
                    $("#organic2").prop("checked", UserData.simulationInput.organic);
                    $("#transition2").prop("checked", UserData.simulationInput.transition);
                    $("#workerSalarySoles2").val(UserData.simulationInput.workerSalarySoles);
                    $("#productionQuintales2").val(UserData.simulationInput.productionQuintales);
                    $("#transportCostSoles2").val(UserData.simulationInput.transportCostSoles);
                    $("#costPriceSolesPerQuintal2").val(UserData.simulationInput.costPriceSolesPerQuintal);
                }
                else if (page.toLowerCase() == "input") {
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

  
