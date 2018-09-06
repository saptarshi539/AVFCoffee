var farmerPhone = localStorage.getItem("farmerPhone");
var UserData = {
    input: {},
    output: {},
    resultChartDataObject: {},
    simulationInput: {},
    simulationOutput: {},
    simulationChartDataObject: {},
    user: {}
};

//$(document).on('ready', function () {
    
    console.log(farmerPhone);
    var globalDataPromise = $.ajax({
        type: "GET",
        url: apiURL + "CellSum/getinput?",
        data: "phoneNumber=" + farmerPhone,
        contentType: "application/json; charset=utf-8",
        success: function (content, status) {
            if (status != 'nocontent') {
                console.log(content)
                UserData.input = content.loginfo.Inputs;
                UserData.output = content.loginfo.Outputs;
                UserData.simulationOutput = content.loginfo.Outputs;
                UserData.simulationInput = content.loginfo.Inputs;
                UserData.user.language = content.loginfo.User.language;
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
                    $("#conventionalFertilizers2").val(UserData.input.expSolesChem);
                    $("#organicFertilizers2").val(UserData.input.expSolesOrg);
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
                    $("#conventionalFertilizers").val(UserData.input.expSolesChem);
                    $("#organicFertilizers").val(UserData.input.expSolesOrg);
                }

                translate()
            }
        },
        error: function () {
            console.log('not successful');
        }

    });
//});


