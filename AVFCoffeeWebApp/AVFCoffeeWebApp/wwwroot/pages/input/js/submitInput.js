
$("#submitInput").click(function () {
    //get user inputs
    var earlyHectares = $("#earlyHectares").val();
    var peakHectares = $("#peakHectares").val();
    var oldHectares =  $("#oldHectares").val();
    var conventional = $("#conventional").prop("checked");
    var organic = $("#organic").prop("checked");
    var transition = $("#transition").prop("checked");
    var workerSalarySoles = $("#workerSalarySoles").val();
    var productionQuintales = $("#productionQuintales").val();
    var transportCostSoles = $("#transportCostSoles").val();
    var costPriceSolesPerQuintal = $("#costPriceSolesPerQuintal").val();
    var convFert = $("#conventionalFertilizers").val();
    var orgFert = $("#organicFertilizers").val();
   
    //make user inputs object 
    var userInputs = {
        "earlyHectares": earlyHectares,
        "peakHectares": peakHectares,
        "oldHectares": oldHectares,
        "conventional": conventional, 
        "organic": organic,
        "transition": transition,
        "workerSalarySoles": workerSalarySoles,
        "productionQuintales": productionQuintales,
        "transportCostSoles": transportCostSoles,
        "costPriceSolesPerQuintal": costPriceSolesPerQuintal,
        "expSolesChem": convFert,
        "expSolesOrg": orgFert
    }


    $.ajax({
        type: "GET",
        url: apiURL + "CellSum/calculate",
        data: "earlyHectares=" + earlyHectares + "&peakHectares=" + peakHectares + "&oldHectares=" + oldHectares +
        "&conventional=" + conventional + "&organic=" + organic + "&transition=" + transition +
        "&workerSalarySoles=" + workerSalarySoles + "&productionQuintales=" + productionQuintales +
        "&transportCostSoles=" + transportCostSoles + "&costPriceSolesPerQuintal=" + costPriceSolesPerQuintal + "&expSolesChem=" + convFert +"&expSolesOrg=" + orgFert, 
        contentType: "application/json; charset=utf-8",
        success: function (result, status) {
            var userInputPromise = saveUserInput(userInputs);
            userInputPromise.then(saveUserOutput(result)); 
        },
        error: function (res, status) {
            if (status === "error") {
                console.log("error");
            }
        }
    });
});


function saveUserInput(userData) {
    var request = JSON.stringify(userData)
    var promise = $.ajax({
        type: "POST",
        url: apiURL + "CellSum/saveinput",
        data: request,
        contentType: "application/json; charset=utf-8",
        success: function (result, status) {

        },
        error: function (res, status) {
        }
    });
    return promise;
}

function saveUserOutput(outputData) {
    var request = JSON.stringify(outputData)
    $.ajax({
        type: "POST",
        url: apiURL + "CellSum/saveoutput",
        contentType: "application/json; charset=utf-8",
        data: request,
        success: function (result, status) {
            // save input and output .then
            window.location.href = '/Results';
        },
        error: function (res, status) {
        }
    });
}