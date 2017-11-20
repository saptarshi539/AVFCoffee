
$("#submitInput").click(function () {
    //get user inputs
    var earlyHectares = $("#earlyHectares").val();
    var peakHectares = $("#peakHectares").val();
    var oldHectares = $("#oldHectares").val();
    var conventional = $("#conventional").prop("checked");
    var organic = $("#organic").prop("checked");
    var transition = $("#transition").prop("checked");
    var workerSalarySoles = $("#workerSalarySoles").val();
    var productionQuintales = $("#productionQuintales").val();
    var transportCostSoles = $("#transportCostSoles").val();
    var costPriceSolesPerQuintal = $("#costPriceSolesPerQuintal").val();

    $.ajax({
        type: "GET",
        url: apiURL + "CellSum/calculate",
        data: "earlyHectares=" + earlyHectares + "&peakHectares=" + peakHectares + "&oldHectares=" + oldHectares +
        "&conventional=" + conventional + "&organic=" + organic + "&transition=" + transition +
        "&workerSalarySoles=" + workerSalarySoles + "&productionQuintales=" + productionQuintales +
        "&transportCostSoles=" + transportCostSoles + "&costPriceSolesPerQuintal=" + costPriceSolesPerQuintal,
        contentType: "application/json; charset=utf-8",
        success: function (result, status) {
            chartDataObject = [];
            //result = { "producer": [.41, .06, .84], "cooperative": [.44, .04, .89] }
            //different ways to work with data, chart wants format below. Will depend on what comes back from server 
            /*chartDataObject = {
                data: [{
                    name: 'Variable',
                    data: [.84, .89]
                }, {
                    name: 'Fixed',
                    data: [.06, .04],
                }, {
                    name: 'Additional',
                    data: [.41, .44],
                }]
            } */
            var variableData = { name: 'Variable', data: [] };
            var fixedData = { name: 'Fixed', data: [] };
            var additionalData = { name: 'Additional', data: [] };

            variableData.data.push(Math.round(result.output.ProducerOutputEnglish.variableCostUSPound * 100) /100);
            fixedData.data.push(Math.round(result.output.ProducerOutputEnglish.fixedCostUSPound * 100) /100);
            additionalData.data.push(Math.round(result.output.ProducerOutputEnglish.totalCostAndDeprUSPound * 100) /100);
            variableData.data.push(result.output.Coop.variableCostUSPound);
            fixedData.data.push(result.output.Coop.fixedCostUSPound);
            additionalData.data.push(result.output.Coop.totalCostAndDeprUSPound);

            chartDataObject.push(variableData);
            chartDataObject.push(fixedData);
            chartDataObject.push(additionalData);

            localStorage.setItem("chartDataObject", JSON.stringify(chartDataObject));

            //go to chart page
            window.location.href = '/Results';
        },
        error: function (res, status) {
            if (status === "error") {
                console.log("error");
            }
        }
    });
});

