
globalDataPromise.done(function (data) {
    updateChartDataObject();
    createChart();
    initSimClickEvent();
});

function initSimClickEvent() {
    $("#submitInput2").click(function () {
        //get user inputs
        var earlyHectares = $("#earlyHectares2").val();
        var peakHectares = $("#peakHectares2").val();
        var oldHectares = $("#oldHectares2").val();
        var conventional = $("#conventional2").prop("checked");
        var organic = $("#organic2").prop("checked");
        var transition = $("#transition2").prop("checked");
        var workerSalarySoles = $("#workerSalarySoles2").val();
        var productionQuintales = $("#productionQuintales2").val();
        var transportCostSoles = $("#transportCostSoles2").val();
        var costPriceSolesPerQuintal = $("#costPriceSolesPerQuintal2").val();

        //make user inputs object 
        var simulatorInputs = {
            "earlyHectares": earlyHectares,
            "peakHectares": peakHectares,
            "oldHectares": oldHectares,
            "conventional": conventional,
            "organic": organic,
            "transition": transition,
            "workerSalarySoles": workerSalarySoles,
            "productionQuintales": productionQuintales,
            "transportCostSoles": transportCostSoles,
            "costPriceSolesPerQuintal": costPriceSolesPerQuintal
        }

        $.ajax({
            type: "GET",
            url: apiURL + "CellSum/calculate",
            data: "earlyHectares=" + earlyHectares + "&peakHectares=" + peakHectares + "&oldHectares=" + oldHectares +
            "&conventional=" + conventional + "&organic=" + organic + "&transition=" + transition +
            "&workerSalarySoles=" + workerSalarySoles + "&productionQuintales=" + productionQuintales +
            "&transportCostSoles=" + transportCostSoles + "&costPriceSolesPerQuintal=" + costPriceSolesPerQuintal,
            contentType: "application/json; charset=utf-8",
            success: function (result, status) {
                //set simulation to new output
                UserData.simulationOutput = result.output

                //create new chart object
                updateChartDataObject();

                //update the chart
               // var newData = UserData.chartDataObject;
                var chart = $('#chartdiv2').highcharts();
                //chart.redraw();
              
                //note - we should not have to destroy and recreate the chart, highcharts has a
                //redraw function that is supposed to reanimate the chart with new data,  but it is 
                //not working here.
                chart.destroy();
                createChart();
            },
            error: function (res, status) {
                if (status === "error") {
                    console.log("error");
                }
            }
        });
    });
}

function updateChartDataObject() {
    var chartDataObject = [];
    var variableData = { name: 'Variable', data: [] };
    var fixedData = { name: 'Fixed', data: [] };
    var additionalData = { name: 'Additional', data: [] };

    //producer - from UserDataObject
    variableData.data.push(Math.round(UserData.output.ProducerOutputEnglish.variableCostUSPound * 100) / 100);
    fixedData.data.push(Math.round(UserData.output.ProducerOutputEnglish.fixedCostUSPound * 100) / 100);
    additionalData.data.push(Math.round(UserData.output.ProducerOutputEnglish.totalCostAndDeprUSPound * 100) / 100);
    
    //simulation
    variableData.data.push(Math.round(UserData.simulationOutput.ProducerOutputEnglish.variableCostUSPound * 100) / 100);
    fixedData.data.push(Math.round(UserData.simulationOutput.ProducerOutputEnglish.fixedCostUSPound * 100) / 100);
    additionalData.data.push(Math.round(UserData.simulationOutput.ProducerOutputEnglish.totalCostAndDeprUSPound * 100) / 100);

    //coop
    variableData.data.push(UserData.output.Coop.variableCostUSPound);
    fixedData.data.push(UserData.output.Coop.fixedCostUSPound);
    additionalData.data.push(UserData.output.Coop.totalCostAndDeprUSPound);

    chartDataObject.push(variableData);
    chartDataObject.push(fixedData);
    chartDataObject.push(additionalData);

    UserData.chartDataObject = chartDataObject;
    console.log(UserData.chartDataObject)
}