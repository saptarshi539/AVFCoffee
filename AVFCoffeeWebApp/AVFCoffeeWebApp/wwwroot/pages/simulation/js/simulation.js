
globalDataPromise.done(function (data) {
    updateSimulationChartDataObject();
    createSimulationChart();
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
        var convFert = $("#conventionalFertilizers2").val();
        var orgFert = $("#organicFertilizers2").val();

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
            "&transportCostSoles=" + transportCostSoles + "&costPriceSolesPerQuintal=" + costPriceSolesPerQuintal + "&expSolesChem=" + convFert + "&expSolesOrg=" + orgFert,
            contentType: "application/json; charset=utf-8",
            success: function (result, status) {
                //set simulation to new output
                UserData.simulationOutput = result.output

                //create new chart object
                updateSimulationChartDataObject();

                //update the chart
               // var newData = UserData.chartDataObject;
                var chart = $('#chartdiv2').highcharts();
                //chart.redraw();
              
                //note - we should not have to destroy and recreate the chart, highcharts has a
                //redraw function that is supposed to reanimate the chart with new data,  but it is 
                //not working here.
                chart.destroy();
                createSimulationChart();
            },
            error: function (res, status) {
                if (status === "error") {
                    console.log("error");
                }
            }
        });
    });
}

function updateSimulationChartDataObject() {
    var chartDataObject = [];
    var variableData = { name: 'Variable', data: [], index: 2 };
    var fixedData = { name: 'Fixed', data: [], index: 1 };
    var additionalData = { name: 'Additional', data: [], index: 0};

    //producer - from UserDataObject
    additionalData.data.push(Math.round(UserData.output.ProducerOutputEnglish.totalCostAndDeprUSPound * 100) / 100);
    fixedData.data.push(Math.round(UserData.output.ProducerOutputEnglish.fixedCostUSPound * 100) / 100);
    variableData.data.push(Math.round(UserData.output.ProducerOutputEnglish.variableCostUSPound * 100) / 100);

    //simulation
    additionalData.data.push(Math.round(UserData.simulationOutput.ProducerOutputEnglish.totalCostAndDeprUSPound * 100) / 100);
    fixedData.data.push(Math.round(UserData.simulationOutput.ProducerOutputEnglish.fixedCostUSPound * 100) / 100);
    variableData.data.push(Math.round(UserData.simulationOutput.ProducerOutputEnglish.variableCostUSPound * 100) / 100);

    chartDataObject.push(variableData);
    chartDataObject.push(fixedData);
    chartDataObject.push(additionalData);

    language.EN.chart.simulationData = chartDataObject
    language.EN.chart.plotlinePriceRecieved = Math.round(UserData.output.ProducerOutputEnglish.breakEvenCostUSPound * 100) / 100
    language.EN.chart.plotlinePriceRecievedText = Math.round(UserData.output.ProducerOutputEnglish.breakEvenCostUSPound * 100) / 100 + "<br /> Price <br/> received"
    language.EN.chart.plotlineWorldPrice = Math.round(UserData.output.ProducerOutputEnglish.futuresPrice * 1000) / 1000
    language.EN.chart.plotlineWorldPriceText = Math.round(UserData.output.ProducerOutputEnglish.futuresPrice * 1000) / 1000 + "<br />  World  <br/> price"

    var chartDataObjectES = [];
    var variableDataES = { name: 'Variable', data: [], index: 2 };
    var fixedDataES = { name: 'Fixed', data: [], index: 1 };
    var additionalDataES = { name: 'Additional', data: [], index: 0 };
    var conversion = language.ES.chart.chartUnitsConversion

    //producer - from UserDataObject
    additionalDataES.data.push(Math.round((UserData.output.ProducerOutputEnglish.totalCostAndDeprUSPound * conversion)* 100) / 100);
    fixedDataES.data.push(Math.round((UserData.output.ProducerOutputEnglish.fixedCostUSPound * conversion) * 100) / 100);
    variableDataES.data.push(Math.round((UserData.output.ProducerOutputEnglish.variableCostUSPound * conversion) * 100) / 100);

    //simulation
    additionalDataES.data.push(Math.round((UserData.simulationOutput.ProducerOutputEnglish.totalCostAndDeprUSPound * conversion) *100) / 100);
    fixedDataES.data.push(Math.round((UserData.simulationOutput.ProducerOutputEnglish.fixedCostUSPound * conversion)* 100) / 100);
    variableDataES.data.push(Math.round((UserData.simulationOutput.ProducerOutputEnglish.variableCostUSPound * conversion)* 100) / 100);
  
    chartDataObjectES.push(variableDataES);
    chartDataObjectES.push(fixedDataES);
    chartDataObjectES.push(additionalDataES);

    language.ES.chart.simulationData = chartDataObjectES
    language.ES.chart.plotlinePriceRecieved = Math.round((UserData.output.ProducerOutputEnglish.breakEvenCostUSPound * conversion) * 100) / 100 
    language.ES.chart.plotlinePriceRecievedText = Math.round((UserData.output.ProducerOutputEnglish.breakEvenCostUSPound * conversion) * 100) / 100 + "<br /> precio <br/> recibido";
    language.ES.chart.plotlineWorldPrice = Math.round((UserData.output.ProducerOutputEnglish.futuresPrice * conversion) * 1000) / 1000 
    language.ES.chart.plotlineWorldPriceText = Math.round((UserData.output.ProducerOutputEnglish.futuresPrice * conversion) * 1000) / 1000 + "<br />  precio  <br/> mundial";

    //UserData.simulationChartDataObject = chartDataObject;
}