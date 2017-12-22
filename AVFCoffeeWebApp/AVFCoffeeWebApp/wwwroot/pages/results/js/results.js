

globalDataPromise.done(function (data) {
    updateResultChartDataObject();
    createResultChart();
});

function updateResultChartDataObject() {
    var chartDataObject = [];
    var variableData = { name: 'Variable', data: [] };
    var fixedData = { name: 'Fixed', data: [] };
    var additionalData = { name: 'Additional', data: [] };

    //producer - from UserDataObject
    variableData.data.push(Math.round(UserData.output.ProducerOutputEnglish.variableCostUSPound * 100) / 100);
    fixedData.data.push(Math.round(UserData.output.ProducerOutputEnglish.fixedCostUSPound * 100) / 100);
    additionalData.data.push(Math.round(UserData.output.ProducerOutputEnglish.totalCostAndDeprUSPound * 100) / 100);

    //coop
    variableData.data.push(UserData.output.Coop.variableCostUSPound);
    fixedData.data.push(UserData.output.Coop.fixedCostUSPound);
    additionalData.data.push(UserData.output.Coop.totalCostAndDeprUSPound);

    chartDataObject.push(variableData);
    chartDataObject.push(fixedData);
    chartDataObject.push(additionalData);

    UserData.resultChartDataObject = chartDataObject;
}
