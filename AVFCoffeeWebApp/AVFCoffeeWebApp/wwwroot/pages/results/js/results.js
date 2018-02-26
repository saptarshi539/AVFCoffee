

globalDataPromise.done(function (data) {
    updateResultChartDataObject();
    createResultChart();
});

function updateResultChartDataObject() {
    var chartDataObject = [];
    var variableData = { name: 'Variable', data: [], index: 2 };
    var fixedData = { name: 'Fixed', data: [], index: 1 };
    var additionalData = { name: 'Additional', data: [] , index: 0};

    //producer - from UserDataObject
    additionalData.data.push(Math.round(UserData.output.ProducerOutputEnglish.totalCostAndDeprUSPound * 100) / 100);
    fixedData.data.push(Math.round(UserData.output.ProducerOutputEnglish.fixedCostUSPound * 100) / 100);
    variableData.data.push(Math.round(UserData.output.ProducerOutputEnglish.variableCostUSPound * 100) / 100);
    

    //coop
    additionalData.data.push(UserData.output.Coop.totalCostAndDeprUSPound);
    fixedData.data.push(UserData.output.Coop.fixedCostUSPound);
    variableData.data.push(UserData.output.Coop.variableCostUSPound);
    

    chartDataObject.push(variableData);
    chartDataObject.push(fixedData);
    chartDataObject.push(additionalData);

    UserData.resultChartDataObject = chartDataObject;
}
