

globalDataPromise.done(function (data) {
    updateResultChartDataObject();
    createResultChart();
});

function updateResultChartDataObject() {
    var chartDataObject = [];
    var variableData = { name: 'Variable', data: [], index: 2 };
    var fixedData = { name: 'Fixed', data: [], index: 1 };
    var additionalData = { name: 'Additional', data: [], index: 0 };

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

    language.EN.chart.data = chartDataObject
    language.EN.chart.plotlinePriceRecieved = Math.round(UserData.output.ProducerOutputEnglish.breakEvenCostUSPound * 100) / 100
    language.EN.chart.plotlinePriceRecievedText = Math.round(UserData.output.ProducerOutputEnglish.breakEvenCostUSPound * 100) / 100 + "<br /> Price <br/> received"
    language.EN.chart.plotlineWorldPrice = Math.round(UserData.output.ProducerOutputEnglish.futuresPrice * 1000) / 1000
    language.EN.chart.plotlineWorldPriceText = Math.round(UserData.output.ProducerOutputEnglish.futuresPrice * 1000) / 1000 + "<br />  World  <br/> price"

    var chartDataObjectES = [];
    var variableDataES = { name: 'Variable', data: [], index: 2 };
    var fixedDataES = { name: 'Fixed', data: [], index: 1 };
    var additionalDataES = { name: 'Additional', data: [], index: 0 };
    var conversion = language.ES.chart.chartUnitsConversion

    //producer - from UserDataObject - convert to 
    additionalDataES.data.push(Math.round((UserData.output.ProducerOutputEnglish.totalCostAndDeprUSPound * conversion) * 100) / 100);
    fixedDataES.data.push(Math.round((UserData.output.ProducerOutputEnglish.fixedCostUSPound * conversion) * 100) / 100);
    variableDataES.data.push(Math.round((UserData.output.ProducerOutputEnglish.variableCostUSPound * conversion) * 100) / 100);

    //coop
    additionalDataES.data.push(Math.round((UserData.output.Coop.totalCostAndDeprUSPound * conversion) * 100) / 100);
    fixedDataES.data.push(Math.round((UserData.output.Coop.fixedCostUSPound * conversion) * 100) /100);
    variableDataES.data.push(Math.round((UserData.output.Coop.variableCostUSPound * conversion) * 100) / 100);

    chartDataObjectES.push(variableDataES);
    chartDataObjectES.push(fixedDataES);
    chartDataObjectES.push(additionalDataES);

    language.ES.chart.data = chartDataObjectES
    language.ES.chart.plotlinePriceRecieved = Math.round((UserData.output.ProducerOutputEnglish.breakEvenCostUSPound * conversion) * 100) / 100
    language.ES.chart.plotlinePriceRecievedText = Math.round((UserData.output.ProducerOutputEnglish.breakEvenCostUSPound * conversion) * 100) / 100 + "<br /> precio <br/> recibido";
    language.ES.chart.plotlineWorldPrice = Math.round((UserData.output.ProducerOutputEnglish.futuresPrice * conversion) * 1000) / 1000
    language.ES.chart.plotlineWorldPriceText = Math.round((UserData.output.ProducerOutputEnglish.futuresPrice * conversion) * 1000) / 1000 + "<br />  precio  <br/> mundial";
}
