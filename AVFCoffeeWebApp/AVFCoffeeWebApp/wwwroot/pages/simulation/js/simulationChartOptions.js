
function createSimulationChart() {
    //var lang = UserData.user.language
    var lang = localStorage.getItem("selectedLanguage")
    var units = localStorage.getItem("selectedUnits")
    var chartLanguage = language[lang]["chart"];
    var data = chartLanguage.simulationData[units]

    
     Highcharts.chart('chartdiv2', {
         colors: ["#009c86", "#05354b", "#96394e", "#0D8ECF", "#2A0CD0", "#CD0D74", "#CC0000", "#00CC00", "#0000CC", "#DDDDDD", "#999999", "#333333", "#990000"],

        chart: {
            type: 'column',
            marginBottom: 100,
            marginTop: 100,
            backgroundColor: '#EFEFEF',
        },
        title: {
            text: chartLanguage.chartTitle
        },
        subtitle: {
            text: chartLanguage.chartSubtitle
        },
        xAxis: {
            categories: chartLanguage.simulationCategories
        },
        yAxis: {
            min: 0,
            title: {
                text: language[lang]["chart"].yaxisLabel[units]
            },
            stackLabels: {
                enabled: true,
                style: {
                    fontWeight: 'bold',
                    color: (Highcharts.theme && Highcharts.theme.textColor) || 'gray'
                },
                formatter: function () {
                    return Highcharts.numberFormat(this.total,2)
                }
            },
            plotLines: [{
                color: '#05354b',
                value: chartLanguage.plotlinePriceRecieved, // Insert your average here
                width: '1',
                zIndex: 99, // To not get stuck below the regular plot lines,
                dashStyle: 'ShortDash',
                label: {
                    text: chartLanguage.plotlinePriceRecievedText,
                    align: 'right',
                    textAlign: 'right',
                    style: {
                        color: '#05354b',
                        fontWeight: 'bold',
                    },
                }
            }, {
                color: '#96394e',
                value: chartLanguage.plotlineWorldPrice,
                width: '1',
                zIndex: 99, // To not get stuck below the regular plot lines,
                dashStyle: 'ShortDash',
                label: {
                    text: chartLanguage.plotlineWorldPriceText,
                    align: 'left',
                    textAlign: 'left',
                    style: {
                        color: '#96394e',
                        fontWeight: 'bold',
                    },
                }
            }
            ]
        },
        legend: {
            align: 'center',
            verticalAlign: 'bottom',
            floating: true,
            backgroundColor: (Highcharts.theme && Highcharts.theme.background2) || 'white',
            borderColor: '#CCC',
            borderWidth: 1,
            shadow: false,
            labelFormatter: function () {
                return chartLanguage.seriesLabel[this.name].name
            }
        },
        tooltip: {
            headerFormat: '<b>{point.x}</b><br/>',
            //pointFormat: '{series.name}: {point.y}<br/>Total: {point.stackTotal}<br/>description[{series.name}]',
            formatter: function () {
                return '<b>' + chartLanguage.seriesLabel[this.series.name].name + '</b>: ' + chartLanguage.seriesLabel[this.series.name].description;
            }
        },
        plotOptions: {
            column: {
                stacking: 'normal',
                dataLabels: {
                    enabled: true,
                    color: (Highcharts.theme && Highcharts.theme.dataLabelsColor) || 'white'
                }
            }
        },
        series: data


    });
}


