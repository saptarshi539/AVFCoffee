﻿
function createSimulationChart() {
    var lang = UserData.user.language
    var chartLanguage = language[lang]["chart"];
    
     Highcharts.chart('chartdiv2', {
        exporting: {
            chartOptions: { // specific options for the exported image
                plotOptions: {
                    series: {
                        dataLabels: {
                            enabled: true
                        }
                    }
                }
            },
            fallbackToExportServer: false
        },
        colors: ["#B9A5AE", "#9D6D82", "#754A5D", "#0D8ECF", "#2A0CD0", "#CD0D74", "#CC0000", "#00CC00", "#0000CC", "#DDDDDD", "#999999", "#333333", "#990000"],
        chart: {
            type: 'column',
            marginBottom: 100,
            marginTop: 100,
            backgroundColor: '#EFEFEF',
        },
        title: {
            text: chartLanguage.chartTitle
        },
        xAxis: {
            categories: chartLanguage.simulationCategories
        },
        yAxis: {
            min: 0,
            title: {
                text: chartLanguage.yaxisLabel
            },
            stackLabels: {
                enabled: true,
                style: {
                    fontWeight: 'bold',
                    color: (Highcharts.theme && Highcharts.theme.textColor) || 'gray'
                }
            },
            plotLines: [{
                color: 'blue',
                value: Math.round(UserData.output.ProducerOutputEnglish.breakEvenCostUSPound * 100) / 100, // Insert your average here
                width: '1',
                zIndex: 2, // To not get stuck below the regular plot lines,
                dashStyle: 'ShortDash',
                label: {
                    text: Math.round(UserData.output.ProducerOutputEnglish.breakEvenCostUSPound * 100) / 100,
                    style: {
                        textAlign: 'right',
                        color: 'blue',
                        fontWeight: 'bold',
                    },
                    x: -30
                }
            }, {
                color: 'red',
                value: 1.34,
                width: '1',
                zIndex: 2, // To not get stuck below the regular plot lines,
                dashStyle: 'ShortDash',
                label: {
                    text: 1.34,
                    style: {
                        textAlign: 'right',
                        color: 'red',
                        fontWeight: 'bold',
                    },
                    x: -30
                }
            }]
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
        series: UserData.simulationChartDataObject


    });
}


