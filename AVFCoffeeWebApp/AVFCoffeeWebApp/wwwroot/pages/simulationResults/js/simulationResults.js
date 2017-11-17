// get language object for rendering chart 
var selected = localStorage.getItem("selectedLanguage")
var chartLanguage = language[selected]["chart"]
var chartData = JSON.parse(localStorage.getItem("chartDataObject"))


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
        categories: ["Producer", "Simulation", "Cooperative"]

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
            color: 'black',
            value: '1.11', // Insert your average here
            width: '1',
            zIndex: 2, // To not get stuck below the regular plot lines,
            dashStyle: 'ShortDash',
            label: {
                text: chartLanguage.plotLineLabel,
                style: {
                    textAlign: 'right',
                    color: 'black',
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
        shadow: false
    },
    tooltip: {
        headerFormat: '<b>{point.x}</b><br/>',
        //pointFormat: '{series.name}: {point.y}<br/>Total: {point.stackTotal}<br/>description[{series.name}]',
        formatter: function () {
            return '<b>' + this.series.name + '</b>: ' + chartLanguage.description[this.series.name];
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
    series: [{
            name: 'Variable',
            data: [.84, .81, .89]
        }, {
            name: 'Fixed',
            data: [.06, .03, .04],
        }, {
            name: 'Additional',
            data: [.41, .40, .38],
        }]
    
    
});



