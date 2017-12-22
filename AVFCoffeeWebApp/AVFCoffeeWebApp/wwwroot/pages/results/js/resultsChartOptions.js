
function createResultChart() {
    
    var lang = UserData.user.language
    var chartLanguage = language[lang]["chart"];
    var chartData = JSON.parse(localStorage.getItem("chartDataObject"))
    console.log(chartLanguage)

    Highcharts.chart('chartdiv1', {
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
            text: UserData.user.language.chartTitle
        },
        xAxis: {
            categories: UserData.user.language.categories

        },
        yAxis: {
            min: 0,
            title: {
                text: UserData.user.language.yaxisLabel
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
                    text: UserData.user.language.plotLineLabel,
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
                return '<b>' + this.series.name + '</b>: ' + UserData.user.language.description[this.series.name];
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
        series: UserData.resultChartDataObject
    });
}