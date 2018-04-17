//global object for the language settings
var language = {
    "ES": {
        "langLink1": "Inglés",
        "langLink2": "Español",
        "home-header": "Bienvenido a Calcucafé",
        "home-subheader": "Una herramienta para ayudarlo a calcular su costo de producción de café",
        "home-demobtn": "Regístrate",
        "home-loginbtn": "Iniciar sesión",
        "demo-screen1": "Se empieza a ingresar sus datos básicos",
        "demo-screen2": "Entónces, se va a ver su deglose de costo",
        "demo-screen3": "Su desglose de costos",
        "demo-screen4": "Su desglose de costos",
        "demo-screen5": "Su desglose de costos",
        "demo-screen6": "Por favor, regístrese para empezar",
        "demo-next": "Siguiente",
        "demo-skip": "Omitir",
        "demo-signup": "Regístrate",
        "layout-navitem1": "Inicio",
        "layout-navitem2": "Datos",
        "layout-navitem3": "Simulación",
        "layout-navitem4": "Cerrar Session",
        "input-question1": "1. ¿Cuántas hectáreas de café tiene de acuerdo la edad de los árboles en su finca?",
        "input-question1-option1": "Hectáreas con árboles jóvenes",
        "input-question1-option1-tooltip": "Arboles que están dando sus primeros frutos.",
        "input-question1-option2": "Hectáreas con árboles maduros",
        "input-question1-option2-tooltip": "Arboles que están dando el máximo de su producción.",
        "input-question1-option3": "Hectáreas con árboles viejos",
        "input-question1-option3-tooltip": "Arboles que están dando menos que en el pasado.",
        "input-question2": "2. ¿Cuál es su método de cultivo?",
        "input-question2-option1": "Orgánico",
        "input-question2-option1-tooltip": "Producción con métodos orgánicos.",
        "input-question2-option2": "Convenciónal",
        "input-question2-option2-tooltip": "Producción en la cual utiliza químicos.",
        "input-question2-option3": "En Transición ",
        "input-question2-option3-tooltip": "Está en el proceso de pasarse de sistema de producción químico a orgánico.",
        "input-question3": "3. ¿Cuánto le paga a sus trabajadores por día?",
        "input-question3-label": "Trabajadores",
        "input-question3-units": "soles/día",
        "input-question3-option1-tooltip": "El sueldo promedio que sus trabajadores ganan en un día. Ejemplo: Jornal.",
        "input-question4": "4. ¿En un año cuántos quintales de café produce en promedio en una hectárea de árboles maduros?",
        "input-question4-label": "Producidos",
        "input-question4-units": "quintales/hectárea",
        "input-question4-option1-tooltip": "Kilogramos de café producidos por hectárea durante un periodo de cosecha normal.",
        "input-submit": "Ingresar",
        "input-question5": "5. ¿ En un año cuánto paga en soles para transportar su café de la granja a el centro de recogida ?",
        "input-question5-label": "Transporte",
        "input-question5-units": "soles/año",
        "input-question5-option1-tooltip": "¿Cuánto paga para el transporte?",
        "input-question6": "6. ¿Qué precio recibió por quintal de café?",
        "input-question6-label": "Precio",
        "input-question6-units": "soles/quintal",
        "input-question6-option1-tooltip": "Ingrese el precio que recibió.",
        "input-question7": "7. En un año de producción normal cuánto gasta en los siguientes insumos ?",
        "input-question7-option1": "Fertilizantes químicos",
        "input-question7-option1-tooltip": "Ingrese el precio gastado en fertilizante por hectárea.",
        "input-question7-option2": "Fertilizantes orgánicos",
        "input-question7-option2-tooltip": "Ingrese el precio gastado en fertilizante por hectárea.",
        "simulation-header1": "Hectárea:",
        "simulation-header2": "Método:",
        "simulation-header3": "Trabajadores:",
        "simulation-header4": "Producción:",
        "simulation-header5": "Transporte:",
        "simulation-header6": "Price:",
        "simulation-header7": "Gasto:",
        "chart": {
            chartTitle: "Desglose de costos",
            chartSubtitle: "Desplácese o haga clic en el gráfico para ver la definición de desglose de costos",
            categories: ["Productor", "Cooperativa"],
            simulationCategories: ["Productor", "Simulación"],
            yaxisLabel: {
                "ES": "Soles por Quintal",
                "EN": "Dólares por Libra"
            },
            plotLineLabel: "Precio<br/> Actual",
            seriesLabel: {
                "Variable": {
                    "name": "Variables",
                    "description": "Los costos variables son aquellos asociados a las cantidades de café producidas en la finca o parcela. Estos incluyen mano de obra y otros insumos necesarios para la producción de cantidades específicas de café."
                },
                "Fixed": {
                    "name": "Fijos",
                    "description": "Los costos fijos tienen que ser pagados sin importaS el nivel de producción de café.Estos incluyen impuestos, membresías a cooperativa entre otros.",
                },
                "Additional": {
                    "name": "Adicionales",
                    "description": "Los costos de depreciación y totales incluyen la depreciación de herramientas y equipos que se usan por más de un periodo así como el costo de oportunidad de los costos iniciales de establecimiento y de la tierra."
                }

            },
            data: {
                "EN": [],
                "ES": []
            },
            simulationData: {
                "EN": [],
                "ES": []
            },
            defaultUnits: "ES",
            altUnits: "EN",
            plotlinePriceRecieved: {
                "EN": "",
                "ES": ""
            },
            plotlineWorldPrice: {
                "EN": "",
                "ES": ""
            },
            plotlinePriceRecievedText: {
                "EN": "",
                "ES": ""
            },
            plotlineWorldPriceText: {
                "EN": "",
                "ES": ""
            },
            chartUnitsConversion: (101.4) * (3.16)
        },
        "chart-unitswitcher": "Ver unidades en: ",
        "chart-altunits": "Dólares por Libra",
        "chart-mainunits": "Soles por Quintal"

    },

    "EN": {
        "langLink1": "English",
        "langLink2": "Spanish",
        "home-header": "Welcome to Calcucafé",
        "home-subheader": "A tool to help you calculate your cost of coffee production",
        "home-demobtn": "Sign Up",
        "home-loginbtn": "Login",
        "demo-screen": "Start by inputting some basic information",
        "demo-screen2": "Then you will see your cost breakdown",
        "demo-screen3": "Your Cost Breakdown",
        "demo-screen4": "Your Cost Breakdown",
        "demo-screen5": "Your Cost Breakdown",
        "demo-screen6": "Please create an account to get started",
        "demo-next": "Next",
        "demo-skip": "Skip",
        "demo-signup": "Sign Up",
        "layout-navitem1": "Home",
        "layout-navitem2": "Input",
        "layout-navitem3": "Simulation",
        "layout-navitem4": "Sign Out",
        "input-question1": "1. How many hectares of each tree do you have?",
        "input-question1-option1": "Young",
        "input-question1-option1-tooltip": "Trees that are not yet producing beans.",
        "input-question1-option2": "Mature",
        "input-question1-option2-tooltip": "Trees that are fully producing.",
        "input-question1-option3": "Old",
        "input-question1-option3-tooltip": "Trees that are past peak production.",
        "input-question2": "2. What is your method of Farming?",
        "input-question2-option1": "Organic",
        "input-question2-option1-tooltip": "Organic production methods.",
        "input-question2-option2": "Conventional",
        "input-question2-option2-tooltip": "Chemical methods of production.",
        "input-question2-option3": "Transition",
        "input-question2-option3-tooltip": "Transitioning to organic production methods.",
        "input-question3": "3. How much do you pay per day to your workers in soles on average?",
        "input-question3-label": "Laborers",
        "input-question3-units": "soles/day",
        "input-question3-option1-tooltip": "How much do you pay for labor.",
        "input-question4": "4. How many quintales of coffee do you produce on average in one year per hectare?",
        "input-question4-label": "Production",
        "input-question4-units": "quintals/hectare",
        "input-question4-option1-tooltip": "Enter your yield in quintales/day.",
        "input-submit": "Submit",
        "input-question5": "5. How​ ​much​ ​do​ ​you​ ​pay​ ​in​ ​soles​ ​to​ ​transport​ ​your​ ​coffee​ ​from​ ​the​ ​farm​ ​to the​ ​collection​ ​center​ ​in​ ​one​ ​year?​",
        "input-question5-label": "Transport",
        "input-question5-units": "soles/year",
        "input-question5-option1-tooltip": "How much do you pay for transport.",
        "input-question6": "6. What​ ​price​ ​did​ ​you​ ​receive​ ​per​ ​quintal​ ​of​ ​coffee?",
        "input-question6-label": "Price",
        "input-question6-units": "soles/quintal",
        "input-question6-option1-tooltip": "Enter the price you recieved.",
        "input-question7": " 7. In one year, and during the pick of production, how much did you spend in your coffee farm in the following inputs per hectare?",
        "input-question7-option1": "Conventional Fertilizers",
        "input-question7-option1-tooltip": "Enter price spent on fertilizer per hectacre.",
        "input-question7-option2": "Organic Fertilizers",
        "input-question7-option2-tooltip": "Enter price spent on fertilizer per hectacre.",
        "simulation-header1": "Hectares:",
        "simulation-header2": "Method:",
        "simulation-header3": "Laborers:",
        "simulation-header4": "Production:",
        "simulation-header5": "Transport:",
        "simulation-header6": "Price:",
        "simulation-header7": "Expenditure:",
        "chart": {
            chartTitle: "Your Farm",
            chartSubtitle: "Hover or click the chart for definition of cost breakdown",
            categories: ["Your Farm", "Co-op Average"],
            simulationCategories: ["Producer", "Simulation"],
            yaxisLabel: {
                "EN": "Dollars per Pound",
                "ES": "Soles per Quintal"
            },
            plotLineLabel: "Current<br/>Price",
            seriesLabel: {
                "Additional": {
                    "name": "Additional",
                    "description": "Total costs includes the depreciation for assets used in more than one harvest cycles, start-up costs,and the opportunity costs of land and farm management."
                },         
                "Fixed": {
                    "name": "Fixed",
                    "description": "Fixed costs must be paid whether or not any coffee is produced. These include cooperative memberships costs, taxes, and supplies."
                },
                "Variable": {
                    "name": "Variable",
                    "description": "Variable Costs are directly related to coffee farm output. These include hired labor and production inputs such as fertilizer or pesticides."
                }
               
            },
            data: {
                "EN": [],
                "ES": []
            },
            simulationData: {
                "EN": [],
                "ES": []
            },
            defaultUnits: "EN",
            altUnits: "ES",
            plotlinePriceRecieved: {
                "EN": "",
                "ES": ""
            },
            plotlineWorldPrice: {
                "EN": "",
                "ES": ""
            },
            plotlinePriceRecievedText: {
                "EN": "",
                "ES": ""
            },
            plotlineWorldPriceText: {
                "EN": "",
                "ES": ""
            },
            chartUnitsConversion: ""
        },
        "chart-unitswitcher": "View units in: ",
        "chart-altunits": "Soles per Quintal",
        "chart-mainunits": "Dollars per Pound"

    }
}
// on each ppage load, translate to the selected languaage
$(document).ready(function () {
    //if (page.toLowerCase() != "simulation") 
    var path = window.location.pathname;
    var page = path.split("/").pop();
    // default to spanish on app load
    // when user logs in their chosen language is used to access the login page in that language
    // the app will stay in that langaue until user clicks trnaslate buttons on top page
    // When app is loaded it will always load in the language that the user chose when logging in that session.
    console.log(page)
    console.log(localStorage.getItem("selectedLanguage"))
    if (page == '') {
        localStorage.setItem("selectedLanguage", "ES")
        localStorage.setItem("defaultUnits", "")
        translate();
    }
    else if (page == 'Demo') {
        localStorage.getItem("selectedLanguage")
        translate();
    }
    else {
        if (!localStorage.getItem("selectedLanguage")) {
            globalDataPromise.then(function (value) {
                selectedLanguage = UserData.user.language
                localStorage.setItem("selectedLanguage", selectedLanguage);
                localStorage.setItem("selectedUnits", language[selectedLanguage].chart.defaultUnits);
                translate();
            })
        }
        else{
            translate();
        }
    }
})

function translate() { 
    //filter the document to pull out just elements with a data-tag attribute
    var datas = $("*").filter("[data-tag]")
    var selected = localStorage.getItem("selectedLanguage")


   //iterate through the data-tags, lookup the lang value and update the element
    datas.each(function (i, e) {
        var data = $(this).attr("data-tag")
        var tooltip = data.split("-");

        if (tooltip[3] == 'tooltip') {
            $(this).attr('data-original-title', language[selected][data]);
        }
        else {
            $(this).html(language[selected][data])
        }
    }) 
}


// click event for front page set language links
$("#english").click(function () {
    localStorage.setItem("selectedLanguage", "EN")
    localStorage.setItem("selectedUnits", "EN")
    $(".switchUnitsAlt").show();
    $(".switchUnitsMain").hide();


    translate()
    if ($('#chartdiv1').highcharts()) {
        createResultChart();
    }
    if ($('#chartdiv2').highcharts()) {
        createSimulationChart();
    }
   
});

$("#spanish").click(function () {
    localStorage.setItem("selectedLanguage", "ES")
    localStorage.setItem("selectedUnits", "ES")
    $(".switchUnitsAlt").show();
    $(".switchUnitsMain").hide();


    translate();
    if ($('#chartdiv1').highcharts()) {
        createResultChart();
    }
    if ($('#chartdiv2').highcharts()) {
        createSimulationChart();
    }
});

//click event for switching units view in chart
$(".switchUnitsMain").click(function () {
    //get current lang
    var lang = localStorage.getItem("selectedLanguage")
    //get default units
    var units = language[lang].chart.defaultUnits;
    localStorage.setItem("selectedUnits", units);

    if ($('#chartdiv1').highcharts()) {
        createResultChart();
    }
    if ($('#chartdiv2').highcharts()) {
        createSimulationChart();
    }
    $(".switchUnitsAlt").show();
    $(".switchUnitsMain").hide();


});

$(".switchUnitsAlt").click(function () {
    //get current lang
    var lang = localStorage.getItem("selectedLanguage")
    //get alt units
    var units = language[lang].chart.altUnits;
    localStorage.setItem("selectedUnits", units);

    if ($('#chartdiv1').highcharts()) {
        createResultChart();
    }
    if ($('#chartdiv2').highcharts()) {
        createSimulationChart();
    }
    $(".switchUnitsMain").show();
    $(".switchUnitsAlt").hide();

});