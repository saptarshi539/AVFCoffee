//global object for the language settings
var language = {
    "spanish": {
        "langLink1": "Inglés",
        "langLink2": "Español",
        "home-header": "Bienvenido a Calcucafé",
        "home-subheader": "Una herramienta para ayudarlo a calcular su costo de producción de café",
        "input-question1": "1. ¿Cuántas hectáreas de café tiene de acuerdo la edad de los árboles en su finca?",
        "input-question1-option1": "Producción​ ​inicial",
        "input-question1-option2": "Producción​ ​máximo",
        "input-question1-option3": "Viejo",
        "input-question2": "2. ¿Cuál es su método de cultivo?",
        "input-question2-option1": "Orgánico",
        "input-question2-option2": "Químico",
        "input-question2-option3": "En Transición ",
        "input-question3": "3. ¿Cuánto les paga a sus trabajadores por día?",
        "input-question3-label": "Trabajadores",
        "input-question3-units": "soles/día",
        "input-question4": "4. ¿Cuál es su productividad cuando sus arboles de café están maduros? ",
        "input-question4-label": "Producidos",
        "input-question4-units": "kilogramos/hectárea",
        "input-submit": "Ingresar",
        "input-question5": "5. ¿Cuánto paga en soles para transportar su café de la granja a el centro de recogida en un año?",
        "input-question5-label": "Transporte",
        "input-question5-units": "soles",
        "input-question6": "6. ¿Qué precio recibió por quintal de café?",
        "input-question6-label": "Precio",
        "input-question6-units": "soles",
        "chart": {
            chartTitle: "Desglose de costos",
            categories: ["Productor", "Cooperativa"],
            description: {
                "Variables": "Los costos variables son aquellos asociados a las cantidades de café producidas en la finca o parcela. Estos incluyen mano de obra y otros insumos necesarios para la producción de cantidades específicas de café.",
                "Fijos": "Los costos fijos tienen que ser pagados sin importaS el nivel de producción de café. Estos incluyen impuestos, membresías a cooperativa entre otros.",
                "Adicionales": "Los costos de depreciación y totales incluyen la depreciación de herramientas y equipos que se usan por más de un periodo así como el costo de oportunidad de los costos iniciales de establecimiento y de la tierra."
            },
            yaxisLabel: "Precio por kilogramo",
            plotLineLabel: "Precio<br/> Actual",
            data: [{
                name: 'Variables',
                data: [.84, .89]
            }, {
                name: 'Fijos',
                data: [.06, .04],
            }, {
                name: 'Adicionales',
                data: [.41, .44],
            }]
        }
    },

    "english": {
        "langLink1": "English",
        "langLink2": "Spanish",
        "home-header": "Welcome to Calcucafé",
        "home-subheader": "A tool to help you calculate your cost of coffee production",
        "input-question1": "1. How many hectares of each tree do you have?",
        "input-question1-option1": "Young",
        "input-question1-option2": "Mature",
        "input-question1-option3": "Old",
        "input-question2": "2. What is your method of Farming?",
        "input-question2-option1": "Organic",
        "input-question2-option2": "Chemical",
        "input-question2-option3": "Transitioning",
        "input-question3": "3. How much do you pay day laborers?",
        "input-question3-label": "Laborers",
        "input-question3-units": "soles/day",
        "input-question4": "4. What is your Productivity?",
        "input-question4-label": "Production",
        "input-question4-units": "pounds/day",
        "input-submit": "Submit",
        "input-question5": "5. How​ ​much​ ​do​ ​you​ ​pay​ ​in​ ​soles​ ​to​ ​transport​ ​your​ ​coffee​ ​from​ ​the​ ​farm​ ​to the​ ​collection​ ​center​ ​in​ ​one​ ​year?​",
        "input-question5-label": "Transport",
        "input-question5-units": "soles",
        "input-question6": "6. What​ ​price​ ​did​ ​you​ ​receive​ ​per​ ​quintal​ ​of​ ​coffee?",
        "input-question6-label": "Price",
        "input-question6-units": "soles",
        "chart": {
            chartTitle: "Your Farm",
            categories: ["Your Farm", "Co-op Average"],
            description: {
                "Variable": "Variable Costs are directly related to coffee farm output. These include hired labor and production inputs such as fertilizer or pesticides.",
                "Fixed": "Fixed costs must be paid whether or not any coffee is produced. These include cooperative memberships costs, taxes, and supplies.",
                "Additional": "Total costs includes the depreciation for assets used in more than one harvest cycles, start-up costs,and the opportunity costs of land and farm management."
            },
            yaxisLabel: "Price per pound",
            plotLineLabel: "Current<br/>Price"
        }
    }
}



function translate() { 
    console.log("translate")
    //filter the document to pull out just elements with a data-tag attribute
    var datas = $("*").filter("[data-tag]")
    var selected = localStorage.getItem("selectedLanguage")

   //iterate through the data-tags, lookup the lang value and update the element
    datas.each(function (i, e) {
        var data = $(this).attr("data-tag")
        $(this).html(language[selected][data])
    }) 
}


// click event for front page set language links
$("#english").click(function () {
    console.log("eng")
    localStorage.setItem("selectedLanguage", "english")
    translate()
});

$("#spanish").click(function () {
    console.log("eng")
    localStorage.setItem("selectedLanguage", "spanish")
    translate();
});

// on each ppage load, translate to the selected languaage
$(document).ready(function () {
    translate();
})