var data;
var globalData;
var prom;
var lang = localStorage.getItem("selectedLanguage");
$(document).on('ready', function () {
    debugger;
    console.log(lang);
    getInputs(apiURL);
    //prom.then(getAdvanced(apiURL));
});

function getInputs(apiURL) {
    debugger;
    $("#Inputs").html("");
    var htmlStr = "";
    var language = "";
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" style="margin-left:24px">' +
            'All Categories' +
            '</div >' +
            '</div >';
        language = lang;
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" style="margin-left:24px">' +
            'Todas las categorias' +
            '</div >' +
            '</div >';
        language = lang;
    }
    var c = 1;
    prom = $.ajax({
        type: "GET",
        url: apiURL + "TechnicianHomeAPI/getinputs",
        data: "language=" + language,
        contentType: "application/json",
        success: function (result) {
            //console.log(result);
            data = result;
            globalData = result;
            for (var key in result) {
                if (key != "Labordesc" && key != "Costdesc" && key != "Equipmentdesc" && key != "Transportationdesc" && key != "Inputs") {
                    htmlStr = htmlStr +

                        '<a style="cursor:pointer" onclick="Inputs' + c + '()">' +
                        '<div class="list-item">' +
                        '<div class="list-title">' + key + '</div>' +
                        '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                        '</div>' +
                        '</a >'
                    c++;
                }
            }
            $("#Inputs").html(htmlStr);
            getAdvanced(apiURL)
            //        for (i = 0; i < result.length; i++ ) {

            //}
        }
    });

}

function Inputs1() {
    debugger;
    var c = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Labor</a>' +
            '</div>' +
            '</div>' +
            '</div>';
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Labor</a>' +
            '</div>' +
            '</div>' +
            '</div>';
    }
    for (var key in data["Labor"]) {

        htmlStr = htmlStr +

            '<a style="cursor:pointer" onclick="LaborInputs' + c + '()">' +
            '<div class="list-item">' +
            '<div class="list-title">' + key + '</div>' +
            '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
            '</div>' +
            '</a >'
        c++;
    }
    $("#Inputs").html(htmlStr);
    data = data["Labor"];
}

function Inputs2() {
    debugger;
    var c = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Additional Income</a>' +
            '</div>' +
            '</div>' +
            '</div>';

        for (var key in data["Additional Income and remunertion"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="AddInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Ingresos adicionales</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var key in data["Ingreso adicional y remuneraciones"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="AddInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    }
    $("#Inputs").html(htmlStr);
    if (lang === "EN") {
        data = data["Additional Income and remunertion"];
    } else {
        data = data["Ingreso adicional y remuneraciones"];
    }
}

function Inputs3() {
    debugger;
    var c = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Cost</a>' +
            '</div>' +
            '</div>' +
            '</div>';

        for (var key in data["Cost of materials and inputs"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="CostInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Costo</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var key in data["Costo de Materiales e insumos"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="CostInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    }
    $("#Inputs").html(htmlStr);
    if (lang === "EN") {
        data = data["Cost of materials and inputs"];
    } else {
        data = data["Costo de Materiales e insumos"];
    }
}

function Inputs4() {
    debugger;
    var c = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Equipment</a>' +
            '</div>' +
            '</div>' +
            '</div>';

        for (var key in data["Equipment and Reusable material"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="EquipmentInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Equipo</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var key in data["Equipo y material reutilizable"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="EquipmentInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    }
    $("#Inputs").html(htmlStr);
    if (lang === "EN") {
        data = data["Equipment and Reusable material"];
    } else {
        data = data["Equipo y material reutilizable"];
    }
}

function Inputs5() {
    debugger;
    var c = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Administrative costs</a>' +
            '</div>' +
            '</div>' +
            '</div>';

        for (var key in data["Administrative costs, taxes and land"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="AdminInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Costes administrativos</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var key in data["Costos administrativos, impuestos y tierra"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="AdminInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    }
    $("#Inputs").html(htmlStr);
    if (lang === "EN") {
        data = data["Administrative costs, taxes and land"];
    } else {
        data = data["Costos administrativos, impuestos y tierra"];
    }
}

function Inputs6() {
    debugger;
    var c = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Transportation</a>' +
            '</div>' +
            '</div>' +
            '</div>';

        for (var key in data["Transportation"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="TransInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    } else {

        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back">' +
            '<a href=""><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Transporte</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var key in data["Transporte"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="TransInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    }
    $("#Inputs").html(htmlStr);
    if (lang === "EN") {
        data = data["Transportation"];
    } else {
        data = data["Transporte"];
    }
}

function CostInputs1() {
    debugger;
    var cost = 1;
    var instr = globalData["Costdesc"];
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs3()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs3()" style="cursor:pointer">Cost</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Materials for germinator</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (cost === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=costinputsfunc1() id="costinputs1' + i + '" value="' + costinputs1[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            cost++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button data-tag="DiscardButton" class="outlined-button-disabled" type="button"> DISCARD CHANGES</button>' +
            '<button data-tag="SaveButton" class="contained-button-disabled" id="costinputssave1" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs3()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs3()" style="cursor:pointer">Costo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Materiales para germinador</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (cost === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=costinputsfunc1() id="costinputs1' + i + '" value="' + costinputs1[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            cost++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="costinputssave1" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function CostInputs2() {
    debugger;
    var instr = globalData["Costdesc"];
    var cost = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs3()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text"style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs3()" style="cursor:pointer">Cost</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Materials for nursery</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (cost === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=costinputsfunc2() id="costinputs2' + i + '" value="' + costinputs2[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            cost++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button data-tag="DiscardButton" class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button data-tag="SaveButton" class="contained-button-disabled" id="costinputssave2" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs3()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text"style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs3()" style="cursor:pointer">Costo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Materiales para vivero</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (cost === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=costinputsfunc2() id="costinputs2' + i + '" value="' + costinputs2[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            cost++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="costinputssave2" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function CostInputs3() {
    debugger;
    var cost = 1;
    var instr = globalData["Costdesc"];
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs3()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs3()" style="cursor:pointer">Cost</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Fertilizers during planting</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (cost === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=costinputsfunc3() id="costinputs3' + i + '" value="' + costinputs3[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            cost++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button data-tag="DiscardButton" class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button data-tag="SaveButton" class="contained-button-disabled" id="costinputssave3" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs3()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs3()" style="cursor:pointer">Costo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Fertilizantes durante la siembra</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (cost === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=costinputsfunc3() id="costinputs3' + i + '" value="' + costinputs3[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            cost++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="costinputssave3" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function CostInputs4() {
    debugger;
    var cost = 1;
    var instr = globalData["Costdesc"];
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs3()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs3()" style="cursor:pointer">Cost</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Fertilizers during vegetative growth</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (cost === 4) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=costinputsfunc4() id="costinputs4' + i + '" value="' + costinputs4[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            cost++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button data-tag="DiscardButton" class="outlined-button-disabled" type="button"> DISCARD CHANGES</button>' +
            '<button data-tag="SaveButton" class="contained-button-disabled" id="costinputssave4" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs3()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs3()" style="cursor:pointer">Costo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Fertilizantes durante el crecimiento vegetativo</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (cost === 4) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=costinputsfunc4() id="costinputs4' + i + '" value="' + costinputs4[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            cost++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="costinputssave4" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function CostInputs5() {
    debugger;
    var cost = 1;
    var instr = globalData["Costdesc"];
    $("#Inputs").html("");
    var htmlStr = "";
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs3()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs3()" style="cursor:pointer">Cost</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Fertilizers during maintenance</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (cost === 5) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=costinputsfunc5() id="costinputs5' + i + '" value="' + costinputs5[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            cost++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button data-tag="DiscardButton" class="outlined-button-disabled" type="button"> DISCARD CHANGES</button>' +
            '<button data-tag="SaveButton" class="contained-button-disabled" id="costinputssave5" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs3()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs3()" style="cursor:pointer">Costo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Fertilizantes durante el mantenimiento</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (cost === 5) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=costinputsfunc5() id="costinputs5' + i + '" value="' + costinputs5[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            cost++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="costinputssave5" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function EquipmentInputs1() {
    debugger;
    var equipment = 1;
    var instr = globalData["Equipmentdesc"];
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back"style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs4()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs4()" style="cursor:pointer">Equipment</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>General Equipment</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (equipment === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=equipmentinputsfunc1() id="equipmentinputs1' + i + '" value="' + equipmentinputs1[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            equipment++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button"> DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="equipmentinputssave1" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back"style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs4()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs4()" style="cursor:pointer">Equipo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Equipo general</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (equipment === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=equipmentinputsfunc1() id="equipmentinputs1' + i + '" value="' + equipmentinputs1[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            equipment++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="equipmentinputssave1" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function EquipmentInputs2() {
    debugger;
    var equipment = 1;
    var instr = globalData["Equipmentdesc"];
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs4()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs4()" style="cursor:pointer">Equipment</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Equipments for Harvest</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (equipment === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=equipmentinputsfunc2() id="equipmentinputs2' + i + '" value="' + equipmentinputs2[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            equipment++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button data-tag="DiscardButton" class="outlined-button-disabled" type="button"> DISCARD CHANGES</button>' +
            '<button data-tag="SaveButton" class="contained-button-disabled" id="equipmentinputssave2" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back"style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs4()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs4()" style="cursor:pointer">Equipo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Equipos para cosecha</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (equipment === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=equipmentinputsfunc2() id="equipmentinputs2' + i + '" value="' + equipmentinputs2[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            equipment++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="equipmentinputssave2" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function EquipmentInputs3() {
    debugger;
    var equipment = 1;
    var instr = globalData["Equipmentdesc"];
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs4()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs4()" style="cursor:pointer">Equipment</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Equipments for processing</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (equipment === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=equipmentinputsfunc3() id="equipmentinputs3' + i + '" value="' + equipmentinputs3[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            equipment++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button data-tag="DiscardButton" class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button data-tag="SaveButton" class="contained-button-disabled" id="equipmentinputssave3" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back"style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs4()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs4()" style="cursor:pointer">Equipo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Equipos para procesamiento</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (equipment === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=equipmentinputsfunc3() id="equipmentinputs3' + i + '" value="' + equipmentinputs3[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            equipment++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="equipmentinputssave3" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function AdminInputs1() {
    debugger;
    var admin = 1;
    $("#Inputs").html("");
    var htmlStr = "";
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs5()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs5()" style="cursor:pointer">Administrative costs</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Cooperative Expenses</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (admin === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=admininputsfunc1() id="admininputs1' + i + '" value="' + admininputs1[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            admin++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button data-tag="DiscardButton" class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button data-tag="SaveButton" class="contained-button-disabled" id="admininputssave1" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);

    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs5()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs5()" style="cursor:pointer">Costes administrativos</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Gastos cooperativos</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (admin === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=admininputsfunc1() id="admininputs1' + i + '" value="' + admininputs1[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            admin++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="admininputssave1" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function AdminInputs2() {
    debugger;
    var admin = 1;
    $("#Inputs").html("");
    var htmlStr = "";
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs5()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs5()" style="cursor:pointer">Administrative costs</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>land</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (admin === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=admininputsfunc2() id="admininputs2' + i + '" value="' + admininputs2[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            admin++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="admininputssave2" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs5()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs5()" style="cursor:pointer">Costes administrativos</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>tierra</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (admin === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=admininputsfunc2() id="admininputs2' + i + '" value="' + admininputs2[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            admin++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="admininputssave2" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function AdminInputs3() {
    debugger;
    var admin = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs5()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs5()" style="cursor:pointer">Administrative costs</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Unexpected events</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (admin === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=admininputsfunc3() id="admininputs3' + i + '" value="' + admininputs3[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            admin++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="admininputssave3" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs5()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs5()" style="cursor:pointer">Costes administrativos</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Eventos inesperados</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (admin === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=admininputsfunc3() id="admininputs3' + i + '" value="' + admininputs3[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            admin++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="admininputssave3" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}


function TransInputs1() {
    debugger;
    var trans = 1;
    var instr = globalData["Transportationdesc"];
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs6()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs6()" style="cursor:pointer">Transport</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Germinator Transport</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (trans === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=transinputsfunc1() name="" id="transinputs1' + i + '" value="' + transinputs1[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            trans++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="transinputssave1" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs6()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs6()" style="cursor:pointer">Transporte</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Transporte del germinador</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (trans === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=transinputsfunc1() name="" id="transinputs1' + i + '" value="' + transinputs1[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            trans++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="transinputssave1" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function TransInputs2() {
    debugger;
    var trans = 1;
    var htmlStr = "";
    var instr = globalData["Transportationdesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs6()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs6()" style="cursor:pointer">Transport</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Nursery Transport</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (trans === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=transinputsfunc2() name="" id="transinputs2' + i + '" value="' + transinputs2[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            trans++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="transinputssave2" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs6()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs6()" style="cursor:pointer">Transporte</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Transporte de guardería</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (trans === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=transinputsfunc2() name="" id="transinputs2' + i + '" value="' + transinputs2[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            trans++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="transinputssave2" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function TransInputs3() {
    debugger;
    var trans = 1;
    var htmlStr = "";
    var instr = globalData["Transportationdesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs6()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs6()" style="cursor:pointer">Transport</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Land preparation and planting Transport</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (trans === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=transinputsfunc3() name="" id="transinputs3' + i + '" value="' + transinputs3[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            trans++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type = "button" > DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="transinputssave3" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs6()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs6()" style="cursor:pointer">Transporte</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Preparación de la tierra y siembra Transporte</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (trans === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=transinputsfunc3() name="" id="transinputs3' + i + '" value="' + transinputs3[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            trans++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="transinputssave3" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function TransInputs4() {
    debugger;
    var trans = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    var instr = globalData["Transportationdesc"];
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs6()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs6()" style="cursor:pointer">Transport</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Other Transport</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (trans === 4) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=transinputsfunc4() id="transinputs4' + i + '" value="' + transinputs4[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            trans++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="transinputssave4" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs6()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs6()" style="cursor:pointer">Transporte</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Otro transporte</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (trans === 4) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=transinputsfunc4() id="transinputs4' + i + '" value="' + transinputs4[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            trans++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="transinputssave4" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function AddInputs1() {
    debugger;
    var additional = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs2()" style="cursor:pointer">Additional Income</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Indirect Income</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (additional === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=addinputsfunc1() id="addinputs1' + i + '" value="' + addinputs1[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            additional++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="addinputssave1" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs2()" style="cursor:pointer">Ingresos adicionales</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Ingresos indirectos</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (additional === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=addinputsfunc1() id="addinputs1' + i + '" value="' + addinputs1[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            additional++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="addinputssave1" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}

function AddInputs2() {
    debugger;
    var additional = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs2()" style="cursor:pointer">Additional Income</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Credit</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (additional === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=addinputsfunc2() id="addinputs2' + i + '" value="' + addinputs2[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            additional++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="addinputssave2" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=Inputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick="Inputs2()" style="cursor:pointer">Ingresos adicionales</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Crédito</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (additional === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=addinputsfunc2() id="addinputs2' + i + '" value="' + addinputs2[i] + '">' +
                        '</form>' +
                        '</div>'
                }
            }
            additional++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="addinputssave2" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData;
}



function LaborInputs1() {
    debugger;
    console.log(data);
    var c = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=BackInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick=BackInputs1() style="cursor:pointer">Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Establishment and vegetative growth years</a>' +
            '</div>' +
            '</div>' +
            '</div>';

        for (var key in data["Labor during establishment and vegetative growth years"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="LaborGerminatorInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=BackInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick=BackInputs1() style="cursor:pointer">Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Establecimiento y años de crecimiento vegetativo</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var key in data["Trabajo durante los años de establecimiento y crecimiento de las plantas de café "]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="LaborGerminatorInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    }
    $("#Inputs").html(htmlStr);
    if (lang === "EN") {
        data = data["Labor during establishment and vegetative growth years"];
    } else {
        data = data["Trabajo durante los años de establecimiento y crecimiento de las plantas de café "];
    }
}
function BackInputs1() {
    data = globalData;
    Inputs1();
}
function LaborInputs2() {
    debugger;
    console.log(data);
    var c = 1;
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=BackInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a onclick=BackInputs1() style="cursor:pointer">Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Labor for applications</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var key in data["Labor for farm maintenance, harvesting and procesing"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="LaborMaintenanceInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=BackInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Mano de obra para aplicaciones</a>' +
            '</div>' +
            '</div>' +
            '</div>';
        for (var key in data["Trabajo para mantenimiento, cosecha y beneficio"]) {

            htmlStr = htmlStr +

                '<a style="cursor:pointer" onclick="LaborMaintenanceInputs' + c + '()">' +
                '<div class="list-item">' +
                '<div class="list-title">' + key + '</div>' +
                '<div class="list-chevron"><img src="/icons/chevron-right.svg" alt=""></div>' +
                '</div>' +
                '</a >'
            c++;
        }
    }

    $("#Inputs").html(htmlStr);
    if (lang === "EN") {
        data = data["Labor for farm maintenance, harvesting and procesing"];
    } else {
        data = data["Trabajo para mantenimiento, cosecha y beneficio"];
    }
}

function LaborMaintenanceInputs1() {
    debugger;
    var maintenance = 1;
    var htmlStr = "";
    var instr = globalData["Labordesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Labor for applications</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Maintain young coffee trees</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc1() id="labormaintenance1' + i + '" value="' + labormaintenance1[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button"> DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave1" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Mano de obra para aplicaciones</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Mantener cafetos jóvenes</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc1() id="labormaintenance1' + i + '" value="' + labormaintenance1[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave1" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];

}

function LaborMaintenanceInputs2() {
    debugger;
    var maintenance = 1;
    var htmlStr = "";
    var instr = globalData["Labordesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Labor for applications</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Harvest young coffee trees</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc2() id="labormaintenance2' + i + '" value="' + labormaintenance2[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave2" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Mano de obra para aplicaciones</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a href="">Cosechar cafetos jóvenes</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc2() id="labormaintenance2' + i + '" value="' + labormaintenance2[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave2" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

function LaborMaintenanceInputs3() {
    debugger;
    var maintenance = 1;
    var instr = globalData["Labordesc"];
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Labor for applications</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a href="">Process young coffee trees</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc3() id="labormaintenance3' + i + '" value="' + labormaintenance3[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave3" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Mano de obra para aplicaciones</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Procesar cafetos jóvenes</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc3() id="labormaintenance3' + i + '" value="' + labormaintenance3[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave3" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

function LaborMaintenanceInputs4() {
    debugger;
    var maintenance = 1;
    var htmlStr = "";
    var instr = globalData["Labordesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Labor for applications</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Maintain matured coffee trees</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 4) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc4() id="labormaintenance4' + i + '" value="' + labormaintenance4[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave4" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Mano de obra para aplicaciones</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Mantener cafetos maduros</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 4) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc4() id="labormaintenance4' + i + '" value="' + labormaintenance4[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave4" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

function LaborMaintenanceInputs5() {
    debugger;
    var maintenance = 1;
    var instr = globalData["Labordesc"];
    var htmlStr = "";
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Labor for applications</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Harvest matured coffee trees</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 5) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc5() id="labormaintenance5' + i + '" value="' + labormaintenance5[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave5" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Mano de obra para aplicaciones</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Cosechar cafetos maduros</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 5) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc5() id="labormaintenance5' + i + '" value="' + labormaintenance5[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave5" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

function LaborMaintenanceInputs6() {
    debugger;
    var maintenance = 1;
    var htmlStr = "";
    var instr = globalData["Labordesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Labor for applications</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Process matured coffee trees</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 6) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc6() id="labormaintenance6' + i + '" value="' + labormaintenance6[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave6" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Mano de obra para aplicaciones</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Procesar cafetales maduros</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 6) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=labormaintenancefunc6() id="labormaintenance6' + i + '" value="' + labormaintenance6[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave6" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

function LaborMaintenanceInputs7() {
    debugger;
    var maintenance = 1;
    var htmlStr = "";
    var instr = globalData["Labordesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Labor for applications</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Maintain old coffee trees</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 7) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=labormaintenancefunc7() name="" id="labormaintenance7' + i + '" value="' + labormaintenance7[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button"> DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave7" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {

        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Mano de obra para aplicaciones</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Mantener cafetos viejos</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 7) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=labormaintenancefunc7() name="" id="labormaintenance7' + i + '" value="' + labormaintenance7[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave7" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

function LaborMaintenanceInputs8() {
    debugger;
    var maintenance = 1;
    var htmlStr = "";
    var instr = globalData["Labordesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Labor for applications</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Harvest old coffee trees</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 8) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=labormaintenancefunc8() name="" id="labormaintenance8' + i + '" value="' + labormaintenance8[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave8" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Mano de obra para aplicaciones</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Cosecha árboles de café viejos</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 8) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=labormaintenancefunc8() name="" id="labormaintenance8' + i + '" value="' + labormaintenance8[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave8" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

function LaborMaintenanceInputs9() {
    debugger;
    var maintenance = 1;
    var htmlStr = "";
    var instr = globalData["Labordesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Labor for applications</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Process old coffee trees</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 9) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=labormaintenancefunc9() name="" id="labormaintenance9' + i + '" value="' + labormaintenance9[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave9" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs2()>Mano de obra para aplicaciones</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Procesar cafetales viejos</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (maintenance === 9) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=labormaintenancefunc9() name="" id="labormaintenance9' + i + '" value="' + labormaintenance9[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            maintenance++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="labormaintenancesave9" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

function LaborGerminatorInputs1() {
    debugger;
    var germinator = 1;
    var htmlStr = "";
    var instr = globalData["Labordesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()>Establishment and vegetative growth years</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Germinator</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (germinator === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=laborgerminatorfunc1() id="laborgerminatorinputs1' + i + '" value="' + laborgerminator1[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            germinator++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="laborgerminatorsave1" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()>Establecimiento y años de crecimiento vegetativo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Germinator</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (germinator === 1) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=laborgerminatorfunc1() id="laborgerminatorinputs1' + i + '" value="' + laborgerminator1[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            germinator++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="laborgerminatorsave1" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

function LaborGerminatorInputs2() {
    debugger;
    var germinator = 1;
    var htmlStr = "";
    var instr = globalData["Labordesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()>Establishment and vegetative growth years</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Nursery</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (germinator === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=laborgerminatorfunc2() id="laborgerminatorinputs2' + i + '" name="" value="' + laborgerminator2[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            germinator++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="laborgerminatorsave2" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()>Establecimiento y años de crecimiento vegetativo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Guardería</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (germinator === 2) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" onchange=laborgerminatorfunc2() id="laborgerminatorinputs2' + i + '" name="" value="' + laborgerminator2[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            germinator++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="laborgerminatorsave2" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

function LaborGerminatorInputs3() {
    debugger;
    var germinator = 1;
    var htmlStr = "";
    var instr = globalData["Labordesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()>Establishment and vegetative growth years</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Land Preparation labor</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (germinator === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=laborgerminatorfunc3() id="laborgerminatorinputs3' + i + '" value="' + laborgerminator3[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            germinator++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DISCARD CHANGES</button>' +
            '<button class="contained-button-disabled" id="laborgerminatorsave3" onclick="saveAdvanced()" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()>Establecimiento y años de crecimiento vegetativo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Trabajos de preparación de tierras</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (germinator === 3) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" type="text" name="" onchange=laborgerminatorfunc3() id="laborgerminatorinputs3' + i + '" value="' + laborgerminator3[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            germinator++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button class="contained-button-disabled" id="laborgerminatorsave3" onclick="saveAdvanced()" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

function LaborGerminatorInputs4() {
    debugger;
    var germinator = 1;
    var htmlStr = "";
    var instr = globalData["Labordesc"];
    $("#Inputs").html("");
    if (lang === "EN") {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">All Categories</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()>Establishment and vegetative growth years</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Vegetative growth labor</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (germinator === 4) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" onchange=laborgerminatorfunc4() type="text" name="" id="laborgerminatorinputs4' + i + '" value="' + laborgerminator4[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            germinator++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button"> DISCARD CHANGES</button>' +
            '<button onclick="saveAdvanced()" id="laborgerminatorsave4" class="contained-button-disabled" type="button">SAVE</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    } else {
        htmlStr = '<div class="breadcrumbs">' +
            '<div class="crumb" id="crumb-back" style="margin-left:24px">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()><img src="arrow-back.svg"></a>' +
            '</div>' +
            '<div class="crumb-text" style="margin-left:72px">' +
            '<div class="crumb">' +
            '<a href="">Todas las categorias</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=BackInputs1()>Labor</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a style="cursor:pointer" onclick=LaborInputs1()>Establecimiento y años de crecimiento vegetativo</a>' +
            '</div>' +
            '<div class="crumb">' +
            '<img src="/icons/chevron-right.svg">' +
            '</div>' +
            '<div class="crumb">' +
            '<a>Trabajo de crecimiento vegetativo</a>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="input-instructions">' +
            instr +
            '</div>';
        for (var d in data) {
            var g = data[d];
            if (germinator === 4) {
                for (var i = 0; i < g.length; i++) {
                    htmlStr = htmlStr +
                        '<div class="list-item">' +
                        '<div class="list-title-value">' + g[i] + '</div>' +
                        '<form class="input-form" autocomplete="off">' +
                        '<input class="input-value" onchange=laborgerminatorfunc4() type="text" name="" id="laborgerminatorinputs4' + i + '" value="' + laborgerminator4[i] + '">' +
                        //'days' +
                        '</form>' +
                        '</div>'
                }
            }
            germinator++;
        }
        htmlStr = htmlStr + '<div class="input-buttons">' +
            '<button class="outlined-button-disabled" type="button">DESCARTAR LOS CAMBIOS</button>' +
            '<button onclick="saveAdvanced()" id="laborgerminatorsave4" class="contained-button-disabled" type="button">SALVAR</button>' +
            '</div >'
        $("#Inputs").html(htmlStr);
    }
    data = globalData["Labor"];
}

//$("#laborgerminatorinputs40").change(function () {
//    alert("The text has been changed.");
//});

function laborgerminatorfunc4() {
    debugger;
    console.log("I am here");
    userAdvancedInputs["lppyWeeding"] = $("#laborgerminatorinputs40").val();
    userAdvancedInputs["lppyOrganic"] = $("#laborgerminatorinputs41").val();
    userAdvancedInputs["lppyChemical"] = $("#laborgerminatorinputs42").val();
    userAdvancedInputs["lppyFoliarSpraying"] = $("#laborgerminatorinputs43").val();
    userAdvancedInputs["lppyOther"] = $("#laborgerminatorinputs44").val();
    $("#laborgerminatorsave4").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function laborgerminatorfunc3() {
    debugger;
    console.log("I am here");
    userAdvancedInputs["lppFieldCleaning"] = $("#laborgerminatorinputs30").val();
    userAdvancedInputs["lppCuttingTrees"] = $("#laborgerminatorinputs31").val();
    userAdvancedInputs["lppWoodCollection"] = $("#laborgerminatorinputs32").val();
    userAdvancedInputs["lppWoodChopping"] = $("#laborgerminatorinputs33").val();
    userAdvancedInputs["lppCoffeeLayout"] = $("#laborgerminatorinputs34").val();
    userAdvancedInputs["lppHoleDigging"] = $("#laborgerminatorinputs35").val();
    userAdvancedInputs["lppSeedlingTransportation"] = $("#laborgerminatorinputs36").val();
    userAdvancedInputs["lppSeedlingTransplant"] = $("#laborgerminatorinputs37").val();
    userAdvancedInputs["lppShadeAdjustment"] = $("#laborgerminatorinputs38").val();
    userAdvancedInputs["lppCompostMixing"] = $("#laborgerminatorinputs39").val();
    userAdvancedInputs["lppOthers"] = $("#laborgerminatorinputs310").val();
    $("#laborgerminatorsave3").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function laborgerminatorfunc2() {
    debugger;
    console.log("I am here");
    userAdvancedInputs["lNurseryConstruction"] = $("#laborgerminatorinputs20").val();
    userAdvancedInputs["lNurseryDrawnPulled"] = $("#laborgerminatorinputs21").val();
    userAdvancedInputs["lNurseryClean"] = $("#laborgerminatorinputs22").val();
    userAdvancedInputs["lNurserySoilPreparationFertilizer"] = $("#laborgerminatorinputs23").val();
    userAdvancedInputs["lNurseryFilledLockedBags"] = $("#laborgerminatorinputs24").val();
    userAdvancedInputs["lNurseryButterflySowing"] = $("#laborgerminatorinputs25").val();
    userAdvancedInputs["lNurseryIrrigation"] = $("#laborgerminatorinputs26").val();
    userAdvancedInputs["lNurseryFoliarApplication"] = $("#laborgerminatorinputs27").val();
    userAdvancedInputs["lNurseryReseeding"] = $("#laborgerminatorinputs28").val();
    userAdvancedInputs["lNurseryOthers"] = $("#laborgerminatorinputs29").val();
    laborgerminator2[0] = $("#laborgerminatorinputs20").val();
    laborgerminator2[1] = $("#laborgerminatorinputs21").val();
    laborgerminator2[2] = $("#laborgerminatorinputs22").val();
    laborgerminator2[3] = $("#laborgerminatorinputs23").val();
    laborgerminator2[4] = $("#laborgerminatorinputs24").val();
    laborgerminator2[5] = $("#laborgerminatorinputs25").val();
    laborgerminator2[6] = $("#laborgerminatorinputs26").val();
    laborgerminator2[7] = $("#laborgerminatorinputs27").val();
    laborgerminator2[8] = $("#laborgerminatorinputs28").val();
    laborgerminator2[9] = $("#laborgerminatorinputs28").val();
    $("#laborgerminatorsave2").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function laborgerminatorfunc1() {
    debugger;
    console.log("I am here");
    userAdvancedInputs["lGerminationSeedCollection"] = $("#laborgerminatorinputs10").val();
    userAdvancedInputs["lGerminationSeedSelection"] = $("#laborgerminatorinputs11").val();
    userAdvancedInputs["lGerminationNurseryConstruction"] = $("#laborgerminatorinputs12").val();
    userAdvancedInputs["lGerminationSeedingSupportIrrigation"] = $("#laborgerminatorinputs13").val();
    userAdvancedInputs["lGerminationOthers"] = $("#laborgerminatorinputs14").val();
    laborgerminator1[0] = $("#laborgerminatorinputs10").val();
    laborgerminator1[1] = $("#laborgerminatorinputs11").val();
    laborgerminator1[2] = $("#laborgerminatorinputs12").val();
    laborgerminator1[3] = $("#laborgerminatorinputs13").val();
    laborgerminator1[4] = $("#laborgerminatorinputs14").val();
    $("#laborgerminatorsave1").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function labormaintenancefunc1() {
    userAdvancedInputs["lhpmyManualWeeding"] = $("#labormaintenance10").val();
    userAdvancedInputs["lhpmyChemicalWeeding"] = $("#labormaintenance11").val();
    userAdvancedInputs["lhpmyOrganicFertilizers"] = $("#labormaintenance12").val();
    userAdvancedInputs["lhpmyChemicalFertilizers"] = $("#labormaintenance13").val();
    userAdvancedInputs["lhpmyFoliarSpraying"] = $("#labormaintenance14").val();
    userAdvancedInputs["lhpmyHedgerowsConstruction"] = $("#labormaintenance15").val();
    userAdvancedInputs["lhpmyShadetreePruning"] = $("#labormaintenance16").val();
    userAdvancedInputs["lhpmyPestControl"] = $("#labormaintenance17").val();
    userAdvancedInputs["lhpmyCoffeeGrowManagement"] = $("#labormaintenance18").val();
    userAdvancedInputs["lhpmyOthers"] = $("#labormaintenance19").val();
    $("#labormaintenancesave1").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function labormaintenancefunc2() {
    userAdvancedInputs["lhphyCoffeeCollecDays"] = $("#labormaintenance20").val();
    userAdvancedInputs["lhphyAdditionDays"] = $("#labormaintenance21").val();
    $("#labormaintenancesave2").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function labormaintenancefunc3() {
    userAdvancedInputs["lhppyFermentation"] = $("#labormaintenance30").val();
    userAdvancedInputs["lhppyWashing"] = $("#labormaintenance31").val();
    userAdvancedInputs["lhppyDrying"] = $("#labormaintenance32").val();
    userAdvancedInputs["lhppyScreening"] = $("#labormaintenance33").val();
    userAdvancedInputs["lhppySelection"] = $("#labormaintenance34").val();
    userAdvancedInputs["lhppyStorage"] = $("#labormaintenance35").val();
    userAdvancedInputs["lhppyCoffeewastewater"] = $("#labormaintenance36").val();
    userAdvancedInputs["lhppyPulpManagement"] = $("#labormaintenance37").val();
    userAdvancedInputs["lhppyOthers"] = $("#labormaintenance38").val();
    $("#labormaintenancesave3").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function labormaintenancefunc4() {
    userAdvancedInputs["lhpmmManualWeeding"] = $("#labormaintenance40").val();
    userAdvancedInputs["lhpmmChemicalWeeding"] = $("#labormaintenance41").val();
    userAdvancedInputs["lhpmmOrganicFertilizers"] = $("#labormaintenance42").val();
    userAdvancedInputs["lhpmmChemicalFertilizers"] = $("#labormaintenance43").val();
    userAdvancedInputs["lhpmmFoliarSpraying"] = $("#labormaintenance44").val();
    userAdvancedInputs["lhpmmHedgerowsConstruction"] = $("#labormaintenance45").val();
    userAdvancedInputs["lhpmmShadetreePruning"] = $("#labormaintenance46").val();
    userAdvancedInputs["lhpmmPestControl"] = $("#labormaintenance47").val();
    userAdvancedInputs["lhpmmCoffeeGrowManagement"] = $("#labormaintenance48").val();
    userAdvancedInputs["lhpmmOthers"] = $("#labormaintenance49").val();
    $("#labormaintenancesave4").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function labormaintenancefunc5() {
    userAdvancedInputs["lhphmCoffeeCollecDays"] = $("#labormaintenance50").val();
    userAdvancedInputs["lhphmAdditionDays"] = $("#labormaintenance51").val();
    $("#labormaintenancesave5").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function labormaintenancefunc6() {
    userAdvancedInputs["lhppmFermentation"] = $("#labormaintenance60").val();
    userAdvancedInputs["lhppmWashing"] = $("#labormaintenance61").val();
    userAdvancedInputs["lhppmDrying"] = $("#labormaintenance62").val();
    userAdvancedInputs["lhppmScreening"] = $("#labormaintenance63").val();
    userAdvancedInputs["lhppmSelection"] = $("#labormaintenance64").val();
    userAdvancedInputs["lhppmStorage"] = $("#labormaintenance65").val();
    userAdvancedInputs["lhppmCoffeewastewater"] = $("#labormaintenance66").val();
    userAdvancedInputs["lhppmPulpManagement"] = $("#labormaintenance67").val();
    userAdvancedInputs["lhppmOthers"] = $("#labormaintenance68").val();
    $("#labormaintenancesave6").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function labormaintenancefunc7() {
    userAdvancedInputs["lhpmoManualWeeding"] = $("#labormaintenance70").val();
    userAdvancedInputs["lhpmoChemicalWeeding"] = $("#labormaintenance71").val();
    userAdvancedInputs["lhpmoOrganicFertilizers"] = $("#labormaintenance72").val();
    userAdvancedInputs["lhpmoChemicalFertilizers"] = $("#labormaintenance73").val();
    userAdvancedInputs["lhpmoFoliarSpraying"] = $("#labormaintenance74").val();
    userAdvancedInputs["lhpmoHedgerowsConstruction"] = $("#labormaintenance75").val();
    userAdvancedInputs["lhpmoShadetreePruning"] = $("#labormaintenance76").val();
    userAdvancedInputs["lhpmoPestControl"] = $("#labormaintenance77").val();
    userAdvancedInputs["lhpmoCoffeeGrowManagement"] = $("#labormaintenance78").val();
    userAdvancedInputs["lhpmoOthers"] = $("#labormaintenance79").val();
    $("#labormaintenancesave7").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}


function labormaintenancefunc8() {
    userAdvancedInputs["lhphoCoffeeCollecDays"] = $("#labormaintenance80").val();
    userAdvancedInputs["lhphoAdditionDays"] = $("#labormaintenance81").val();
    $("#labormaintenancesave8").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function labormaintenancefunc9() {
    userAdvancedInputs["lhppoFermentation"] = $("#labormaintenance90").val();
    userAdvancedInputs["lhppoWashing"] = $("#labormaintenance91").val();
    userAdvancedInputs["lhppoDrying"] = $("#labormaintenance92").val();
    userAdvancedInputs["lhppoScreening"] = $("#labormaintenance93").val();
    userAdvancedInputs["lhppoSelection"] = $("#labormaintenance94").val();
    userAdvancedInputs["lhppoStorage"] = $("#labormaintenance95").val();
    userAdvancedInputs["lhppoCoffeewastewater"] = $("#labormaintenance96").val();
    userAdvancedInputs["lhppoPulpManagement"] = $("#labormaintenance97").val();
    userAdvancedInputs["lhppoOthers"] = $("#labormaintenance98").val();
    $("#labormaintenancesave9").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function addinputsfunc1() {
    userAdvancedInputs["iiFood"] = $("#addinputs10").val();
    userAdvancedInputs["iiAdditionalTransfers"] = $("#addinputs11").val();
    userAdvancedInputs["iiDaysoftraining"] = $("#addinputs12").val();
    $("#addinputssave1").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function addinputsfunc2() {
    userAdvancedInputs["icCreditfromcooperative"] = $("#addinputs20").val();
    userAdvancedInputs["icCreditfromcooperativeTime"] = $("#addinputs21").val();
    userAdvancedInputs["icCreditfromcooperativeInterest"] = $("#addinputs22").val();
    userAdvancedInputs["icCreditfromagent"] = $("#addinputs23").val();
    userAdvancedInputs["icCreditfromagentTime"] = $("#addinputs24").val();
    userAdvancedInputs["icCreditfromagentInterest"] = $("#addinputs25").val();
    $("#addinputssave2").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function transinputsfunc1() {
    userAdvancedInputs["tgSeedPurchase"] = $("#transinputs10").val();
    userAdvancedInputs["tgWoodTransportation"] = $("#transinputs11").val();
    userAdvancedInputs["tgSandTransportation"] = $("#transinputs12").val();
    userAdvancedInputs["tgOthers"] = $("#transinputs13").val();
    $("#transinputssave1").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function transinputsfunc2() {
    userAdvancedInputs["tnSoilTransportation"] = $("#transinputs20").val();
    userAdvancedInputs["tnSacksMaterialShopping"] = $("#transinputs21").val();
    userAdvancedInputs["tnOthers"] = $("#transinputs22").val();
    $("#transinputssave2").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function transinputsfunc3() {
    userAdvancedInputs["tlpWoodTransportation"] = $("#transinputs30").val();
    userAdvancedInputs["tlpCompostTransportation"] = $("#transinputs31").val();
    userAdvancedInputs["tlpPlantTransportation"] = $("#transinputs32").val();
    userAdvancedInputs["tlpOthers"] = $("#transinputs33").val();
    $("#transinputssave3").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function transinputsfunc4() {
    userAdvancedInputs["tOtherEquipment"] = $("#transinputs40").val();
    userAdvancedInputs["tOtherLaborTransportation"] = $("#transinputs41").val();
    userAdvancedInputs["tOtherCoffeeTransportation"] = $("#transinputs42").val();
    userAdvancedInputs["tOtherSupervisingActivities"] = $("#transinputs43").val();
    userAdvancedInputs["tOthers"] = $("#transinputs44").val();
    $("#transinputssave4").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function admininputsfunc1() {
    userAdvancedInputs["accApplicationFee"] = $("#admintinputs10").val();
    userAdvancedInputs["accAnnualMembership"] = $("#admintinputs11").val();
    userAdvancedInputs["accLifeInsurance"] = $("#admintinputs12").val();
    userAdvancedInputs["accfloCertification"] = $("#admintinputs13").val();
    userAdvancedInputs["accOrganicCertification"] = $("#admintinputs14").val();
    $("#admininputssave1").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function admininputsfunc2() {
    userAdvancedInputs["aclLandValue"] = $("#admintinputs20").val();
    userAdvancedInputs["aclPropertyTax"] = $("#admintinputs21").val();
    $("#admininputssave2").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function admininputsfunc3() {
    userAdvancedInputs["acuSuperviseInvest"] = $("#admintinputs30").val();
    userAdvancedInputs["acuAdministInvest"] = $("#admintinputs31").val();
    userAdvancedInputs["acuTrainingInvest"] = $("#admintinputs32").val();
    userAdvancedInputs["acuExtraOrdInvest"] = $("#admintinputs33").val();
    $("#admininputssave3").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function equipmentinputsfunc1() {
    userAdvancedInputs["egeManualSprayer"] = $("#equipmentinputs10").val();
    userAdvancedInputs["egeLifespam1"] = $("#equipmentinputs11").val();
    userAdvancedInputs["egeMachetes"] = $("#equipmentinputs12").val();
    userAdvancedInputs["egeLifespam2"] = $("#equipmentinputs13").val();
    userAdvancedInputs["egeShovel"] = $("#equipmentinputs14").val();
    userAdvancedInputs["egeLifespam3"] = $("#equipmentinputs15").val();
    userAdvancedInputs["egeHoe"] = $("#equipmentinputs16").val();
    userAdvancedInputs["egeLifespam4"] = $("#equipmentinputs17").val();
    userAdvancedInputs["egeWheelBarrow"] = $("#equipmentinputs18").val();
    userAdvancedInputs["egeLifespam5"] = $("#equipmentinputs19").val();
    userAdvancedInputs["egeLime"] = $("#equipmentinputs110").val();
    userAdvancedInputs["egeLifespam6"] = $("#equipmentinputs111").val();
    userAdvancedInputs["egeAuger"] = $("#equipmentinputs112").val();
    userAdvancedInputs["egeLifespam7"] = $("#equipmentinputs113").val();
    userAdvancedInputs["egeMetalBar"] = $("#equipmentinputs114").val();
    userAdvancedInputs["egeLifespam8"] = $("#equipmentinputs115").val();
    userAdvancedInputs["egeHose"] = $("#equipmentinputs116").val();
    userAdvancedInputs["egeLifespam9"] = $("#equipmentinputs117").val();
    userAdvancedInputs["egeSprinklers"] = $("#equipmentinputs118").val();
    userAdvancedInputs["egeLifespam10"] = $("#equipmentinputs119").val();
    userAdvancedInputs["egeChainSaw"] = $("#equipmentinputs120").val();
    userAdvancedInputs["egeLifespam11"] = $("#equipmentinputs121").val();
    userAdvancedInputs["egeHandSaw"] = $("#equipmentinputs122").val();
    userAdvancedInputs["egeLifespam12"] = $("#equipmentinputs123").val();
    userAdvancedInputs["egeMotorPump"] = $("#equipmentinputs124").val();
    userAdvancedInputs["egeLifespam13"] = $("#equipmentinputs125").val();
    userAdvancedInputs["egePrunningScissors"] = $("#equipmentinputs126").val();
    userAdvancedInputs["egeLifespam14"] = $("#equipmentinputs127").val();
    userAdvancedInputs["egeAxe"] = $("#equipmentinputs128").val();
    userAdvancedInputs["egeLifespam15"] = $("#equipmentinputs129").val();
    $("#equipmentinputssave1").attr('style', 'background-color:#00838F; float:right; border-color:bisque');

}

function equipmentinputsfunc2() {
    userAdvancedInputs["eehScale"] = $("#equipmentinputs20").val();
    userAdvancedInputs["eehLifespam1"] = $("#equipmentinputs21").val();
    userAdvancedInputs["eehVehicle"] = $("#equipmentinputs22").val();
    userAdvancedInputs["eehLifespam2"] = $("#equipmentinputs23").val();
    userAdvancedInputs["eehWorkAnimal"] = $("#equipmentinputs24").val();
    userAdvancedInputs["eehLifespam3"] = $("#equipmentinputs25").val();
    userAdvancedInputs["eehMotorcycle"] = $("#equipmentinputs26").val();
    userAdvancedInputs["eehLifespam4"] = $("#equipmentinputs27").val();
    userAdvancedInputs["eehBags"] = $("#equipmentinputs28").val();
    userAdvancedInputs["eehLifespam5"] = $("#equipmentinputs29").val();
    userAdvancedInputs["eehSack"] = $("#equipmentinputs210").val();
    userAdvancedInputs["eehLifespam6"] = $("#equipmentinputs211").val();
    userAdvancedInputs["eehStraw"] = $("#equipmentinputs212").val();
    userAdvancedInputs["eehLifespam7"] = $("#equipmentinputs213").val();
    userAdvancedInputs["eehBaskets"] = $("#equipmentinputs214").val();
    userAdvancedInputs["eehLifespam8"] = $("#equipmentinputs215").val();
    userAdvancedInputs["eehBoxes"] = $("#equipmentinputs216").val();
    userAdvancedInputs["eehLifespam9"] = $("#equipmentinputs217").val();
    userAdvancedInputs["eehOthers"] = $("#equipmentinputs218").val();
    userAdvancedInputs["eehLifespam10"] = $("#equipmentinputs219").val();
    $("#equipmentinputssave2").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function equipmentinputsfunc3() {
    userAdvancedInputs["eepPulperMachine"] = $("#equipmentinputs30").val();
    userAdvancedInputs["eepLifespam1"] = $("#equipmentinputs31").val();
    userAdvancedInputs["eepTolca"] = $("#equipmentinputs32").val();
    userAdvancedInputs["eepLifespam2"] = $("#equipmentinputs33").val();
    userAdvancedInputs["eepEngine"] = $("#equipmentinputs34").val();
    userAdvancedInputs["eepLifespam3"] = $("#equipmentinputs35").val();
    userAdvancedInputs["eepTanks"] = $("#equipmentinputs36").val();
    userAdvancedInputs["eepLifespam4"] = $("#equipmentinputs37").val();
    userAdvancedInputs["eepWaterChannel"] = $("#equipmentinputs38").val();
    userAdvancedInputs["eepLifespam5"] = $("#equipmentinputs39").val();
    userAdvancedInputs["eepPVCPipes"] = $("#equipmentinputs310").val();
    userAdvancedInputs["eepLifespam6"] = $("#equipmentinputs311").val();
    userAdvancedInputs["eepFilteringSystem"] = $("#equipmentinputs312").val();
    userAdvancedInputs["eepLifespam7"] = $("#equipmentinputs313").val();
    userAdvancedInputs["eepScreeningMachine"] = $("#equipmentinputs314").val();
    userAdvancedInputs["eepLifespam8"] = $("#equipmentinputs315").val();
    userAdvancedInputs["eepDesmucilaginador"] = $("#equipmentinputs316").val();
    userAdvancedInputs["eepLifespam9"] = $("#equipmentinputs317").val();
    userAdvancedInputs["eepMotorPump"] = $("#equipmentinputs318").val();
    userAdvancedInputs["eepLifespam10"] = $("#equipmentinputs319").val();
    userAdvancedInputs["eepOthersWetInput"] = $("#equipmentinputs320").val();
    userAdvancedInputs["eepLifespam11"] = $("#equipmentinputs321").val();
    userAdvancedInputs["eepConcrete"] = $("#equipmentinputs322").val();
    userAdvancedInputs["eepLifespam12"] = $("#equipmentinputs323").val();
    userAdvancedInputs["eepPlastic"] = $("#equipmentinputs324").val();
    userAdvancedInputs["eepLifespam13"] = $("#equipmentinputs325").val();
    userAdvancedInputs["eepRake"] = $("#equipmentinputs326").val();
    userAdvancedInputs["eepLifespam14"] = $("#equipmentinputs327").val();
    userAdvancedInputs["eepBroom"] = $("#equipmentinputs328").val();
    userAdvancedInputs["eepLifespam15"] = $("#equipmentinputs329").val();
    userAdvancedInputs["eepStorageRoom"] = $("#equipmentinputs330").val();
    userAdvancedInputs["eepLifespam16"] = $("#equipmentinputs331").val();
    userAdvancedInputs["eepOthersDryInput"] = $("#equipmentinputs332").val();
    userAdvancedInputs["eepLifespam17"] = $("#equipmentinputs333").val();
    $("#equipmentinputssave3").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function costinputsfunc1() {

    userAdvancedInputs["costGerminator"] = $("#costinputs10").val();
    userAdvancedInputs["costGerminatorSeeds"] = $("#costinputs11").val();
    userAdvancedInputs["costGerminatorSeedbed"] = $("#costinputs12").val();
    userAdvancedInputs["costGerminatorSandSubstrate"] = $("#costinputs13").val();
    userAdvancedInputs["costGerminatorCalciumSulfide"] = $("#costinputs14").val();
    userAdvancedInputs["costGerminatorLime"] = $("#costinputs15").val();
    userAdvancedInputs["costGerminatorPlastic"] = $("#costinputs16").val();
    userAdvancedInputs["costGerminatorOthers"] = $("#costinputs17").val();
    $("#costinputssave1").attr('style', 'background-color:#00838F; float:right; border-color:bisque');


}

function costinputsfunc2() {
    userAdvancedInputs["costNurseryFertilizer"] = $("#costinputs20").val();
    userAdvancedInputs["costNurseryPlasticBags"] = $("#costinputs21").val();
    userAdvancedInputs["costNurseryNetting"] = $("#costinputs22").val();
    userAdvancedInputs["costNurseryStuds"] = $("#costinputs23").val();
    userAdvancedInputs["costNurseryWire"] = $("#costinputs24").val();
    userAdvancedInputs["costNurseryCiclonics"] = $("#costinputs25").val();
    userAdvancedInputs["costNurseryStaples"] = $("#costinputs26").val();
    userAdvancedInputs["costNurserySoil"] = $("#costinputs27").val();
    userAdvancedInputs["costNurseryBioFert"] = $("#costinputs28").val();
    userAdvancedInputs["costNurseryAgroChemicals"] = $("#costinputs29").val();
    userAdvancedInputs["costNurseryFungicide"] = $("#costinputs210").val();
    userAdvancedInputs["costNurseryPhosphoricRock"] = $("#costinputs211").val();
    userAdvancedInputs["costNurseryOthers"] = $("#costinputs212").val();
    $("#costinputssave2").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function costinputsfunc3() {
    userAdvancedInputs["costFLPPOrganicFert"] = $("#costinputs30").val();
    userAdvancedInputs["costFLPPChemicalFert"] = $("#costinputs31").val();
    $("#costinputssave3").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function costinputsfunc4() {
    userAdvancedInputs["costFVGOrganicFert"] = $("#costinputs40").val();
    userAdvancedInputs["costFVGChemicalFert"] = $("#costinputs41").val();
    $("#costinputssave4").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

function costinputsfunc5() {
    userAdvancedInputs["costFMOtherFert"] = $("#costinputs50").val();
    userAdvancedInputs["costFMOrganicFoliar"] = $("#costinputs51").val();
    userAdvancedInputs["costFMChemicalFoliar"] = $("#costinputs52").val();
    userAdvancedInputs["costFMGasFuel"] = $("#costinputs53").val();
    userAdvancedInputs["costFMOthers"] = $("#costinputs54").val();
    $("#costinputssave5").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
}

//make user input object
var userAdvancedInputs = {
    "LGerminationSeedCollection": $("#laborgerminatorinputs10").val(),
    "LGerminationSeedSelection": $("#laborgerminatorinputs11").val(),
    "LGerminationNurseryConstruction": $("#laborgerminatorinputs12").val(),
    "LGerminationSeedingSupportIrrigation": $("#laborgerminatorinputs13").val(),
    "LGerminationOthers": $("#laborgerminatorinputs14").val(),
    "LNurseryConstruction": $("#laborgerminatorinputs20").val(),
    "LNurseryDrawnPulled": $("#laborgerminatorinputs21").val(),
    "LNurseryClean": $("#laborgerminatorinputs22").val(),
    "LNurserySoilPreparationFertilizer": $("#laborgerminatorinputs23").val(),
    "LNurseryFilledLockedBags": $("#laborgerminatorinputs24").val(),
    "LNurseryButterflySowing": $("#laborgerminatorinputs25").val(),
    "LNurseryIrrigation": $("#laborgerminatorinputs26").val(),
    "LNurseryFoliarApplication": $("#laborgerminatorinputs27").val(),
    "LNurseryReseeding": $("#laborgerminatorinputs28").val(),
    "LNurseryOthers": $("#laborgerminatorinputs29").val(),
    "LPPFieldCleaning": $("#laborgerminatorinputs30").val(),
    "LPPCuttingTrees": $("#laborgerminatorinputs31").val(),
    "LPPWoodCollection": $("#laborgerminatorinputs32").val(),
    "LPPWoodChopping": $("#laborgerminatorinputs33").val(),
    "LPPCoffeeLayout": $("#laborgerminatorinputs34").val(),
    "LPPHoleDigging": $("#laborgerminatorinputs35").val(),
    "LPPSeedlingTransportation": $("#laborgerminatorinputs36").val(),
    "LPPSeedlingTransplant": $("#laborgerminatorinputs37").val(),
    "LPPShadeAdjustment": $("#laborgerminatorinputs38").val(),
    "LPPCompostMixing": $("#laborgerminatorinputs39").val(),
    "LPPOthers": $("#laborgerminatorinputs310").val(),
    "LPPYWeeding": $("#laborgerminatorinputs40").val(),
    "LPPYOrganic": $("#laborgerminatorinputs41").val(),
    "LPPYChemical": $("#laborgerminatorinputs42").val(),
    "LPPYFoliarSpraying": $("#laborgerminatorinputs43").val(),
    "LPPYOther": $("#laborgerminatorinputs43").val(),
    "LHPMYManualWeeding": $("#labormaintenance10").val(),
    "LHPMYChemicalWeeding": $("#labormaintenance11").val(),
    "LHPMYOrganicFertilizers": $("#labormaintenance12").val(),
    "LHPMYChemicalFertilizers": $("#labormaintenance13").val(),
    "LHPMYFoliarSpraying": $("#labormaintenance14").val(),
    "LHPMYHedgerowsConstruction": $("#labormaintenance15").val(),
    "LHPMYShadetreePruning": $("#labormaintenance16").val(),
    "LHPMYPestControl": $("#labormaintenance17").val(),
    "LHPMYCoffeeGrowManagement": $("#labormaintenance18").val(),
    "LHPMYOthers": $("#labormaintenance19").val(),
    "LHPHYCoffeeCollecDays": $("#labormaintenance20").val(),
    "LHPHYAdditionDays": $("#labormaintenance21").val(),
    "LHPPYFermentation": $("#labormaintenance30").val(),
    "LHPPYWashing": $("#labormaintenance31").val(),
    "LHPPYDrying": $("#labormaintenance32").val(),
    "LHPPYScreening": $("#labormaintenance33").val(),
    "LHPPYSelection": $("#labormaintenance34").val(),
    "LHPPYStorage": $("#labormaintenance35").val(),
    "LHPPYCoffeewastewater": $("#labormaintenance36").val(),
    "LHPPYPulpManagement": $("#labormaintenance37").val(),
    "LHPPYOthers": $("#labormaintenance38").val(),
    "LHPMMManualWeeding": $("#labormaintenance40").val(),
    "LHPMMChemicalWeeding": $("#labormaintenance41").val(),
    "LHPMMOrganicFertilizers": $("#labormaintenance42").val(),
    "LHPMMChemicalFertilizers": $("#labormaintenance43").val(),
    "LHPMMFoliarSpraying": $("#labormaintenance44").val(),
    "LHPMMHedgerowsConstruction": $("#labormaintenance45").val(),
    "LHPMMShadetreePruning": $("#labormaintenance46").val(),
    "LHPMMPestControl": $("#labormaintenance47").val(),
    "LHPMMCoffeeGrowManagement": $("#labormaintenance48").val(),
    "LHPMMOthers": $("#labormaintenance49").val(),
    "LHPHMCoffeeCollecDays": $("#labormaintenance50").val(),
    "LHPHMAdditionDays": $("#labormaintenance51").val(),
    "LHPPMFermentation": $("#labormaintenance60").val(),
    "LHPPMWashing": $("#labormaintenance61").val(),
    "LHPPMDrying": $("#labormaintenance62").val(),
    "LHPPMScreening": $("#labormaintenance63").val(),
    "LHPPMSelection": $("#labormaintenance64").val(),
    "LHPPMStorage": $("#labormaintenance65").val(),
    "LHPPMCoffeewastewater": $("#labormaintenance66").val(),
    "LHPPMPulpManagement": $("#labormaintenance67").val(),
    "LHPPMOthers": $("#labormaintenance68").val(),
    "LHPMOManualWeeding": $("#labormaintenance70").val(),
    "LHPMOChemicalWeeding": $("#labormaintenance71").val(),
    "LHPMOOrganicFertilizers": $("#labormaintenance72").val(),
    "LHPMOChemicalFertilizers": $("#labormaintenance73").val(),
    "LHPMOFoliarSpraying": $("#labormaintenance74").val(),
    "LHPMOHedgerowsConstruction": $("#labormaintenance75").val(),
    "LHPMOShadetreePruning": $("#labormaintenance76").val(),
    "LHPMOPestControl": $("#labormaintenance77").val(),
    "LHPMOCoffeeGrowManagement": $("#labormaintenance78").val(),
    "LHPMOOthers": $("#labormaintenance79").val(),
    "LHPHOCoffeeCollecDays": $("#labormaintenance80").val(),
    "LHPHOAdditionDays": $("#labormaintenance81").val(),
    "LHPPOFermentation": $("#labormaintenance90").val(),
    "LHPPOWashing": $("#labormaintenance91").val(),
    "LHPPODrying": $("#labormaintenance92").val(),
    "LHPPOScreening": $("#labormaintenance93").val(),
    "LHPPOSelection": $("#labormaintenance94").val(),
    "LHPPOStorage": $("#labormaintenance95").val(),
    "LHPPOCoffeewastewater": $("#labormaintenance96").val(),
    "LHPPOPulpManagement": $("#labormaintenance97").val(),
    "LHPPOOthers": $("#labormaintenance98").val(),
    "IIFood": $("#addinputs10").val(),
    "IIAdditionalTransfers": $("#addinputs11").val(),
    "IIDaysoftraining": $("#addinputs12").val(),
    "ICCreditfromcooperative": $("#addinputs20").val(),
    "ICCreditfromcooperativeTime": $("#addinputs21").val(),
    "ICCreditfromcooperativeInterest": $("#addinputs22").val(),
    "ICCreditfromagent": $("#addinputs23").val(),
    "ICCreditfromagentTime": $("#addinputs24").val(),
    "ICCreditfromagentInterest": $("#addinputs25").val(),
    "CostGerminator": $("#costinputs10").val(),
    "CostGerminatorSeeds": $("#costinputs11").val(),
    "CostGerminatorSeedbed": $("#costinputs12").val(),
    "CostGerminatorSandSubstrate": $("#costinputs13").val(),
    "CostGerminatorCalciumSulfide": $("#costinputs14").val(),
    "CostGerminatorLime": $("#costinputs15").val(),
    "CostGerminatorPlastic": $("#costinputs16").val(),
    "CostGerminatorOthers": $("#costinputs17").val(),
    "CostNurseryFertilizer": $("#costinputs20").val(),
    "CostNurseryPlasticBags": $("#costinputs21").val(),
    "CostNurseryNetting": $("#costinputs22").val(),
    "CostNurseryStuds": $("#costinputs23").val(),
    "CostNurseryWire": $("#costinputs24").val(),
    "CostNurseryCiclonics": $("#costinputs25").val(),
    "CostNurseryStaples": $("#costinputs26").val(),
    "CostNurserySoil": $("#costinputs27").val(),
    "CostNurseryBioFert": $("#costinputs28").val(),
    "CostNurseryAgroChemicals": $("#costinputs29").val(),
    "CostNurseryFungicide": $("#costinputs210").val(),
    "CostNurseryPhosphoricRock": $("#costinputs211").val(),
    "CostNurseryOthers": $("#costinputs212").val(),
    "CostFLPPOrganicFert": $("#costinputs30").val(),
    "CostFLPPChemicalFert": $("#costinputs31").val(),
    "CostFVGOrganicFert": $("#costinputs40").val(),
    "CostFVGChemicalFert": $("#costinputs41").val(),
    "CostFMOtherFert": $("#costinputs50").val(),
    "CostFMOrganicFoliar": $("#costinputs51").val(),
    "CostFMChemicalFoliar": $("#costinputs52").val(),
    "CostFMGasFuel": $("#costinputs53").val(),
    "CostFMOthers": $("#costinputs54").val(),
    "EGEManualSprayer": $("#equipmentinputs10").val(),
    "EGELifespam1": $("#equipmentinputs11").val(),
    "EGEMachetes": $("#equipmentinputs12").val(),
    "EGELifespam2": $("#equipmentinputs13").val(),
    "EGEShovel": $("#equipmentinputs14").val(),
    "EGELifespam3": $("#equipmentinputs15").val(),
    "EGEHoe": $("#equipmentinputs16").val(),
    "EGELifespam4": $("#equipmentinputs17").val(),
    "EGEWheelBarrow": $("#equipmentinputs18").val(),
    "EGELifespam5": $("#equipmentinputs19").val(),
    "EGELime": $("#equipmentinputs110").val(),
    "EGELifespam6": $("#equipmentinputs111").val(),
    "EGEAuger": $("#equipmentinputs112").val(),
    "EGELifespam7": $("#equipmentinputs113").val(),
    "EGEMetalBar": $("#equipmentinputs114").val(),
    "EGELifespam8": $("#equipmentinputs115").val(),
    "EGEHose": $("#equipmentinputs116").val(),
    "EGELifespam9": $("#equipmentinputs117").val(),
    "EGESprinklers": $("#equipmentinputs118").val(),
    "EGELifespam10": $("#equipmentinputs119").val(),
    "EGEChainSaw": $("#equipmentinputs120").val(),
    "EGELifespam11": $("#equipmentinputs121").val(),
    "EGEHandSaw": $("#equipmentinputs122").val(),
    "EGELifespam12": $("#equipmentinputs123").val(),
    "EGEMotorPump": $("#equipmentinputs124").val(),
    "EGELifespam13": $("#equipmentinputs125").val(),
    "EGEPrunningScissors": $("#equipmentinputs126").val(),
    "EGELifespam14": $("#equipmentinputs127").val(),
    "EGEAxe": $("#equipmentinputs128").val(),
    "EGELifespam15": $("#equipmentinputs129").val(),
    "EEHScale": $("#equipmentinputs20").val(),
    "EEHLifespam1": $("#equipmentinputs21").val(),
    "EEHVehicle": $("#equipmentinputs22").val(),
    "EEHLifespam2": $("#equipmentinputs23").val(),
    "EEHWorkAnimal": $("#equipmentinputs24").val(),
    "EEHLifespam3": $("#equipmentinputs25").val(),
    "EEHMotorcycle": $("#equipmentinputs26").val(),
    "EEHLifespam4": $("#equipmentinputs27").val(),
    "EEHBags": $("#equipmentinputs28").val(),
    "EEHLifespam5": $("#equipmentinputs29").val(),
    "EEHSack": $("#equipmentinputs210").val(),
    "EEHLifespam6": $("#equipmentinputs211").val(),
    "EEHStraw": $("#equipmentinputs212").val(),
    "EEHLifespam7": $("#equipmentinputs213").val(),
    "EEHBaskets": $("#equipmentinputs214").val(),
    "EEHLifespam8": $("#equipmentinputs215").val(),
    "EEHBoxes": $("#equipmentinputs216").val(),
    "EEHLifespam9": $("#equipmentinputs217").val(),
    "EEHOthers": $("#equipmentinputs218").val(),
    "EEHLifespam10": $("#equipmentinputs219").val(),
    "EEPPulperMachine": $("#equipmentinputs30").val(),
    "EEPLifespam1": $("#equipmentinputs31").val(),
    "EEPTolca": $("#equipmentinputs32").val(),
    "EEPLifespam2": $("#equipmentinputs33").val(),
    "EEPEngine": $("#equipmentinputs34").val(),
    "EEPLifespam3": $("#equipmentinputs35").val(),
    "EEPTanks": $("#equipmentinputs36").val(),
    "EEPLifespam4": $("#equipmentinputs37").val(),
    "EEPWaterChannel": $("#equipmentinputs38").val(),
    "EEPLifespam5": $("#equipmentinputs39").val(),
    "EEPPVCPipes": $("#equipmentinputs310").val(),
    "EEPLifespam6": $("#equipmentinputs311").val(),
    "EEPFilteringSystem": $("#equipmentinputs312").val(),
    "EEPLifespam7": $("#equipmentinputs313").val(),
    "EEPScreeningMachine": $("#equipmentinputs314").val(),
    "EEPLifespam8": $("#equipmentinputs315").val(),
    "EEPDesmucilaginador": $("#equipmentinputs316").val(),
    "EEPLifespam9": $("#equipmentinputs317").val(),
    "EEPMotorPump": $("#equipmentinputs318").val(),
    "EEPLifespam10": $("#equipmentinputs319").val(),
    "EEPOthersWetInput": $("#equipmentinputs320").val(),
    "EEPLifespam11": $("#equipmentinputs321").val(),
    "EEPConcrete": $("#equipmentinputs322").val(),
    "EEPLifespam12": $("#equipmentinputs323").val(),
    "EEPPlastic": $("#equipmentinputs324").val(),
    "EEPLifespam13": $("#equipmentinputs325").val(),
    "EEPRake": $("#equipmentinputs326").val(),
    "EEPLifespam14": $("#equipmentinputs327").val(),
    "EEPBroom": $("#equipmentinputs328").val(),
    "EEPLifespam15": $("#equipmentinputs329").val(),
    "EEPStorageRoom": $("#equipmentinputs330").val(),
    "EEPLifespam16": $("#equipmentinputs331").val(),
    "EEPOthersDryInput": $("#equipmentinputs332").val(),
    "EEPLifespam17": $("#equipmentinputs333").val(),
    "ACCApplicationFee": $("#admintinputs10").val(),
    "ACCAnnualMembership": $("#admintinputs11").val(),
    "ACCLifeInsurance": $("#admintinputs12").val(),
    "ACCFLOCertification": $("#admintinputs13").val(),
    "ACCOrganicCertification": $("#admintinputs14").val(),
    "ACLLandValue": $("#admintinputs20").val(),
    "ACLPropertyTax": $("#admintinputs21").val(),
    "ACUSuperviseInvest": $("#admintinputs30").val(),
    "ACUAdministInvest": $("#admintinputs31").val(),
    "ACUTrainingInvest": $("#admintinputs32").val(),
    "ACUExtraOrdInvest": $("#admintinputs33").val(),
    "TGSeedPurchase": $("#transinputs10").val(),
    "TGWoodTransportation": $("#transinputs11").val(),
    "TGSandTransportation": $("#transinputs12").val(),
    "TGOthers": $("#transinputs13").val(),
    "TNSoilTransportation": $("#transinputs20").val(),
    "TNSacksMaterialShopping": $("#transinputs21").val(),
    "TNOthers": $("#transinputs22").val(),
    "TLPWoodTransportation": $("#transinputs30").val(),
    "TLPCompostTransportation": $("#transinputs31").val(),
    "TLPPlantTransportation": $("#transinputs32").val(),
    "TLPOthers": $("#transinputs33").val(),
    "TOtherEquipment": $("#transinputs40").val(),
    "TOtherLaborTransportation": $("#transinputs41").val(),
    "TOtherCoffeeTransportation": $("#transinputs42").val(),
    "TOtherSupervisingActivities": $("#transinputs43").val(),
    "TOthers": $("#transinputs44").val(),
}
//userAdvancedInputs = JSON.stringify(userAdvancedInputs);
var laborgerminator1 = [];
var laborgerminator2 = [];
var laborgerminator3 = [];
var laborgerminator4 = [];
var labormaintenance1 = [];
var labormaintenance2 = [];
var labormaintenance3 = [];
var labormaintenance4 = [];
var labormaintenance5 = [];
var labormaintenance6 = [];
var labormaintenance7 = [];
var labormaintenance8 = [];
var labormaintenance9 = [];
var addinputs1 = [];
var addinputs2 = [];
var costinputs1 = [];
var costinputs2 = [];
var costinputs3 = [];
var costinputs4 = [];
var costinputs5 = [];
var equipmentinputs1 = [];
var equipmentinputs2 = [];
var equipmentinputs3 = [];
var admininputs1 = [];
var admininputs2 = [];
var admininputs3 = [];
var transinputs1 = [];
var transinputs2 = [];
var transinputs3 = [];
var transinputs4 = [];

function getAdvanced() {
    debugger;
    var promise = $.ajax({
        type: "GET",
        url: apiURL + "TechnicianHomeAPI/getinputvalues",
        //data: request,
        contentType: "application/json; charset=utf-8",
        success: function (result, status) {
            console.log(result);
            userAdvancedInputs = result;
            laborgerminator1[0] = result["lGerminationSeedCollection"];
            laborgerminator1[1] = result["lGerminationSeedSelection"];
            laborgerminator1[2] = result["lGerminationNurseryConstruction"];
            laborgerminator1[3] = result["lGerminationSeedingSupportIrrigation"];
            laborgerminator1[4] = result["lGerminationOthers"];
            laborgerminator2[0] = result["lNurseryConstruction"];
            laborgerminator2[1] = result["lNurseryDrawnPulled"];
            laborgerminator2[2] = result["lNurseryClean"];
            laborgerminator2[3] = result["lNurserySoilPreparationFertilizer"];
            laborgerminator2[4] = result["lNurseryFilledLockedBags"];
            laborgerminator2[5] = result["lNurseryButterflySowing"];
            laborgerminator2[6] = result["lNurseryIrrigation"];
            laborgerminator2[7] = result["lNurseryFoliarApplication"];
            laborgerminator2[8] = result["lNurseryReseeding"];
            laborgerminator2[9] = result["lNurseryOthers"];
            laborgerminator3[0] = result["lppFieldCleaning"];
            laborgerminator3[1] = result["lppCuttingTrees"];
            laborgerminator3[2] = result["lppWoodCollection"];
            laborgerminator3[3] = result["lppWoodChopping"];
            laborgerminator3[4] = result["lppCoffeeLayout"];
            laborgerminator3[5] = result["lppHoleDigging"];
            laborgerminator3[6] = result["lppSeedlingTransportation"];
            laborgerminator3[7] = result["lppSeedlingTransplant"];
            laborgerminator3[8] = result["lppShadeAdjustment"];
            laborgerminator3[9] = result["lppCompostMixing"];
            laborgerminator3[10] = result["lppOthers"];
            laborgerminator4[0] = result["lppyWeeding"];
            laborgerminator4[1] = result["lppyOrganic"];
            laborgerminator4[2] = result["lppyChemical"];
            laborgerminator4[3] = result["lppyFoliarSpraying"];
            laborgerminator4[4] = result["lppyOther"];
            labormaintenance1[0] = result["lhpmyManualWeeding"];
            labormaintenance1[1] = result["lhpmyChemicalWeeding"];
            labormaintenance1[2] = result["lhpmyOrganicFertilizers"];
            labormaintenance1[3] = result["lhpmyChemicalFertilizers"];
            labormaintenance1[4] = result["lhpmyFoliarSpraying"];
            labormaintenance1[5] = result["lhpmyHedgerowsConstruction"];
            labormaintenance1[6] = result["lhpmyShadetreePruning"];
            labormaintenance1[7] = result["lhpmyPestControl"];
            labormaintenance1[8] = result["lhpmyCoffeeGrowManagement"];
            labormaintenance1[9] = result["lhpmyOthers"];
            labormaintenance2[0] = result["lhphyCoffeeCollecDays"];
            labormaintenance2[1] = result["lhphyAdditionDays"];
            labormaintenance3[0] = result["lhppyFermentation"];
            labormaintenance3[1] = result["lhppyWashing"];
            labormaintenance3[2] = result["lhppyDrying"];
            labormaintenance3[3] = result["lhppyScreening"];
            labormaintenance3[4] = result["lhppySelection"];
            labormaintenance3[5] = result["lhppyStorage"];
            labormaintenance3[6] = result["lhppyCoffeewastewater"];
            labormaintenance3[7] = result["lhppyPulpManagement"];
            labormaintenance3[8] = result["lhppyOthers"];
            labormaintenance4[0] = result["lhpmmManualWeeding"];
            labormaintenance4[1] = result["lhpmmChemicalWeeding"];
            labormaintenance4[2] = result["lhpmmOrganicFertilizers"];
            labormaintenance4[3] = result["lhpmmChemicalFertilizers"];
            labormaintenance4[4] = result["lhpmmFoliarSpraying"];
            labormaintenance4[5] = result["lhpmmHedgerowsConstruction"];
            labormaintenance4[6] = result["lhpmmShadetreePruning"];
            labormaintenance4[7] = result["lhpmmPestControl"];
            labormaintenance4[8] = result["lhpmmCoffeeGrowManagement"];
            labormaintenance4[9] = result["lhpmmOthers"];
            labormaintenance5[0] = result["lhphmCoffeeCollecDays"];
            labormaintenance5[1] = result["lhphmAdditionDays"];
            labormaintenance6[0] = result["lhppmFermentation"];
            labormaintenance6[1] = result["lhppmWashing"];
            labormaintenance6[2] = result["lhppmDrying"];
            labormaintenance6[3] = result["lhppmScreening"];
            labormaintenance6[4] = result["lhppmSelection"];
            labormaintenance6[5] = result["lhppmStorage"];
            labormaintenance6[6] = result["lhppmCoffeewastewater"];
            labormaintenance6[7] = result["lhppmPulpManagement"];
            labormaintenance6[8] = result["lhppmOthers"];
            labormaintenance7[0] = result["lhpmoManualWeeding"];
            labormaintenance7[1] = result["lhpmoChemicalWeeding"];
            labormaintenance7[2] = result["lhpmoOrganicFertilizers"];
            labormaintenance7[3] = result["lhpmoChemicalFertilizers"];
            labormaintenance7[4] = result["lhpmoFoliarSpraying"];
            labormaintenance7[5] = result["lhpmoHedgerowsConstruction"];
            labormaintenance7[6] = result["lhpmoShadetreePruning"];
            labormaintenance7[7] = result["lhpmoPestControl"];
            labormaintenance7[8] = result["lhpmoCoffeeGrowManagement"];
            labormaintenance7[9] = result["lhpmoOthers"];
            labormaintenance8[0] = result["lhphoCoffeeCollecDays"];
            labormaintenance8[1] = result["lhphoAdditionDays"];
            labormaintenance9[0] = result["lhppoFermentation"];
            labormaintenance9[1] = result["lhppoWashing"];
            labormaintenance9[2] = result["lhppoDrying"];
            labormaintenance9[3] = result["lhppoScreening"];
            labormaintenance9[4] = result["lhppoSelection"];
            labormaintenance9[5] = result["lhppoStorage"];
            labormaintenance9[6] = result["lhppoCoffeewastewater"];
            labormaintenance9[7] = result["lhppoPulpManagement"];
            labormaintenance9[8] = result["lhppoOthers"];
            addinputs1[0] = result["iiFood"];
            addinputs1[1] = result["iiAdditionalTransfers"];
            addinputs1[2] = result["iiDaysoftraining"];
            addinputs2[0] = result["icCreditfromcooperative"];
            addinputs2[1] = result["icCreditfromcooperativeTime"];
            addinputs2[2] = result["icCreditfromcooperativeInterest"];
            addinputs2[3] = result["icCreditfromagent"];
            addinputs2[4] = result["icCreditfromagentTime"];
            addinputs2[5] = result["icCreditfromagentInterest"];
            costinputs1[0] = result["costGerminator"];
            costinputs1[1] = result["costGerminatorSeeds"];
            costinputs1[2] = result["costGerminatorSeedbed"];
            costinputs1[3] = result["costGerminatorSandSubstrate"];
            costinputs1[4] = result["costGerminatorCalciumSulfide"];
            costinputs1[5] = result["costGerminatorLime"];
            costinputs1[6] = result["costGerminatorPlastic"];
            costinputs1[7] = result["costGerminatorOthers"];
            costinputs2[0] = result["costNurseryFertilizer"];
            costinputs2[1] = result["costNurseryPlasticBags"];
            costinputs2[2] = result["costNurseryNetting"];
            costinputs2[3] = result["costNurseryStuds"];
            costinputs2[4] = result["costNurseryWire"];
            costinputs2[5] = result["costNurseryCiclonics"];
            costinputs2[6] = result["costNurseryStaples"];
            costinputs2[7] = result["costNurserySoil"];
            costinputs2[8] = result["costNurseryBioFert"];
            costinputs2[9] = result["costNurseryAgroChemicals"];
            costinputs2[10] = result["costNurseryFungicide"];
            costinputs2[11] = result["costNurseryPhosphoricRock"];
            costinputs2[12] = result["costNurseryOthers"];
            costinputs3[0] = result["costFLPPOrganicFert"];
            costinputs3[1] = result["costFLPPChemicalFert"];
            costinputs4[0] = result["costFVGOrganicFert"];
            costinputs4[1] = result["costFVGChemicalFert"];
            costinputs5[0] = result["costFMOtherFert"];
            costinputs5[1] = result["costFMOrganicFoliar"];
            costinputs5[2] = result["costFMChemicalFoliar"];
            costinputs5[3] = result["costFMGasFuel"];
            costinputs5[4] = result["costFMOthers"];
            equipmentinputs1[0] = result["egeManualSprayer"];
            equipmentinputs1[1] = result["egeLifespam1"];
            equipmentinputs1[2] = result["egeMachetes"];
            equipmentinputs1[3] = result["egeLifespam2"];
            equipmentinputs1[4] = result["egeShovel"];
            equipmentinputs1[5] = result["egeLifespam3"];
            equipmentinputs1[6] = result["egeHoe"];
            equipmentinputs1[7] = result["egeLifespam4"];
            equipmentinputs1[8] = result["egeWheelBarrow"];
            equipmentinputs1[9] = result["egeLifespam5"];
            equipmentinputs1[10] = result["egeLime"];
            equipmentinputs1[11] = result["egeLifespam6"];
            equipmentinputs1[12] = result["egeAuger"];
            equipmentinputs1[13] = result["egeLifespam7"];
            equipmentinputs1[14] = result["egeMetalBar"];
            equipmentinputs1[15] = result["egeLifespam8"];
            equipmentinputs1[16] = result["egeHose"];
            equipmentinputs1[17] = result["egeLifespam9"];
            equipmentinputs1[18] = result["egeSprinklers"];
            equipmentinputs1[19] = result["egeLifespam10"];
            equipmentinputs1[20] = result["egeChainSaw"];
            equipmentinputs1[21] = result["egeLifespam11"];
            equipmentinputs1[22] = result["egeHandSaw"];
            equipmentinputs1[23] = result["egeLifespam12"];
            equipmentinputs1[24] = result["egeMotorPump"];
            equipmentinputs1[25] = result["egeLifespam13"];
            equipmentinputs1[26] = result["egePrunningScissors"];
            equipmentinputs1[27] = result["egeLifespam14"];
            equipmentinputs1[28] = result["egeAxe"];
            equipmentinputs1[29] = result["egeLifespam15"];
            equipmentinputs2[0] = result["eehScale"];
            equipmentinputs2[1] = result["eehLifespam1"];
            equipmentinputs2[2] = result["eehVehicle"];
            equipmentinputs2[3] = result["eehLifespam2"];
            equipmentinputs2[4] = result["eehWorkAnimal"];
            equipmentinputs2[5] = result["eehLifespam3"];
            equipmentinputs2[6] = result["eehMotorcycle"];
            equipmentinputs2[7] = result["eehLifespam4"];
            equipmentinputs2[8] = result["eehBags"];
            equipmentinputs2[9] = result["eehLifespam5"];
            equipmentinputs2[10] = result["eehSack"];
            equipmentinputs2[11] = result["eehLifespam6"];
            equipmentinputs2[12] = result["eehStraw"];
            equipmentinputs2[13] = result["eehLifespam7"];
            equipmentinputs2[14] = result["eehBaskets"];
            equipmentinputs2[15] = result["eehLifespam8"];
            equipmentinputs2[16] = result["eehBoxes"];
            equipmentinputs2[17] = result["eehLifespam9"];
            equipmentinputs2[18] = result["eehOthers"];
            equipmentinputs2[19] = result["eehLifespam10"];
            equipmentinputs3[0] = result["eepPulperMachine"];
            equipmentinputs3[1] = result["eepLifespam1"];
            equipmentinputs3[2] = result["eepTolca"];
            equipmentinputs3[3] = result["eepLifespam2"];
            equipmentinputs3[4] = result["eepEngine"];
            equipmentinputs3[5] = result["eepLifespam3"];
            equipmentinputs3[6] = result["eepTanks"];
            equipmentinputs3[7] = result["eepLifespam4"];
            equipmentinputs3[8] = result["eepWaterChannel"];
            equipmentinputs3[9] = result["eepLifespam5"];
            equipmentinputs3[10] = result["eeppvcPipes"];
            equipmentinputs3[11] = result["eepLifespam6"];
            equipmentinputs3[12] = result["eepFilteringSystem"];
            equipmentinputs3[13] = result["eepLifespam7"];
            equipmentinputs3[14] = result["eepScreeningMachine"];
            equipmentinputs3[15] = result["eepLifespam8"];
            equipmentinputs3[16] = result["eepDesmucilaginador"];
            equipmentinputs3[17] = result["eepLifespam9"];
            equipmentinputs3[18] = result["eepMotorPump"];
            equipmentinputs3[19] = result["eepLifespam10"];
            equipmentinputs3[20] = result["eepOthersWetInput"];
            equipmentinputs3[21] = result["eepLifespam11"];
            equipmentinputs3[22] = result["eepConcrete"];
            equipmentinputs3[23] = result["eepLifespam12"];
            equipmentinputs3[24] = result["eepPlastic"];
            equipmentinputs3[25] = result["eepLifespam13"];
            equipmentinputs3[26] = result["eepRake"];
            equipmentinputs3[27] = result["eepLifespam14"];
            equipmentinputs3[28] = result["eepBroom"];
            equipmentinputs3[29] = result["eepLifespam15"];
            equipmentinputs3[30] = result["eepStorageRoom"];
            equipmentinputs3[31] = result["eepLifespam16"];
            equipmentinputs3[32] = result["eepOthersDryInput"];
            equipmentinputs3[33] = result["eepLifespam17"];
            admininputs1[0] = result["accApplicationFee"];
            admininputs1[1] = result["accAnnualMembership"];
            admininputs1[2] = result["accLifeInsurance"];
            admininputs1[3] = result["accfloCertification"];
            admininputs1[4] = result["accOrganicCertification"];
            admininputs2[0] = result["aclLandValue"];
            admininputs2[1] = result["aclPropertyTax"];
            admininputs3[0] = result["acuSuperviseInvest"];
            admininputs3[1] = result["acuAdministInvest"];
            admininputs3[2] = result["acuTrainingInvest"];
            admininputs3[3] = result["acuExtraOrdInvest"];
            transinputs1[0] = result["tgSeedPurchase"];
            transinputs1[1] = result["tgWoodTransportation"];
            transinputs1[2] = result["tgSandTransportation"];
            transinputs1[3] = result["tgOthers"];
            transinputs2[0] = result["tnSoilTransportation"];
            transinputs2[1] = result["tnSacksMaterialShopping"];
            transinputs2[2] = result["tnOthers"];
            transinputs3[0] = result["tlpWoodTransportation"];
            transinputs3[1] = result["tlpCompostTransportation"];
            transinputs3[2] = result["tlpPlantTransportation"];
            transinputs3[3] = result["tlpOthers"];
            transinputs4[0] = result["tOtherEquipment"];
            transinputs4[1] = result["tOtherLaborTransportation"];
            transinputs4[2] = result["tOtherCoffeeTransportation"];
            transinputs4[3] = result["tOtherSupervisingActivities"];
            transinputs4[4] = result["tOthers"];
            console.log(laborgerminator1);
        },
        error: function (res, status) {
        }
    });

}
function saveAdvanced() {

    debugger;
    console.log(userAdvancedInputs);
    var request = JSON.stringify(userAdvancedInputs);
    console.log(request);
    var promise = $.ajax({
        type: "POST",
        url: apiURL + "TechnicianHomeAPI/saveinputvalues",
        data: request,
        contentType: "application/json; charset=utf-8",
        success: function (result, status) {

            $("#laborgerminatorsave1").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#laborgerminatorsave2").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#laborgerminatorsave3").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#laborgerminatorsave4").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#labormaintenancesave1").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#labormaintenancesave2").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#labormaintenancesave3").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#labormaintenancesave4").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#labormaintenancesave5").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#labormaintenancesave6").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#labormaintenancesave7").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#labormaintenancesave8").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#labormaintenancesave9").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#costinputssave1").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#costinputssave2").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#costinputssave3").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#costinputssave4").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#costinputssave5").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#transinputssave1").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#transinputssave2").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#transinputssave3").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#transinputssave4").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#equipmentinputssave1").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#equipmentinputssave2").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#equipmentinputssave3").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#admininputssave1").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#admininputssave2").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#admininputssave3").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#addinputssave1").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#addinputssave2").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
        },
        error: function (res, status) {
        }
    });
}