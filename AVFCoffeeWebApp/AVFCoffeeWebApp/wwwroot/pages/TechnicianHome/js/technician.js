

$(document).on('ready', function () {
    debugger;
    Metrics(apiURL);
});

function change() {
    $("#Metrics").removeClass("mdl-navigation__link")
    $("#Metrics").addClass("mdl-navigation__link is-active");
}

function press() {
    $("#coffeeparchment").val("Choose...");
    $("#length").val("Choose...");
    $("#farmarea").val("Choose...");
    $("#weight").val("Choose...");
    $("#capacity").val("Choose...");
    $("#currency").val("Choose...");
}

function pressadvanced() {
    $("#seedcoll").val("");
    $("#seedsel").val("");
    $("#germconst").val("");
    $("#germmain").val("");
    $("#other").val("");
}

function slideinput() {
    window.location.href = "/AdvancedInputs";
}

function slidemetric() {
    window.location.href = "/TechnicianHome";
}

function saveMetrics() {
    var data1 = new Array();
    data1.push($("#coffeeparchment").val());
    data1.push($("#length").val());
    data1.push($("#farmarea").val());
    data1.push($("#weight").val());
    data1.push($("#capacity").val());
    data1.push($("#currency").val());
    console.log(data1);
    var data = JSON.stringify(data1);
    $.ajax({
        type: "POST",
        url: apiURL + "TechnicianHomeAPI/savemetrics",
        data: data,
        contentType: "application/json",
        success: function (result) {
            $("#savebutton").attr('style', 'background-color:#FFFFFF; float:right; border-color:bisque');
            $("#savebutton").attr('onclick', '');
        }
    });
}

function change() {
    $("#savebutton").attr('style', 'background-color:#00838F; float:right; border-color:bisque');
    $("#savebutton").attr('onclick', 'saveMetrics()');
}

function Metrics(apiURL) {
    debugger;
    $.ajax({
        type: "GET",
        url: apiURL + "TechnicianHomeAPI/metrics",
        contentType: "application/json",
        success: function (result) {
            console.log(result);
            var info = result["technicianloginfo"];
            var metrics = info["Metrics"];
            for (var property in metrics) {

                //console.log(property, metrics[property]);
                if (metrics[property]) {
                    if (property === "coffeemeasurekilograms") {
                        $("#coffeeparchment").val('Kilograms');
                    }
                    if (property === "coffeemeasurepounds") {
                        $("#coffeeparchment").val('Pounds');
                    }
                    if (property === "coffeemeasurecargas") {
                        $("#coffeeparchment").val('Cargas');
                    }
                    if (property === "coffeemeasurequintales") {
                        $("#coffeeparchment").val('Quintales');
                    }
                    if (property === "coffeemeasurearrobas") {
                        $("#coffeeparchment").val('Arrobas');
                    }

                    if (property === "lengthmeasurefeet") {
                        $("#length").val('Feet');
                    }
                    if (property === "lengthmeasuremeters") {
                        $("#length").val('Meters');
                    }
                    if (property === "farmareameasurehectares") {
                        $("#farmarea").val('Hectares');
                    }
                    if (property === "farmareameasuremanzanas") {
                        $("#farmarea").val('Manzanas');
                    }
                    if (property === "applicationmeasurekilograms") {
                        $("#weight").val('Kilograms');
                    }

                    if (property === "applicationmeasurepounds") {
                        $("#weight").val('Pounds');
                    }
                    if (property === "capacitymeasuregallons") {
                        $("#capacity").val('Gallons');
                    }

                    if (property === "capacitymeasureliters") {
                        $("#capacity").val('Liters');
                    }

                    if (property === "currencyboliviaboliviano") {
                        $("#currency").val('Bolivian Boliviano');
                    }
                    if (property === "currencybrazilreal") {
                        $("#currency").val('Brazilian Real');
                    }
                    if (property === "currencycolombiapeso") {
                        $("#currency").val('Colombian Peso');
                    }
                    if (property === "currencycostaricacolon") {
                        $("#currency").val('Costa Rican Colon');
                    }
                    if (property === "currencycubapeso") {
                        $("#currency").val('Cuban Peso');
                    }
                    if (property === "currencyguatemalaquetzal") {
                        $("#currency").val('Guatemalan Quetzal');
                    }
                    if (property === "currencyhaitigourde") {
                        $("#currency").val('Haitian Gourde');
                    }
                    if (property === "currencyhonduraslempira") {
                        $("#currency").val('Honduran Lempira');
                    }
                    if (property === "currencyjamaicadollar") {
                        $("#currency").val('Jamaican Dollar');
                    }
                    if (property === "currencymexicopeso") {
                        $("#currency").val('Mexican Peso');
                    }
                    if (property === "currencynicaraguacordoba") {
                        $("#currency").val('Nicaraguan Cordoba');
                    }
                    if (property === "currencyperusol") {
                        $("#currency").val('Peruvian Sol');
                    }
                    if (property === "currencyusdollar") {
                        $("#currency").val('USD');
                    }
                    if (property === "currencyvenezuelabolivar") {
                        $("#currency").val('Venezuelan Bolivar');
                    }
                    
                }
            }
            //$("#savebutton").attr('style', 'background-color:#ffffff; float:right; border-color:bisque');
        },
        error: function (res, status) {
            if (status === "error") {
                console.log("error");
            }
        }
    });
}

function laborslide() {
    $("#slider").html("");
    $("#slider").html(
        '<div class="justify-content-between">'
        + '<a href="#!" id="laborback" onclick="laborback()" class="btn btn-md u-btn-outline-primary g-mr-10 g-mb-15" style="-webkit-text-fill-color:white;background-color:#00838F;border-color:white;vertical-align:initial;">'
        + '<i class="hs-admin-angle-left" style="margin-left:-10px;margin-right:8px;font-size:small"></i>'
        + 'Labor'
        + '</a > '
        + '</div>'
        + '<ul class="list-unstyled" >'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1" style="cursor:pointer" onclick="laborsubmenu()">'
        + ' <div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Labor during establishment and vegetative growth years</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1">'
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Labor for farm maintenance, harvesting and processing</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '</ul>'
    );
}

function laborback() {
    $("#slider").html("");
    $("#slider").html(
        '<ul class="list-unstyled">'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1" style="cursor:pointer" onclick="laborslide()">'
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Labor</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1">'
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Additional Income and Remuneration</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1">'
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Cost of materials and Inputs</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1">'
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Equipment and reusable materials</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1">'
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Administrative Costs, taxes and land</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1">'
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Transportation</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '</ul>'
    );
}

function laborsubmenu() {
    $("#slider").html("");
    $("#slider").html(
        '<div class="justify-content-between">'
        + '<a href="#!" id="laborback" onclick="laborslide()" class="btn btn-md u-btn-outline-primary g-mr-10 g-mb-15" style="-webkit-text-fill-color:white;background-color:#00838F;border-color:white;vertical-align:initial;">'
        + '<i class="hs-admin-angle-left" style="margin-left:-10px;margin-right:8px;font-size:small"></i>'
        + 'Labor during establishment and vegetative growth years'
        + '</a > '
        + '</div>'
        + '<ul class="list-unstyled" >'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1" style="cursor:pointer" onclick="laborsubsubmenu()">'
        + ' <div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Germinator Labor</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1">'
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Nursery labor</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1" > '
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Land Preparation and Sowing labor</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1" > '
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Labor for the year corresponding to vegetative growth</strong>'
        + '<i class="hs-admin-angle-right"></i>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '</ul>'
    );
}

function laborsubsubmenu() {
    $("#slider").html("");
    $("#slider").html(
        '<div class="justify-content-between">'
        + '<a href="#!" id="laborback" onclick="laborsubmenu()" class="btn btn-md u-btn-outline-primary g-mr-10 g-mb-15" style="-webkit-text-fill-color:white;background-color:#00838F;border-color:white;vertical-align:initial;">'
        + '<i class="hs-admin-angle-left" style="margin-left:-10px;margin-right:8px;font-size:small"></i>'
        + 'Germinator labor'
        + '</a > '
        + '</div>'
        + '<ul class="list-unstyled" >'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1">'
        + ' <div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Describe in days, how much time is invested for one hectare of coffee.'
        + ' Each working day represents six hours of effective work(Ex: 3 hours = 0.5 days; 12 hours = 2 days).'
        + ' The total number of days is equal to: Number of people * Days * Number of times per year.'
        + ' Eg:If one activity requires 2 people, working 1 day and this activity is performed 3 times per year, then total days = 2*1*3=6.'
        + ' Write 0 if the activity is not done.</strong>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1">'
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Seed Collection</strong>'
        + '<div class=form-inline>'
        + '<input class="form-control form-rounded form-control-md" type="text" id="seedcoll" value="1.87" style="width:auto">'
        + '<strong style="margin-left:10px">Soles</strong>'
        + '</div>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1" > '
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Seed Selection</strong>'
        + '<div class=form-inline>'
        + '<input class="form-control form-rounded form-control-md" type="text" id="seedsel" value="1.52" style="width:auto">'
        + '<strong style="margin-left:10px">Soles</strong>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1" > '
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Germinator Construction</strong>'
        + '<div class=form-inline>'
        + '<input class="form-control form-rounded form-control-md" type="text" id="germconst" value="4.03" style="width:auto">'
        + '<strong style="margin-left:10px">Soles</strong>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1" > '
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Germinator maintenance-irrigation</strong>'
        + '<div class=form-inline>'
        + '<input class="form-control form-rounded form-control-md" type="text" id="germmain" value="8.82" style="width:auto" >'
        + '<strong style="margin-left:10px">Soles</strong>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1" > '
        + '<div class="media-body">'
        + '<div class="d-flex justify-content-between">'
        + '<strong>Other</strong>'
        + '<div class=form-inline>'
        + '<input class="form-control form-rounded form-control-md" type="text" id="other" value="0.88" style="width:auto" >'
        + '<strong style="margin-left:10px">Soles</strong>'
        + '</div>'
        + '</div>'
        + '</li>'
        + '</ul>'
    );
}