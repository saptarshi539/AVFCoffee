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

function slideinput() {
    window.location.href = "/AdvancedInputs";
}

function slidemetric() {
    window.location.href = "/TechnicianHome";
}

function laborslide() {
    debugger;
    $("#slider").html("");
    $("#slider").html(
        '<div class="justify-content-between">'
        + '<a href="#!" id="laborback" onclick="laborback()" class="btn btn-md u-btn-outline-primary g-mr-10 g-mb-15" style="-webkit-text-fill-color:white;background-color:#00838F;border-color:white;vertical-align:initial;">'
        + '<i class="hs-admin-angle-left" style="margin-left:-10px;margin-right:8px;font-size:small"></i>'
        + 'Labor'
        + '</a > '
        + '</div>'
        + '<ul class="list-unstyled" >'
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1" onclick="laborslide()">'
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
        + '<li class="media g-brd-around g-brd-gray-light-v4 g-pa-20 g-mb-minus-1" onclick="laborslide()">'
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