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