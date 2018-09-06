$("#technician-login-button").click(function () {
    var selected = localStorage.getItem("selectedLanguage")
    window.location.href = "Account/SignIn?Lang=" + selected
})

//$("#farmerlogin").click(function () {
//    var phoneNumber = $("#home-demobtn").val();
//    console.log(phoneNumber);
//    $.ajax({
//        type: "GET",
//        url: apiURL + "CellSum/FarmerLogin",
//        data: phoneNumber,
//        contentType: "application/json; charset=utf-8",
//        success: function (result, status) {
//            console.log(result);
//        },
//        error: function (res, status) {
//            if (status === "error") {
//                console.log("error");
//            }
//        }
//    });
//})

function abc(apiURL) {
    debugger;
    var phoneNumber = $("#farmer-login-text-field").val();
    localStorage.setItem("farmerPhone", phoneNumber);
    console.log(phoneNumber);
    $.ajax({
        type: "GET",
        url: apiURL + "CellSum/FarmerLogin?",
        data: "phoneNumber=" + phoneNumber,
        contentType: "application/json; charset=utf-8",
        success: function (result, status) {
            //console.log(result);
            if (result === true) {
                console.log("true")
                window.location.href = "/input"
            } else {
                $("#incorrectPhone").html("Please enter correct Phone number");
            }

        },
        error: function (res, status) {
            if (status === "error") {
                console.log("error");
            }
        }
    });
}


