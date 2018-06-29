$("#loginBtn").click(function () {
    var selected = localStorage.getItem("selectedLanguage")
    window.location.href = "Account/SignIn?Lang=" + selected
})

$("#farmerlogin").click(function () {
    var phoneNumber = $("#home-demobtn").val();
    console.log(data);
    $.ajax({
        type: "GET",
        url: apiURL + "CellSum/FarmerLogin",
        data: phoneNumber,
        contentType: "application/json; charset=utf-8",
        success: function (result, status) {
            
        },
        error: function (res, status) {
            if (status === "error") {
                console.log("error");
            }
        }
    });
})


