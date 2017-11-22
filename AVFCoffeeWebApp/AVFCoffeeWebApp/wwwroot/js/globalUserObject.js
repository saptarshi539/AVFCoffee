var UserData = {
    input: {},
    output: {},
    userData: {}
};

 
 $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        url: apiURL + "CellSum/getinput",
        success: function (content, status) {
            if (status != 'nocontent') {
                UserData.input = content.loginfo.Inputs;
                UserData.output = content.loginfo.Outputs;
                //UserData.userData = content.loginfo.User

                //load users data into input page
                $("#earlyHectares").val(UserData.input.earlyHectares);
                $("#peakHectares").val(UserData.input.peakHectares);
                $("#oldHectares").val(UserData.input.oldHectares);
                $("#conventional").prop("checked", UserData.input.conventional);
                $("#organic").prop("checked", UserData.input.organic);
                $("#transition").prop("checked", UserData.input.transition);
                $("#workerSalarySoles").val(UserData.input.workerSalarySoles);
                $("#productionQuintales").val(UserData.input.productionQuintales);
                $("#transportCostSoles").val(UserData.input.transportCostSoles);
                $("#costPriceSolesPerQuintal").val(UserData.input.costPriceSolesPerQuintal);
            }
        },
        error: function () {
            console.log('not successful');
        }

});

  
