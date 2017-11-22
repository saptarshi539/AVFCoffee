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
            console.log(content)
            UserData.input = content.loginfo.Inputs
            console.log(UserData.input)
           // if (status != 'nocontent') {
          //      console.log
          //      
          //  }

        }
        ,
        error: function () {
            console.log('not successful');
        }

    });

  
