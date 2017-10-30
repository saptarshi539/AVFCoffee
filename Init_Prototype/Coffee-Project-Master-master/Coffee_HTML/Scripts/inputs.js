
// Globals (We may wanted to put these in a class later)
var GlobalPassedFarm;

// This function parses the data from the url
parseFarmFromUrl();





$(document).ready(function(){
    

	$('#submit').on('click', function(){
        alert("gahh")

		console.log('clicked');
		var resultsJSON = collectData();

		// Construct our url
        
        //Chris's way
//		var currentPath = window.location.pathname.split('/').slice(1, -1).join('/');

        //Liza's way 
		var currentPath = window.location.pathname
        currentPath= currentPath.substr(0,currentPath.lastIndexOf('/'));
        
        
      //  var currentPath = window.location.pathname
        var placeToGo = currentPath + '/' + 'Results.html';
		var url = placeToGo + '?' + JSON.stringify(resultsJSON);
        
		// Now we go to our new page
		window.location.href = url;
	})

	$('div.container.toppadding').on('click', function(){

		console.log('parent clicked');
	})


})


function collectData(){
    var name = $('#name').val();

	var resultJSON = {'name':name}

	return resultJSON
}