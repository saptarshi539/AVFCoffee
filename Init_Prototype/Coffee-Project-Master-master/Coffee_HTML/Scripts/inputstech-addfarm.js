
$(document).ready(function(){

	$('#submitFarm').on('click', function(){
		console.log('clicked');
		var resultsJSON = collectData();

		// Construct our url
        
        //Chris's way
//		var currentPath = window.location.pathname.split('/').slice(1, -1).join('/');

        //Liza's way 
		var currentPath = window.location.pathname
        currentPath= currentPath.substr(0,currentPath.lastIndexOf('/'));
        
        
      //  var currentPath = window.location.pathname
        var placeToGo = currentPath + '/' + 'farm1.html';
		var url = placeToGo + '?' + JSON.stringify(resultsJSON);
        
		// Now we go to our new page
		window.location.href = url;
	})

	$('div.container.toppadding').on('click', function(){

		console.log('parent clicked');
	})


})


function collectData(){

	var youngTrees = $('#youngTreeInput').val();
	var matureTrees = $('#matureTreeInput').val();
	var oldTrees = $('#oldTreeInput').val();
	var farmingMethod = $('input[name="Radios"]').val();
	var workerPay = $('#workerPay').val();
	var productivity = $('#productivity').val();

	var resultJSON = {'youngTrees':youngTrees, 'matureTrees':matureTrees, 'oldTrees':oldTrees, 'farmingMethod':farmingMethod, 'workerPay':workerPay, 'productivity':productivity}

	return resultJSON
}