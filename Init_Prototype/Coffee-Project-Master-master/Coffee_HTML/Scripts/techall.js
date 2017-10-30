
// Globals (We may wanted to put these in a class later)
var GlobalPassedFarm;

// This function parses the data from the url
parseFarmFromUrl();



$(document).ready(function(){

	// Check if data was passed (if GlobalPassedFarm is undefined)
	if (GlobalPassedFarm){
		addFarmIcon();
	}
  
})

function parseFarmFromUrl(){

	// Get the text after the url
	var query = window.location.search.substring(1).split("&");

	// Decode the variable that we took from the url
	query = decodeURIComponent(query[0]);
   var nameval= JSON.parse(query);
   nameval = nameval.name;
	if (query != ""){
		// Parse the text into json format
		GlobalPassedFarm = JSON.parse(query);
	}
	

}

function addFarmIcon(){
  //get name from url
    var query = window.location.search.substring(1).split("&");
   query = JSON.parse(decodeURIComponent(query[0]));
    
    
    var name =  query.name;


	var farmHtml = ['<td id="FarmExample">',
                        '<a href="Farm1.html">',
						'<img class="img-responsive text-center farmicon" src="assets/img/farmicon.svg" alt="farm9">',
						name,
					'</td>'].join('');

	// Get the container for the farm icon

	$('tbody.farmtext tr:nth-child(3)').append(farmHtml);

	// Attach event onto our newly added element
	$('#FarmExample').on('click', function(){
		console.log('I am clicked');
		// Go to URL we want and pass data along. See 'inputstech-addfarm' for details on how to pass data through url. 
	});


}
function collectNames(){
    
}