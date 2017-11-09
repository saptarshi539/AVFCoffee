// when user clicks next link, it checks href for the next div to show 
// and hides the others 

$('.demo-screen').hide();
$('#demo-screen1').fadeIn(1000);

$('.slide').click(function () {
    var next = $(this).attr("href");
    if (next.startsWith("#")) {
        $('.demo-screen').hide();
        $(next).fadeIn(1000)
    }
});

