$(document).ready(function () {
    $(document).on("scroll", onScroll);
    
    //smoothscroll
    $('a[href^="#"]').on('click', function (e) {
        e.preventDefault();
        $(document).off("scroll");
        
        $('a').each(function () {
            $(this).removeClass('active');
        })
        $(this).addClass('active');
      
        var target = this.hash,
            menu = target;
        $target = $(target);
        $('html, body').stop().animate({
            'scrollTop': $target.offset().top+2
        }, 500, 'swing', function () {
            window.location.hash = target;
            $(document).on("scroll", onScroll);
        });
    });
});

function onScroll(event){
    var scrollPos = $(document).scrollTop();
    $('.collapse a').each(function () {
        var currLink = $(this);
        var refElement = $(currLink.attr("href"));
        if (refElement.position().top <= scrollPos && refElement.position().top + refElement.height() > scrollPos) {
            $('.collapse ul li a').removeClass("active");
            currLink.addClass("active");
        }
        else{
            currLink.removeClass("active");
        }
    });
}
function myFunction() {
var name = document.getElementById("name").value;
var email = document.getElementById("email").value;
var comments = document.getElementById('comments').value;

// Returns successful data submission message when the entered information is stored in database.
var dataString = 'name1=' + name + '&email1=' + email + '&comments1=' + comments ;
if (name == '' || email == '' || comments == '' ) {
alert("Please Fill All Fields");
} else {
// AJAX code to submit form.
$.ajax({
type: "POST",
url: "ajaxjs.php",
data: dataString,
cache: false,
success: function(html) {
alert(html);
}
});
}
return false;
}
function initMap() {
    var myLatLng = {lat: 17.490753, lng: 78.352628};
    // Create a map object and specify the DOM element for display.
    var map = new google.maps.Map(document.getElementById('googleMap'), {
      center: myLatLng,
      scrollwheel: false,
      zoom: 16
    });
    
    // Create a marker and set its position.
    var marker = new google.maps.Marker({
      map: map,
      position: myLatLng,
      title: 'Anil Company,Miyapur,INDIA'
    });
  }
  
var filePath = "D:/New System/Learning/Company/Email_Address_Data.xlsx"; 

function saveToExcel() 
{ 
var myApp = new ActiveXObject("Excel.Application"); 
myApp.visible = true; 
var xlCellTypeLastCell = 11; 
var myWorkbook = myApp.Workbooks.Open(filePath); 
var myWorksheet = myWorkbook.Worksheets(1); 
myWorksheet.Activate; 
objRange = myWorksheet.UsedRange; 
objRange.SpecialCells(xlCellTypeLastCell).Activate; 
newRow = myApp.ActiveCell.Row + 1; 
alert('newRow : '+newRow); 
strNewCell = "A" + newRow; 
alert('strNewCell : '+ strNewCell); 
myApp.Range(strNewCell).Activate; 
myWorksheet.Cells(newRow,1).value = f1.emailID.value; 
myApp.Workbooks.Close; 
myApp.Close; 
alert('Data successfully saved'); 
} 
