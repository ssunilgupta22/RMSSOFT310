function openTabs(evt, cityName) {
    // Declare all variables
    var i, tabcontent, tablinks;

    // Get all elements with class="tabcontent" and hide them
    tabcontent = document.getElementsByClassName("tabcontent");
    for (i = 0; i < tabcontent.length; i++) {
        tabcontent[i].style.display = "none";
    }

    // Get all elements with class="tablinks" and remove the class "active"
    tablinks = document.getElementsByClassName("tablinks");
    for (i = 0; i < tablinks.length; i++) {
        tablinks[i].className = tablinks[i].className.replace(" active", "");
    }

    // Show the current tab, and add an "active" class to the link that opened the tab
    document.getElementById(cityName).style.display = "block";
    evt.currentTarget.className += " active";
} 

function onsidebarMenu(){
    document.getElementById("sidebarMenu").style.display = "block";
    
    document.getElementById("sidebarMenu").innerHTML = '<ul class="sidebarMenuInner" id="sidebarMenuInner">'+
    '<li><a href="./download_install.html">Download and Install</a></li>'+
    '<li><a href="./settings.html">Change Settings</a></li>'+
    '<li><a href="./othsettings.html">Change Print Page Settings</a></li>'+
    '<li><a href="./openaccounts.html">Add Supplier, Customer Account</a></li>'+
    '<li><a href="./itemsentry.html">Add/Modify Items</a></li>'+
    '<li><a href="./pbpsettings.html">Customize Sale Bill PrintOut Page</a></li>'+
    '<li><a href="./printersetup.html">Printer SetUp</a></li>'+
    '<li><a href="./gridtablesettings.html">Change Grid Table Settings</a></li>'+
    '<li><a href="./purchase.html">Purchase Panel</a></li>'+
    '<li><a href="./sale.html">Sales Panel</a></li>'+
    '<li><a href="./searchpursale.html">Search Purchase/Sale/Order</a></li>'+
    '<li><a href="./stockadj.html">Stock Adjustment</a></li>'+
    '<li><a href="./gstrreports.html">GST Reports</a></li>'+
    '<li><a href="./receipt.html">Receipt or Payment</a></li>'+
    '<li><a href="./ledger.html">Ledger A/C Search</a></li>'+
    '<li><a href="./backuprestore.html">BackUp and Restore Data</a></li></ul>'
    ;
}

function onsidebarMenuClose() {
    document.getElementById("sidebarMenu").style.display = "none";
    
}

function openburgarNav() {
    
    document.getElementById("bugarnav").style.width = "25%";
    document.getElementById("bugarnav").style.height = "100%";
    document.getElementById("x_Button").innerHTML = "x";
    document.getElementById("burgaropenid").innerHTML = "";
    document.getElementById("sidebarMenuInner").innerHTML = 
    '<li><a href="/rmsdoc/download_install.html">Download and Install</a></li>'+
    '<li><a href="/rmsdoc/settings.html">Change Settings</a></li>'+
    '<li><a href="/rmsdoc/othsettings.html">Change Print Page Settings</a></li>'+
    '<li><a href="/rmsdoc/openaccounts.html">Add Supplier, Customer Account</a></li>'+
    '<li><a href="/rmsdoc/itemsentry.html">Add/Modify Items</a></li>'+
    '<li><a href="/rmsdoc/pbpsettings.html">Customize Sale Bill PrintOut Page</a></li>'+
    '<li><a href="/rmsdoc/printersetup.html">Printer SetUp</a></li>'+
    '<li><a href="/rmsdoc/gridtablesettings.html">Change Grid Table Settings</a></li>'+
    '<li><a href="/rmsdoc/purchase.html">Purchase Panel</a></li>'+
    '<li><a href="/rmsdoc/sale.html">Sales Panel</a></li>'+
    '<li><a href="/rmsdoc/searchpursale.html">Search Purchase/Sale/Order</a></li>'+
    '<li><a href="/rmsdoc/stockadj.html">Stock Adjustment</a></li>'+
    '<li><a href="/rmsdoc/gstrreports.html">GST Reports</a></li>'+
    '<li><a href="/rmsdoc/receipt.html">Receipt or Payment</a></li>'+
    '<li><a href="/rmsdoc/ledger.html">Ledger A/C Search</a></li>'+
    '<li><a href="/rmsdoc/backuprestore.html">BackUp and Restore Data</a></li>'
    ;
}

function closeburgarNav() {

    document.getElementById("bugarnav").style.width = "0%";
    document.getElementById("bugarnav").style.height = "0%";
    document.getElementById("x_Button").innerHTML = ""; 
    document.getElementById("navcontent").innerHTML = "";
    document.getElementById("burgaropenid").innerHTML = "&#9776 OPEN";
}


var slideIndex = 0;
showSlides();

function showSlides() {
  var i;
  var slides = document.getElementsByClassName("mySlides");
  for (i = 0; i < slides.length; i++) {
    slides[i].style.display = "none";
  }
  slideIndex++;
  if (slideIndex > slides.length) {slideIndex = 1}
  try{
    slides[slideIndex-1].style.display = "block";
  }catch(err){}
  setTimeout(showSlides, 2000); // Change image every 2 seconds
} 