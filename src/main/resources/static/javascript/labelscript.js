$(document).ready(function(){
$(‘input[type=”file”]’).change(function(e){
var fileName = e.target.files[0].name;
document.getElementById(“spanid”).innerHTML = “File ‘” + fileName + “‘ is selected.”;
});
});