
function test() {
  var opt = document.createElement("option");
  opt.value = "testOption";
  opt.innerHTML = "Test Option";

  document.getElementById("testSelect").append(opt);
  document.getElementById("testSelect").removeAttribute("disabled");
  document.getElementById("testLabel").innerHTML = "TESTTESTTEST";
}
