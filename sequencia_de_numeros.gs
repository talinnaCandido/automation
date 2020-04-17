function myFunction() {
  var repeticao = Number(Browser.inputBox('Digite um número :)'));
  
  for(var i=0; i < repeticao; i++){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getCurrentCell().setValue(i+1);
    spreadsheet.getCurrentCell().offset(1, 0).activate();
  }
};
