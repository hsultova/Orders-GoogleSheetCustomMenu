//The event handler triggered when opening the spreadsheet. 
//Creates menu items and refresh the table.
function onOpen() {
  SpreadsheetApp.getUi() 
      .createMenu('Поръчки')
      .addItem('Нова поръчка', 'newOrderForm')
      .addItem('Опресни', 'refresh')
      .addItem('Извадка по дата', 'newReport')
      .addToUi();

  refresh();    
}

//New Order
function newOrderForm() {
   var html = HtmlService.createHtmlOutputFromFile('NewOrder')
      .setWidth(400)
      .setHeight(300);
      SpreadsheetApp.getUi()
      .showModalDialog(html, 'Нова поръчка');
}

//Adds new order as sheet row. Refresh the table.
function addnewOrder(formObject) {
  var sheet = SpreadsheetApp.getActiveSheet();
   
  sheet.appendRow([formObject.name,
                formObject.description,
                Utilities.formatDate(new Date(formObject.date), 'GMT +02:00', 'dd.MM.yyyy') ]);

  refresh();
}

//New Report
function newReport() {
   var html = HtmlService.createHtmlOutputFromFile('Report')
      .setWidth(400)
      .setHeight(300);
      SpreadsheetApp.getUi()
      .showModalDialog(html, 'Извадка');
}

function addNewReport(formObject) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();  
  var values = sheet.getDataRange().getValues(); 
  
  var newName = formObject.date;
  if(formObject.name){
    newName = formObject.name;
  }

  var newSheet;
  var newTable;
  if(formObject.documentType === "newSheet"){
    newSheet = addNewSheet(newName);
  }
  else if (formObject.documentType === "newTable"){
    newTable = SpreadsheetApp.create(newName);
    newTable.setSpreadsheetLocale("bg");
    newTable.setSpreadsheetTimeZone("GMT +02:00");
  }

  var row;
  var cell;
  var date;
  var formDate = new Date(Utilities.formatDate(new Date(formObject.date), "GMT +02:00", "yyyy-MM-dd"));

  for (var r=0; r<values.length; r++) {
    for (var c=0; c<values.length; c++) {
      row = values[r];
      cell = values[r][c];
      date = new Date(Utilities.formatDate(new Date(cell), "GMT +02:00", "yyyy-MM-dd"));
      if(date != "Invalid Date"){  
        if(date.getTime() === formDate.getTime()){
          if(newSheet){
            newSheet.appendRow(row);
          }
          else if(newTable){
            newTable.appendRow(row);
          }
        }
      }
    }
  }
}

//Refresh the data
function refresh() {
  //Refresh data, check dates and color them
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();

  var today = new Date();
  var tomorrow = getDateAfter(today, 1);

  var tomorrowRule =  SpreadsheetApp.newConditionalFormatRule()
    .whenDateEqualTo(tomorrow)
    .setBackground("#ff7a7a") // light red
    .setRanges([range])
    .build(); 
  var previousDateRule =  SpreadsheetApp.newConditionalFormatRule()
    .whenDateBefore(today)
    .setBackground("#9e9a9a") //grey
    .setRanges([range])
    .build();  
  var todayRule =  SpreadsheetApp.newConditionalFormatRule()
    .whenDateEqualTo(today)
    .setBackground("#d10011") //red
    .setRanges([range])
    .build();  
  var futureDateRule =  SpreadsheetApp.newConditionalFormatRule()
    .whenDateAfter(today)
    .setBackground("#6b9c4f") //green
    .setRanges([range])
    .build(); 
  
  sheet.clearConditionalFormatRules();
  var rules = sheet.getConditionalFormatRules();
  rules.push(tomorrowRule);
  rules.push(previousDateRule);
  rules.push(todayRule);
  rules.push(futureDateRule);

  sheet.setConditionalFormatRules(rules);   
}

//Adds a new sheet by specified name. If the sheet with that name exists, deletes it before creating the new one.
function addNewSheet(name){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var newSheet = sheet.getSheetByName(name);
  if (newSheet != null) {
    sheet.deleteSheet(newSheet);
  }
  return sheet.insertSheet(name);
}

//Gets date after specified days
function getDateAfter(date, days){
  return new Date(date.getTime() + (days * 24 * 60 * 60 * 1000));
}
