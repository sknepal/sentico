// Developed by Subigya Kumar Nepal
// Thelacunablog.com
// @SkNepal
// August 23, 2015

function myFunction() {
  var api = 'IndicoAPIKey';
  analyse('NameOfTheSheet', 'EmailOfTheSender', api);
  analyse('NameOfTheSheet', 'EmailOfTheSender', api);
}

function analyse(sheetName, email, api){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var threads = GmailApp.search("from:" + email); // search for emails from a particular sender
  var length = threads.length;
  for (var i = 0; i < length; i++) { // go through each email of that sender
    var msg= threads[i].getMessages();
    var raw = msg[0].getPlainBody(); // get body of the email
    var wordCount = raw.match(/\S+/g).length; // get number of words in the email
    var data = { // prepare a json of the body text of the email in order to send POST request with it
    "data":raw
    };
    var payload = JSON.stringify(data);
    var options = {
      "method" : "POST",
      "contentType" : "application/json",
      "payload" : payload
    };
    var url = "http://apiv2.indico.io/sentiment?key=" + api;
    var response = UrlFetchApp.fetch(url, options).getContentText(); // a POST request with the email's body text to check its sentiment
    response = JSON.parse(response);
    var positivity = response.results;  // get the response of the request
    var positivityFixed = positivity.toFixed(); // round-off the value
    sheet.appendRow([wordCount, positivity, positivityFixed]); // add it to the spreadsheet
  }
  createScatterChart(sheet, length);
  createPieChart(sheet, length);
}

function createScatterChart(sheet, length) {
  var scatterChart = sheet.newChart()
     .setChartType(Charts.ChartType.SCATTER)
     .addRange(sheet.getRange('A1:B' + length))
     .setPosition(25, 5, 0, 0)
     .asScatterChart()
     .setTitle('Number of words vs Positivity')
     .setXAxisTitle('Number of words')
     .setYAxisTitle('Positivity')
     .setXAxisLogScale()
     .build();
  sheet.insertChart(scatterChart);
}

function createPieChart(sheet, length){
  calculateTotal(sheet, length);
  var pieChart = sheet.newChart()
     .setChartType(Charts.ChartType.PIE)
     .asPieChart()
     .addRange(sheet.getRange('F2:G3'))
     .setPosition(5, 5, 0, 0)
     .setTitle('Email Sentiment Analysis')
     .build();
  sheet.insertChart(pieChart);
}

function calculateTotal(sheet, length){
  sheet.getRange('F2').setValue("Positive (~1)");
  sheet.getRange('F3').setValue("Negative (~0) ");
  var sum = 0;
  for (var i=1; i<length ; i++){
      sum += sheet.getRange('C' + String(i)).getValue(); // add the 1s
  }
  sheet.getRange('G2').setValue(sum);
  sheet.getRange('G3').setValue(length-sum); // get the 0s.
}

