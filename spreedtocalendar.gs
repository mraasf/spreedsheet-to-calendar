/**

 * Exportar eventos da tabela para o calendário

 */

function CriarAgenda() {

  var sheet = SpreadsheetApp.getActiveSheet();

  var headerRows = 1;  // Pula título das colunas

  var range = sheet.getDataRange();

  var data = range.getValues();

  var calId = "ufboad22q625akul149m2ula2g@group.calendar.google.com";

  var cal = CalendarApp.getCalendarById(calId);

  for (i=0; i<data.length; i++) {

    if (i < headerRows) continue; // Skip header row(s)

    var row = data[i];

    var date = row[10];  

    var title = row[1];  //coluna título

    var tstart = row[5]; //coluna data

    var tstop = row[6]; //coluna hora

    var loc = row[4]; // coluna local

    var desc = row[3]; //coluna descrição

    var id = row[15];

    var nucleo = row[2];

    var conv = row[11];

    // checa se o evento exixste

    try {

      var event = cal.getEventSeriesById(id);

    }

    catch (e) {

      

    }

    if (!event) {
       
      var newEvent = cal.createEvent(title+' '+tstop+' '+nucleo, new Date(tstart), new Date(tstart), {description:desc,location:loc}).getId();
      Logger.log('title '+title+' start '+tstart+' stop '+tstop+' desc '+desc+' loc'+loc+' concat '+conv); 
      
      
      row[15] = newEvent;  // grava event ID

    }

   

    debugger;

  }

  // lê todos os ids da tabela
  range.setValues(data);

 

}
