/*window.open è strutturato come window.open('link',...). per altre funzioni, copiare tutto e cambiare il link*/
/*al momento non servono a nulla queste funzioni perchè devono essere assegnate a pulsanti e non a celle*/

function Mail_PreAllerta() {
  var js = " \
<script> \
window.open('https://mail.google.com/mail/u/0/#sent/KtbxLxgNMxDHgdJztGMfJLdnpDJCLXltDB?compose=CllgCKHQdBCznMgXsXjhFqhzwGfJtnqGJJcCGNPfkpkQQJzpLGzrCXklFFKjJQrRqNlsGdJslBq', 'width=800, height=600'); \
google.script.host.close(); \
</script> \
";
  var html = HtmlService.createHtmlOutput(js)
  .setHeight(10)
  .setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sto caricando');
}


function Modifica_Modulo() {
  var js = " \
<script> \
window.open('https://docs.google.com/forms/d/1Ti8iXIreRWbrRrbOluKLwyr-UQ-bmhpLCGZBxwuPbSo/edit', 'width=800, height=600'); \
google.script.host.close(); \
</script> \
";
  var html = HtmlService.createHtmlOutput(js)
  .setHeight(10)
  .setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sto caricando');
}

function Mail_Attivazione() {
  var js = " \
<script> \
window.open('https://mail.google.com/mail/u/1/#search/emergenza+maltempo/FMfcgxvzLNbvQfWLgCmsQFGtmKHFlKDz?compose=jrjtXGkcsqMfJTKhNrVfplxxQxRNHTRbmZNDhXBMKcsVhnRnwqtpTtKqkzXKqhPRhtqLrbKp', 'width=800, height=600'); \
google.script.host.close(); \
</script> \
";
  var html = HtmlService.createHtmlOutput(js)
  .setHeight(10)
  .setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sto caricando');
}

function Apertura_Chiusura_Report() {
  var js = " \
<script> \
window.open('https://drive.google.com/drive/u/0/folders/1Kw_O8WswI2rYZh9_KsDL3VdTh8JU_NpI', 'width=800, height=600'); \
google.script.host.close(); \
</script> \
";
  var html = HtmlService.createHtmlOutput(js)
  .setHeight(10)
  .setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sto caricando');
}

function Benefici_di_Legge() {
  var js = " \
<script> \
window.open('https://drive.google.com/drive/u/0/folders/1cUpcL8jUD6omISZxghU2FFGSKUPfoSYs', 'width=800, height=600'); \
google.script.host.close(); \
</script> \
";
  var html = HtmlService.createHtmlOutput(js)
  .setHeight(10)
  .setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sto caricando');
}

/*funzione per apertura determinato tab. non usare per assegnazione script. NON TOCCARE*/
function showSheetByName(Name) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(Name);
  SpreadsheetApp.setActiveSheet(sheet);
}

/*apertura tab turni sol. Per gli altri basta cambiare il nome e usare il nome di questo per assegnare lo script*/
function Turni_SOL() {
  showSheetByName("turni S.O.L.");
}

/*aggiornamento data colonna b*/
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() == "Guida All'Emergenza") {
    var r = e.source.getActiveRange();
    if (r.getColumn() == 1) {
      sheet.getRange(r.getRow(),r.getColumn()+1).setValue(new Date());
    }
  }
}


/*reset colonna A, fogli risposte, fogli moduli operativi ==> reset emergenza, da usare una volta archiviata l'emergenza*/

function reset(){
  
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.alert(
    'Conferma il reset',
    "Sei sicuro di voler salvare e resettare il report?\nUna volta resettato, non potrai tornare indietro",
    ui.ButtonSet.YES_NO);
  
  // Process the user's response.
  if (result == ui.Button.YES) {
    ui.alert('Premi OK per resettare e attendi il messaggio "Reset completato!"');
    
    /*copia il file nella cartella id folder*/
    /*foglio di origine*/
    var source_sheet = SpreadsheetApp.getActiveSpreadsheet();
    /*id cartella di destinazione*/
    var dest_Folder = DriveApp.getFolderById("1u14oM0TImeW9N6NET22RVNiOc62alJpb");
    /*rinomina foglio di destinazione*/
    var source_name=SpreadsheetApp.getActiveSpreadsheet().getRange("d3:f3").getValue();
    var dest_file = SpreadsheetApp.create(source_name); /*fino a qui ok*/
    /*copia dei singoli sheets nel foglio di destinazione*/
    for (i in source_sheet.getSheets()){
    source_sheet.getSheets()[i].copyTo(dest_file);
    }
    DriveApp.getFileById(dest_file.getId()).makeCopy(source_name, dest_Folder);
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(dest_file.getId()));
    
    /*resetta*/
    /*colonna A, B, F, G*/
    var column_A=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Guida All'Emergenza");
    column_A.getRange("a5:a12").clearContent();
    column_A.getRange("a16:a49").clearContent();
    var column_B=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Guida All'Emergenza");
    column_B.getRange("b3:b12").clearContent();
    column_B.getRange("b16:b49").clearContent();
    var column_FG=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Guida All'Emergenza");
    column_FG.getRange("f3:g12").clearContent();
    column_FG.getRange("f16:g49").clearContent();
    column_FG.getRange("g7:g7").setValue("L'account Gaia da utilizzare deve essere del Delegato Area 3");
    column_FG.getRange("g17:g17").setValue("L'account Gaia da utilizzare deve essere del Delegato Area 3");
    column_FG.getRange("g23:g23").setValue("Assicurarsi di avere ICE e CF dei volontari che si devono attivare");
    var nome_emergenza=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Guida All'Emergenza");
    nome_emergenza.getRange("d3:f3").clearContent();
    nome_emergenza.getRange("d14:f14").clearContent();
    nome_emergenza.getRange("d3:f3").setValue("Nome Emergenza");
    nome_emergenza.getRange("d14:f14").setValue("Nome Emergenza");
    
    /*fogli risposta moduli preallerta/attivazione*/
    var Alert_Answer=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NON TOCCARE - Preallerta");
    var start = 2;
    var end_1 = Alert_Answer.getLastRow() -1;
    Alert_Answer.deleteRows(start, end_1);
    var Activated_Answer=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NON TOCCARE - Attivazione");
    var end_2 = Activated_Answer.getLastRow() -1;
    Activated_Answer.deleteRows(start, end_2);
    
    /*moduli operativi*/
    /*sol*/
    var sol_shift=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("turni S.O.L.");
    sol_shift.getRange("d4:f4").clearContent();
    sol_shift.getRange("h4").clearContent();
    sol_shift.getRange("e6:f6").clearContent();
    sol_shift.getRange("h6").clearContent();
    sol_shift.getRange("l6").clearContent();
    sol_shift.getRange("n6").clearContent();
    sol_shift.getRange("c10:q13").clearContent();
    sol_shift.getRange("b16:q18").clearContent();
    sol_shift.getRange("d4:f4").setValue("----");
    sol_shift.getRange("h4").setValue("----");
    sol_shift.getRange("e6:f6").setValue("----");
    sol_shift.getRange("h6").setValue("----");
    sol_shift.getRange("l6").setValue("----");
    sol_shift.getRange("n6").setValue("----");
    /*2° turno*/
    sol_shift.getRange("e23:f23").clearContent();
    sol_shift.getRange("h23").clearContent();
    sol_shift.getRange("l23").clearContent();
    sol_shift.getRange("n23").clearContent();
    sol_shift.getRange("c27:q30").clearContent();
    sol_shift.getRange("b33:q35").clearContent();
    sol_shift.getRange("e23:f23").setValue("----");
    sol_shift.getRange("h23").setValue("----");
    sol_shift.getRange("l23").setValue("----");
    sol_shift.getRange("n23").setValue("----");
    /*3° turno*/
    sol_shift.getRange("e39:f39").clearContent();
    sol_shift.getRange("h39").clearContent();
    sol_shift.getRange("l39").clearContent();
    sol_shift.getRange("n39").clearContent();
    sol_shift.getRange("c43:q46").clearContent();
    sol_shift.getRange("b49:q51").clearContent();
    sol_shift.getRange("e39:f39").setValue("----");
    sol_shift.getRange("h39").setValue("----");
    sol_shift.getRange("l39").setValue("----");
    sol_shift.getRange("n39").setValue("----");
    /*4° turno*/
    sol_shift.getRange("e55:f55").clearContent();
    sol_shift.getRange("h55").clearContent();
    sol_shift.getRange("l55").clearContent();
    sol_shift.getRange("n55").clearContent();
    sol_shift.getRange("c59:q62").clearContent();
    sol_shift.getRange("b65:q67").clearContent();
    sol_shift.getRange("e55:f55").setValue("----");
    sol_shift.getRange("h55").setValue("----");
    sol_shift.getRange("l55").setValue("----");
    sol_shift.getRange("n55").setValue("----");
    /*5° turno*/
    sol_shift.getRange("e71:f71").clearContent();
    sol_shift.getRange("h71").clearContent();
    sol_shift.getRange("l71").clearContent();
    sol_shift.getRange("n71").clearContent();
    sol_shift.getRange("c75:q78").clearContent();
    sol_shift.getRange("b81:q83").clearContent();
    sol_shift.getRange("e71:f71").setValue("----");
    sol_shift.getRange("h71").setValue("----");
    sol_shift.getRange("l71").setValue("----");
    sol_shift.getRange("n71").setValue("----");
    /*6° turno*/
    sol_shift.getRange("e87:f87").clearContent();
    sol_shift.getRange("h87").clearContent();
    sol_shift.getRange("l87").clearContent();
    sol_shift.getRange("n87").clearContent();
    sol_shift.getRange("c91:q94").clearContent();
    sol_shift.getRange("b97:q99").clearContent();
    sol_shift.getRange("e87:f87").setValue("----");
    sol_shift.getRange("h87").setValue("----");
    sol_shift.getRange("l87").setValue("----");
    sol_shift.getRange("n87").setValue("----");
    
    /*LOG Chiamate*/
    var call_log=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log Chiamate");
    call_log.getRange("a4:h100").clearContent();
    
    /*LOG mezzi*/
    var vehicle_log=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log Mezzi");
    vehicle_log.getRange("a4:l100").clearContent();
    
    /*BLS*/
    var bls=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BLS");
    /*1° turno*/
    bls.getRange("j2:k3").clearContent();
    bls.getRange("m2:m3").clearContent();
    bls.getRange("c5").clearContent();
    bls.getRange("e5").clearContent();
    bls.getRange("g5").clearContent();
    bls.getRange("i5").clearContent();
    bls.getRange("k5").clearContent();
    bls.getRange("m5").clearContent();
    bls.getRange("c7:e10").clearContent();
    bls.getRange("g7:i10").clearContent();
    bls.getRange("k7:m10").clearContent();
    /*2° turno*/
    bls.getRange("j12:k13").clearContent();
    bls.getRange("m12:m13").clearContent();
    bls.getRange("c15").clearContent();
    bls.getRange("e15").clearContent();
    bls.getRange("g15").clearContent();
    bls.getRange("i15").clearContent();
    bls.getRange("k15").clearContent();
    bls.getRange("m15").clearContent();
    bls.getRange("c17:e20").clearContent();
    bls.getRange("g17:i20").clearContent();
    bls.getRange("k17:m20").clearContent();
    /*3° turno*/
    bls.getRange("j22:k23").clearContent();
    bls.getRange("m22:m23").clearContent();
    bls.getRange("c25").clearContent();
    bls.getRange("e25").clearContent();
    bls.getRange("g25").clearContent();
    bls.getRange("i25").clearContent();
    bls.getRange("k25").clearContent();
    bls.getRange("m25").clearContent();
    bls.getRange("c27:e30").clearContent();
    bls.getRange("g27:i30").clearContent();
    bls.getRange("k27:m30").clearContent();
    /*4° turno*/
    bls.getRange("j32:k33").clearContent();
    bls.getRange("m32:m33").clearContent();
    bls.getRange("c35").clearContent();
    bls.getRange("e35").clearContent();
    bls.getRange("g35").clearContent();
    bls.getRange("i35").clearContent();
    bls.getRange("k35").clearContent();
    bls.getRange("m35").clearContent();
    bls.getRange("c37:e40").clearContent();
    bls.getRange("g37:i40").clearContent();
    bls.getRange("k37:m40").clearContent();
    /*5° turno*/
    bls.getRange("j42:k43").clearContent();
    bls.getRange("m42:m43").clearContent();
    bls.getRange("c45").clearContent();
    bls.getRange("e45").clearContent();
    bls.getRange("g45").clearContent();
    bls.getRange("i45").clearContent();
    bls.getRange("k45").clearContent();
    bls.getRange("m45").clearContent();
    bls.getRange("c47:e50").clearContent();
    bls.getRange("g47:i50").clearContent();
    bls.getRange("k47:m50").clearContent();
    /*6° turno*/
    bls.getRange("j52:k53").clearContent();
    bls.getRange("m52:m53").clearContent();
    bls.getRange("c55").clearContent();
    bls.getRange("e55").clearContent();
    bls.getRange("g55").clearContent();
    bls.getRange("i55").clearContent();
    bls.getRange("k55").clearContent();
    bls.getRange("m55").clearContent();
    bls.getRange("c57:e60").clearContent();
    bls.getRange("g57:i60").clearContent();
    bls.getRange("k57:m60").clearContent();
    
    /*ALS*/
    var als=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ALS");
    /*1° turno*/
    als.getRange("j3:k4").clearContent();
    als.getRange("m3:m4").clearContent();
    als.getRange("c6").clearContent();
    als.getRange("e6").clearContent();
    als.getRange("g6").clearContent();
    als.getRange("i6").clearContent();
    als.getRange("k6").clearContent();
    als.getRange("m6").clearContent();
    als.getRange("c8:e12").clearContent();
    als.getRange("g8:i12").clearContent();
    als.getRange("k8:m12").clearContent();
    /*2° turno*/
    als.getRange("j14:k15").clearContent();
    als.getRange("m14:m15").clearContent();
    als.getRange("c18").clearContent();
    als.getRange("e18").clearContent();
    als.getRange("g18").clearContent();
    als.getRange("i18").clearContent();
    als.getRange("k18").clearContent();
    als.getRange("m18").clearContent();
    als.getRange("c20:e24").clearContent();
    als.getRange("g20:i24").clearContent();
    als.getRange("k20:m24").clearContent();
    /*3° turno*/
    als.getRange("j26:k27").clearContent();
    als.getRange("m26:m27").clearContent();
    als.getRange("c30").clearContent();
    als.getRange("e30").clearContent();
    als.getRange("g30").clearContent();
    als.getRange("i30").clearContent();
    als.getRange("k30").clearContent();
    als.getRange("m30").clearContent();
    als.getRange("c32:e36").clearContent();
    als.getRange("g32:i36").clearContent();
    als.getRange("k32:m36").clearContent();
    /*4° turno*/
    als.getRange("j38:k39").clearContent();
    als.getRange("m38:m39").clearContent();
    als.getRange("c42").clearContent();
    als.getRange("e42").clearContent();
    als.getRange("g42").clearContent();
    als.getRange("i42").clearContent();
    als.getRange("k42").clearContent();
    als.getRange("m42").clearContent();
    als.getRange("c44:e48").clearContent();
    als.getRange("g44:i48").clearContent();
    als.getRange("k44:m48").clearContent();
    /*5° turno*/
    als.getRange("j50:k51").clearContent();
    als.getRange("m50:m51").clearContent();
    als.getRange("c54").clearContent();
    als.getRange("e54").clearContent();
    als.getRange("g54").clearContent();
    als.getRange("i54").clearContent();
    als.getRange("k54").clearContent();
    als.getRange("m54").clearContent();
    als.getRange("c56:e60").clearContent();
    als.getRange("g56:i60").clearContent();
    als.getRange("k56:m60").clearContent();
    /*6° turno*/
    als.getRange("j62:k63").clearContent();
    als.getRange("m62:m63").clearContent();
    als.getRange("c66").clearContent();
    als.getRange("e66").clearContent();
    als.getRange("g66").clearContent();
    als.getRange("i66").clearContent();
    als.getRange("k66").clearContent();
    als.getRange("m66").clearContent();
    als.getRange("c68:e72").clearContent();
    als.getRange("g68:i72").clearContent();
    als.getRange("k68:m72").clearContent();
    
    /*SAP*/
    var sap=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SAP");
    /*1° turno*/
    sap.getRange("j2:k3").clearContent();
    sap.getRange("m2:m3").clearContent();
    sap.getRange("c5").clearContent();
    sap.getRange("e5").clearContent();
    sap.getRange("g5").clearContent();
    sap.getRange("i5").clearContent();
    sap.getRange("k5").clearContent();
    sap.getRange("m5").clearContent();
    sap.getRange("c7:e10").clearContent();
    sap.getRange("g7:i10").clearContent();
    sap.getRange("k7:m10").clearContent();
    
    /*2° turno*/
    sap.getRange("j12:k13").clearContent();
    sap.getRange("m12:m13").clearContent();
    sap.getRange("c15").clearContent();
    sap.getRange("e15").clearContent();
    sap.getRange("g15").clearContent();
    sap.getRange("i15").clearContent();
    sap.getRange("k15").clearContent();
    sap.getRange("m15").clearContent();
    sap.getRange("c17:e20").clearContent();
    sap.getRange("g17:i20").clearContent();
    sap.getRange("k17:m20").clearContent();
    
    /*3° turno*/
    sap.getRange("j22:k23").clearContent();
    sap.getRange("m22:m23").clearContent();
    sap.getRange("c25").clearContent();
    sap.getRange("e25").clearContent();
    sap.getRange("g25").clearContent();
    sap.getRange("i25").clearContent();
    sap.getRange("k25").clearContent();
    sap.getRange("m25").clearContent();
    sap.getRange("c27:e30").clearContent();
    sap.getRange("g27:i30").clearContent();
    sap.getRange("k27:m30").clearContent();
    
    /*4° turno*/
    sap.getRange("j32:k33").clearContent();
    sap.getRange("m32:m33").clearContent();
    sap.getRange("c35").clearContent();
    sap.getRange("e35").clearContent();
    sap.getRange("g35").clearContent();
    sap.getRange("i35").clearContent();
    sap.getRange("k35").clearContent();
    sap.getRange("m35").clearContent();
    sap.getRange("c37:e40").clearContent();
    sap.getRange("g37:i40").clearContent();
    sap.getRange("k37:m40").clearContent();
    
    /*5° turno*/
    sap.getRange("j42:k43").clearContent();
    sap.getRange("m42:m43").clearContent();
    sap.getRange("c45").clearContent();
    sap.getRange("e45").clearContent();
    sap.getRange("g45").clearContent();
    sap.getRange("i45").clearContent();
    sap.getRange("k45").clearContent();
    sap.getRange("m45").clearContent();
    sap.getRange("c47:e50").clearContent();
    sap.getRange("g47:i50").clearContent();
    sap.getRange("k47:m50").clearContent();
    
    /*6° turno*/
    sap.getRange("j52:k53").clearContent();
    sap.getRange("m52:m53").clearContent();
    sap.getRange("c55").clearContent();
    sap.getRange("e55").clearContent();
    sap.getRange("g55").clearContent();
    sap.getRange("i55").clearContent();
    sap.getRange("k55").clearContent();
    sap.getRange("m55").clearContent();
    sap.getRange("c57:e60").clearContent();
    sap.getRange("g57:i60").clearContent();
    sap.getRange("k57:m60").clearContent();
    
    /*DAP*/
    var dap=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Squadre Monitoraggio");
    /*1° turno*/
    dap.getRange("e3:f3").clearContent();
    dap.getRange("h3").clearContent();
    dap.getRange("l3").clearContent();
    dap.getRange("n3:q3").clearContent();
    dap.getRange("l5:q5").clearContent();
    dap.getRange("c8:q12").clearContent();
    
    /*2° turno*/
    dap.getRange("e17:f17").clearContent();
    dap.getRange("h17").clearContent();
    dap.getRange("l17").clearContent();
    dap.getRange("n17:q17").clearContent();
    dap.getRange("l19:q19").clearContent();
    dap.getRange("c22:q26").clearContent();
    
    /*3° turno*/
    dap.getRange("e31:f31").clearContent();
    dap.getRange("h31").clearContent();
    dap.getRange("l31").clearContent();
    dap.getRange("n31:q31").clearContent();
    dap.getRange("l33:q33").clearContent();
    dap.getRange("c36:q40").clearContent();
    
    /*4° turno*/
    dap.getRange("e45:f45").clearContent();
    dap.getRange("h45").clearContent();
    dap.getRange("l45").clearContent();
    dap.getRange("n45:q45").clearContent();
    dap.getRange("l47:q47").clearContent();
    dap.getRange("c50:q54").clearContent();
    
    /*5° turno*/
    dap.getRange("e59:f59").clearContent();
    dap.getRange("h59").clearContent();
    dap.getRange("l59").clearContent();
    dap.getRange("n59:q59").clearContent();
    dap.getRange("l61:q61").clearContent();
    dap.getRange("c64:q68").clearContent();
    
       
  } else {
    ui.alert('Reset annullato');
  }
    ui.alert('Reset completato!');
}

