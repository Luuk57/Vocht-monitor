const SHEET_NAME = "Metingen";

function doGet() {
  return ContentService.createTextOutput("VOCHT MONITOR is actief.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) sh = ss.insertSheet(SHEET_NAME);

    const headers = [
      "Opslagtijd","Start datum","Eind datum","Dialyse","Toeslag kg",
      "Gewicht vóór kg","Drooggewicht kg","Gewicht na kg","Geplande UF kg",
      "Werkelijke UF kg","Bloeddruk / HR / MAP","Drinkvocht ml",
      "Eten / overig ml","Totaal vocht ml","Dranklog","Notities"
    ];

    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.getRange(1,1,1,headers.length).setFontWeight("bold");
      sh.setFrozenRows(1);
    }

    sh.appendRow([
      new Date(),
      data.startDate || "",
      data.endDate || "",
      data.dialyseType || "",
      data.toeslag ?? "",
      data.preWeight ?? "",
      data.dryWeight ?? "",
      data.postWeight ?? "",
      data.geplandeUF ?? "",
      data.werkelijkeUF ?? "",
      data.bloeddruk || "",
      data.drinkvocht ?? 0,
      data.etenVocht ?? 0,
      data.totaalVocht ?? 0,
      data.drankLog || "",
      data.notities || ""
    ]);

    sh.autoResizeColumns(1, headers.length);

    return ContentService.createTextOutput(JSON.stringify({
      success:true, message:"Opgeslagen"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success:false, error:String(err)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}