var app = SpreadsheetApp;
var spreadsheet = app.getActiveSpreadsheet();
var drive = DriveApp
var id = app.getActiveSpreadsheet().getId();
var folder=drive.getFolderById("1iquSJgdRUbaCnjt9cb0DWWfzcxU0P2Pi")
var sheetImpressao = spreadsheet.getSheetByName("menu")
function clearFiles(){
  let files = folder.getFiles();
  while(files.hasNext()){
    files.next().setTrashed(true); 
  }
  
}
function enviarpdf(){
  clearFiles()
  let tempSpreadsheet=app.open(drive.getFileById(id).makeCopy("tmp_spreadsheet",folder))
  let sheets = tempSpreadsheet.getSheets()
  let nomes = [];
  let emails = [];
  let values = sheetImpressao.getDataRange().getValues()
    values.map((elemt,ind,obj)=>{
               if(ind>0){
      if (elemt[0]!="null"){ 
          nomes.push(elemt[0])
          emails.push(elemt[1])

      }

    }
  });
  for (i=0;i<nomes.length-1;i++){
      var nome = nomes[i].trim();
      let tempSpreadsheet=app.open(drive.getFileById(id).makeCopy("tmp_spreadsheet",folder))
      sheets.map(elem=>{
                 Logger.log("-----------")
                 var nomeplan = elem.getSheetName().trim();
                 if(nomeplan!=nome){
                   tempSpreadsheet.deleteSheet(elem)
                 }
                             
      });
                  let pdf2 = tempSpreadsheet.getBlob().getAs("application/pdf").setName("pdf");
                  let newFile=folder.createFile(pdf2);
                  drive.getFileById(tempSpreadsheet.getId()).setTrashed(true)
                  let pdf = folder.getFilesByName("pdf").next().getAs("application/pdf")
                  MailApp.sendEmail(emails[i].trim(),"Assunto","Corpo email",{attachments:[pdf]})

     Logger.log(nome);
               
}
}
