function onOpen(e) {

    var menu = SpreadsheetApp.getUi().createMenu('EduXa');
    menu.addItem('Enviar mail', 'sendEmails');
    menu.addToUi();
}

function sendEmails(){
    var html = HtmlService.createTemplateFromFile('body')
      .evaluate()
      .setWidth(550) 
      .setHeight(650)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    SpreadsheetApp.getUi()
                  .showModalDialog(html, 'Enviar Mail');
}


function sendMail(subject,body){
  
    var sheet = SpreadsheetApp.getActiveSheet();
    var dataRange = sheet.getDataRange();
     
    Logger.log('body: %s',body);
  
    for (var i = 2; i < dataRange.getNumRows()+1; i++) {
        var email = dataRange.getCell(i,1).getValue();
        var bodyCloned = {body: body.toString()};
       
        if (email.length > 1){
          
          Logger.log("num cols: %s", dataRange.getNumColumns()+1);
          
          for (var j=1; j < dataRange.getNumColumns()+1; j++){
            
            var field = dataRange.getCell(i,j).getValue();
            Logger.log("<"+j+">");
            var re = new RegExp("<"+j+">", 'g');
            bodyCloned.body = bodyCloned.body.replace(re,field); 
          
          }
         
          Logger.log('%s',email);
          
          MailApp.sendEmail({to:email, 
                             subject:subject,
                             htmlBody: bodyCloned.body});
          
        }

    }
}
