function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Extract Emails').addItem('Extractupi','upiExtract')
      .addToUi()

}

function upiExtract() {
  var ss = SpreadsheetApp.openById('1T2t3GCpoV4n5UFj1mAgToPMW9bKtqNNm-hx1FTeGlwk');
  var sheet = ss.getSheetByName('Sheet6');
  var parsedlabel = GmailApp.createLabel("HdfcParsed")

  var hdfcthreads = GmailApp.search('from:(alerts.hdfcbank.net) (to VPA)')
  for (var o=0;o<hdfcthreads.length;o++){
      var tlabel = hdfcthreads[o].getLabels()
      if (tlabel.includes(parsedlabel)) {
        }
      
      else {
        var msgs = GmailApp.getMessagesForThread(hdfcthreads[o])
        for (var k=0;k<msgs.length;k++){
          val = [[]]
          var bod = msgs[k].getPlainBody()
          var start = bod.indexOf('Rs.')
          var end = bod.indexOf(bod.match(/\d{2}-\d{2}-\d{2}/))
          var row = bod.slice(start,end+10)
          if (row) {
            var content = row.split(" ")
            let amt = content[0].slice(3)
            let deb = content[3]
            let to = content[9]
            let dt = content[11].trim().slice(0,8)
            Logger.log(content)
            Logger.log(amt+ deb + to + dt)
            val = [[dt,amt,to]]
            sheet.getRange(sheet.getLastRow()+1,1,1,3).setValues(val)
          }
          else {
            
          }
          
       }
      }
    

  }
  parsedlabel.addToThreads(hdfcthreads)
  Logger.log("import finished")
  
  
}


