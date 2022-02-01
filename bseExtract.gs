function bseExtract() {
  var ss = SpreadsheetApp.openById('###-####');
  var sheet = ss.getSheetByName('Sheet7');
  var parsedlabel = GmailApp.createLabel("bseParsed")

  var bsethreads = GmailApp.search('from:(mgrpt@bseindia.in) (Details of your BSE trade)')
  for (var o=0;o<bsethreads.length;o++){
      var tlabel = bsethreads[o].getLabels()
      if (tlabel.includes(parsedlabel)) {
        }
      
      else {
        var msgs = GmailApp.getMessagesForThread(bsethreads[o])
        for (var k=0;k<msgs.length;k++){
            val = [[]]
            var bod = msgs[k].getPlainBody()
            var start = bod.indexOf('Buy Sell')
            var end = bod.indexOf('Total Trade Value')
            var table = bod.slice(start+18,end-74)
            var row = table.split("</tr>")
          // Logger.log(table)
              for (var h=0;h<row.length;h++) {
                if (row[h]!="") {
                  var values = [[]]
                  var val = row[h].split("</td>")
                  let scrip = val[1].slice(5);
                  let dt = val[2].slice(5,15)
                  let time = val[7].slice(20)
                  let price = val[9].match(/\d+.\d+/)
                  let qty = val[8].match(/\d+/)
                  let action = val[11].match(/[B|S]/)

                  values =[[dt+" " +time,scrip,price,qty,action]]
                  sheet.getRange(sheet.getLastRow()+1,1,1,5).setValues(values)

                  Logger.log(scrip + dt + " " + time + price + qty + action)
                  
                  }
              }
          }
      }
 }
 parsedlabel.addToThreads(bsethreads)
 Logger.log("import finished")
}
