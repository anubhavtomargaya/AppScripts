function extract() {
  var ss = SpreadsheetApp.openById('1T2t3GCpoV4n5UFj1mAgToPMW9bKtqNNm-hx1FTeGlwk');
  var sheet = ss.getSheetByName('Sheet1');
  var labelname = "CoinAllotment";
  
  var newlabel = "coinParsed";
  var userlabels = GmailApp.getUserLabels()
  
  if (userlabels.includes(newlabel) ) {}
  else {
    var parsedlabel = GmailApp.createLabel(newlabel);
  }
  var threads = GmailApp.search ("subject:" + "allotment report" );
  // var rawthreads = GmailApp.search ("subject:" + "allotment report" );
    // var allotlabel = GmailApp.createLabel(labelname);
    // allotlabel.addToThreads(rawthreads)
    // Logger.log("label created")
  

  // var threads = GmailApp.search ("label:" + allotlabel.getName());


  var messages = GmailApp.getMessagesForThreads (threads);
  for (var i = 0 ; i < messages.length; i++) {
    for (var j = 0; j < messages[i].length; j++) {
      var values = [[]]
      var msgbody = messages[i][j].getPlainBody()
      var limitbody = msgbody.indexOf("Zerodha")
      var start = msgbody.match(/\d{4}-/)
      var strt = msgbody.indexOf(start)
      var impmsg = msgbody.slice(strt,limitbody-3)
      var date = impmsg.match(/(\d{4}-\d{2}-\d{2})/)
      var dateind = impmsg.indexOf(date)

      // var splitbody = impmsg.split(" ")
      var amt = impmsg.match(/(Rs.[^a-zA-Z]+ )+/g)
      // var info = impmsg.match(/(\d{2}-\d{2}-\d+\D+)+ /)
      // var fund = impmsg.match(/ [a-zA-Z]+ .* \D+/g)
      var fund = impmsg.match(/[a-zA-Z0-9 ]+/g)
      var ff = impmsg.toLowerCase().match(/.+[0-9]{4}/)
  
  
      // var fname = fund[0].slice(0,fund[0].length-5)
      var date = impmsg.match(/(\d{4}-\d{2}-\d+) [a-zA-Z 0-9]+ /g)
      Logger.log(impmsg)
      Logger.log(date)
      for  (var k=0; k<fund.length; k++) {
        var deets = amt[k].split(" ")
        var extractamt = deets[1]
        var nav = deets[4]
        var qty = deets[5].substring(0,5)
        values = [[date[k],fund[k],extractamt,nav,qty]]
      //   // Logger.log(date)
        sheet.getRange(sheet.getLastRow()+1, 1, 1, 5).setValues(values);
      // }

    }
}

parsedlabel.addToThreads(threads)
// allotlabel.removeFromThreads(threads);
// GmailApp.deleteLabel(allotlabel)
// Logger.log(threads)


   


}
