var SpreadSheetID = "1DfLJVyB58SuOBEAKeg9YYC4WNE3JYsw-J60PytvcwnU"
var SheetName = "Chemical Inventory"


function BacteriaInventory() {
  var ss = SpreadsheetApp.openById(SpreadSheetID);
  var inventory = ss.getSheetByName(SheetName);

  var bac_inv = getData(inventory);

  const now = new Date();
  const MILLS_PER_DAY = 1000 * 60 * 60 * 24;
  var plus_month = new Date(now.getTime() + 30*MILLS_PER_DAY)

  // console.log(bac_inv);
  // console.log(Object.keys(bac_inv).length);

  // for(var i = 0, l= bac_inv.length; i<l ; i++){
  //   if (bac_inv[i]['Earliest EXP Date'] < plus_month){
  //     console.log('order ' + bac_inv[i]['Name'] + ' now');
  //   }
  // }

  // // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Date/parse
  // // parsing does not evaluate an empty earliest exp date, if statement is not true
  // // add !bac_inv[i]['Earliest EXP Date']
  // for(var i = 0, l= bac_inv.length; i<l ; i++){
  //   if (Date.parse(bac_inv[i]['Earliest EXP Date']) < plus_month || !bac_inv[i]['Earliest EXP Date']){
  //     console.log('order ' + bac_inv[i]['Name'] + ' now');
  //   }
  // }

  restock = [];
  
  // l should be number of rows
  for(var i = 0, l= bac_inv.length; i<l ; i++){
    if (bac_inv[i]['Quantity in Stock'] <= bac_inv[i]['Reorder Level'] || bac_inv[i]['Earliest EXP Date'] < plus_month || !bac_inv[i]['Earliest EXP Date']){
      item = bac_inv[i]['Name'];
      size = bac_inv[i]['Weight/Volume'];
      num_left = bac_inv[i]['Quantity in Stock'];
      exp = bac_inv[i]['Earliest EXP Date'];

      // TODO: HOW TO HANDLE BLANKS AND NAs: dont send email or do?
      if (!bac_inv[i]['Earliest EXP Date'] || bac_inv[i]['Earliest EXP Date'] == "NA"){
        // MailApp.sendEmail({to: "joelle.carbonell@enthalpy.com", subject: item + ", " + size, htmlBody: num_left + " units of " + item + ", " + size + " left. The earliest expiration date is blank", noReply:true})
        restock.push(bac_inv[i]);
      }

      // "ambhatnagar@montrose-env.com"

      else{
        if (bac_inv[i]['Earliest EXP Date'] < plus_month){
          restock.push(bac_inv[i]);
          // MailApp.sendEmail({to: "ambhatnagar@montrose-env.com", subject: item + ", " + size + " expires soon", htmlBody: num_left + " units of " + item + ", " + size + " left. The earliest expiration date is " + Utilities.formatDate(exp, 'America/New_York', 'MMMM dd, yyyy'), noReply:true})

          // MailApp.sendEmail({to: "joelle.carbonell@enthalpy.com", subject: item + ", " + size + " expires soon", htmlBody: num_left + " units of " + item + ", " + size + " left. The earliest expiration date is " + Utilities.formatDate(exp, 'America/New_York', 'MMMM dd, yyyy'), noReply:true})
        }

        else{
          restock.push(bac_inv[i]);
          // MailApp.sendEmail({to: "ambhatnagar@montrose-env.com", subject: item + ", " + size, htmlBody: num_left + " units of " + item + ", " + size + " left. The earliest expiration date is " + Utilities.formatDate(exp, 'America/New_York', 'MMMM dd, yyyy'), noReply:true})

          // MailApp.sendEmail({to: "joelle.carbonell@enthalpy.com", subject: item + ", " + size, htmlBody: num_left + " units of " + item + ", " + size + " left. The earliest expiration date is " + Utilities.formatDate(exp, 'America/New_York', 'MMMM dd, yyyy'), noReply:true})
        }
      }
    }
  }

  // console.log(restock.length);
  // console.log(restock);

  MailApp.sendEmail({to: "ambhatnagar@montrose-env.com",
                subject: "Items to Restock",
                htmlBody: printStuff(restock),
                noReply:true});


  MailApp.sendEmail({to: "joelle.carbonell@enthalpy.com",
                  subject: "Items to Restock",
                  htmlBody: printStuff(restock),
                  noReply:true});
}

function getData(inventory){
  var dataArray = [];
  var rows = inventory.getRange(3,1,inventory.getLastRow()-1, inventory.getLastColumn()).getValues();

  for(var i = 0, l= rows.length; i<l ; i++){
    if (rows[i][2] !== ''){
      var dataRow = rows[i];
      var record = {};
      record['Name'] = dataRow[2];
      record['Supplier'] = dataRow[3];
      record['Weight/Volume'] = dataRow[8];
      record['Quantity in Stock'] = dataRow[9];
      record['Reorder Level'] = dataRow[10];
      record['Earliest EXP Date'] = dataRow[12];
      record['EXP Dates'] = dataRow[13];
      dataArray.push(record);
    }
  }
  return dataArray;
}

// https://stackoverflow.com/questions/58767859/how-to-send-a-table-in-an-email-in-google-scripts
// please let this format stuff!
function printStuff(restock){
  string = "<html><body><br><table border=1><tr><th>Item</th><th>Quantity</th><th>Expiration</th></tr></br>";
  for (var i=0; i<restock.length; i++){
    exp = restock[i]['Earliest EXP Date'];
    string = string + "<tr>";

    // just have to keep separating NA and blanks from everything else
    if (!exp || exp == "NA"){
      temp = `<td> ${restock[i]['Name']} </td><td> ${restock[i]['Quantity in Stock']} </td>`;
      // temp = JSON.stringify("<td>" + restock[i]['Name']) + "</td>" + "<td>" + JSON.stringify(restock[i]['Quantity in Stock']) + "</td>";
      string = string.concat(temp);
    }
    else{
      temp = `<td> ${restock[i]['Name']} </td><td> ${restock[i]['Quantity in Stock']}  </td><td> ${Utilities.formatDate(exp, 'America/New_York', 'MMMM dd, yyyy')}</td>`;
      // temp = JSON.stringify("<td>" + restock[i]['Name']) + "</td>" + "<td>" + JSON.stringify(restock[i]['Quantity in Stock'])  + "</td>" + "<td>" + JSON.stringify(Utilities.formatDate(exp, 'America/New_York', 'MMMM dd, yyyy')) + "</td>";
      // temp = JSON.stringify(restock[i]['Name']) + ":,,,, [quantity] [" + JSON.stringify(restock[i]['Quantity in Stock']) + "],,,,[expiration] [" + JSON.stringify(Utilities.formatDate(exp, 'America/New_York', 'MMMM dd, yyyy')) + "] |***********************************************************|";
      // temp = JSON.stringify(Utilities.formatDate(exp, 'America/New_York', 'MMMM dd, yyyy'));
      // temp = JSON.stringify(restock[i]['Name']) + ">>>>>>>" + JSON.stringify(restock[i]['Quantity in Stock']) + "<<<<<<<" + JSON.stringify(Date.parse(restock[i]['Earliest EXP Date']).format("mmm d, yyyy")) + "|````````````|";
      // temp = JSON.stringify(restock[i]['Name']) + ": " + JSON.stringify(restock[i]['Quantity in Stock']) + JSON.stringify(Utilities.formatDate(Date.parse(restock[i]['Earliest EXP Date'], 'America/New_York', 'MMMM dd, yyyy'))) + "\n";
      string = string.concat(temp);
    }
    string = string + "</tr>";
  }
  string = string + "</table></body></html>";
  return string;
}
