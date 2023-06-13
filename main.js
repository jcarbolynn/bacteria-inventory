var SpreadSheetID = "1DfLJVyB58SuOBEAKeg9YYC4WNE3JYsw-J60PytvcwnU"
var SheetName = "Chemical Inventory"


function BacteriaInventory() {
  var ss = SpreadsheetApp.openById(SpreadSheetID);
  var inventory = ss.getSheetByName(SheetName);

  var bac_inv = getData(inventory);

  // to get time frame for reordering due to expiration
  const now = new Date();
  const MILLS_PER_DAY = 1000 * 60 * 60 * 24;
  var plus_month = new Date(now.getTime() + 30*MILLS_PER_DAY)

  restock = [];
  
  // adds only items that need to be restocked to restick object
  for(var i = 0, l= bac_inv.length; i<l ; i++){
    if (bac_inv[i]['Quantity in Stock'] <= bac_inv[i]['Reorder Level'] || bac_inv[i]['Earliest EXP Date'] < plus_month || !bac_inv[i]['Earliest EXP Date']){
      item = bac_inv[i]['Name'];
      size = bac_inv[i]['Weight/Volume'];
      num_left = bac_inv[i]['Quantity in Stock'];
      exp = bac_inv[i]['Earliest EXP Date'];

      if (!bac_inv[i]['Earliest EXP Date'] || bac_inv[i]['Earliest EXP Date'] == "NA"){
        restock.push(bac_inv[i]);
      }

      else{
        if (bac_inv[i]['Earliest EXP Date'] < plus_month){
          restock.push(bac_inv[i]);
        }

        else{
          restock.push(bac_inv[i]);
        }
      }
    }
  }

  // sends email to these people about items to be restocked
  // look at line 76 for formatting email
  MailApp.sendEmail({to: EMAIL1,
                subject: "Items to Restock",
                htmlBody: printStuff(restock),
                noReply:true});
  MailApp.sendEmail({to: EMAIL2,
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

// makes restock object a string so it can be emailed, also formats information into a table for easy interpretation
// https://stackoverflow.com/questions/58767859/how-to-send-a-table-in-an-email-in-google-scripts
function printStuff(restock){
  string = "<html><body><br><table border=1><tr><th>Item</th><th>Quantity</th><th>Expiration</th></tr></br>";
  for (var i=0; i<restock.length; i++){
    exp = restock[i]['Earliest EXP Date'];
    string = string + "<tr>";

    // just have to keep separating NA and blanks from everything else
    if (!exp || exp == "NA"){
      temp = `<td> ${restock[i]['Name']} </td><td> ${restock[i]['Quantity in Stock']} </td>`;
      string = string.concat(temp);
    }
    else{
      temp = `<td> ${restock[i]['Name']} </td><td> ${restock[i]['Quantity in Stock']}  </td><td> ${Utilities.formatDate(exp, 'America/New_York', 'MMMM dd, yyyy')}</td>`;
      string = string.concat(temp);
    }
    string = string + "</tr>";
  }
  string = string + "</table></body></html>";
  return string;
}
