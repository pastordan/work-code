function autoFillGoogleDocFromForm(e) {
  /* ########## BRING DATA IN FROM FORM ########## */

  //e.values is an array of form values
  //Start with general information
  var requestDate = e.values[0].slice(0,9);
  var completedBy = e.values[1];
  var vName = e.values[2];
  var vStreet = e.values[3];
  var vCity = e.values[4];
  var memo = e.values[5];

  /* ########## MILEAGE ########## */
  var mileage = e.values[6];

  //Mileage 1
  var milesDate1 = e.values[7],
      milesPurpose1 = e.values[8],
      milesClass1 = e.values[9], // returns three digit department code for billing
      start1 = e.values[10],
      end1 = e.values[11],
      startMiles1 = e.values[12],
      endMiles1 = e.values[13],
      personalMiles1 = e.values[14],
      mileageNotes1 = e.values[15];
  
  //Mileage 2
  var milesDate2 = e.values[17],
      milesPurpose2 = e.values[18],
      milesClass2 = e.values[19],
      start2 = e.values[20],
      end2 = e.values[21],
      startMiles2 = e.values[22],
      endMiles2 = e.values[23],
      personalMiles2 = e.values[24],
      mileageNotes2 = e.values[25];
  
  //Mileage 3
  var milesDate3 = e.values[27],
      milesPurpose3 = e.values[28],
      milesClass3 = e.values[29],
      start3 = e.values[30],
      end3 = e.values[31],
      startMiles3 = e.values[32],
      endMiles3 = e.values[33],
      personalMiles3 = e.values[34],
      mileageNotes3 = e.values[35];

  //Mileage 4
  var milesDate4 = e.values[37],
      milesPurpose4 = e.values[38],
      milesClass4 = e.values[39],
      start4 = e.values[40],
      end4 = e.values[41],
      startMiles4 = e.values[42],
      endMiles4 = e.values[43],
      personalMiles4 = e.values[44],
      mileageNotes4 = e.values[45];

  //Mileage 5
  var milesDate5 = e.values[47],
      milesPurpose5 = e.values[48],
      milesClass5 = e.values[49],
      start5 = e.values[50],
      end5 = e.values[51],
      startMiles5 = e.values[52],
      endMiles5 = e.values[53],
      personalMiles5 = e.values[54],
      mileageNotes5 = e.values[55];


  /* ########## OTHER CHARGES ########## */

  //Request 1
  var cCode1 = e.values[57], // returns three digit department code for billing
      meetings1 = e.values[58], // meetings-other: expense categories broken down into line items
      office1 = e.values[59],
      contracting1 = e.values[60],
      events1 = e.values[61],
      other1 = e.values[62],
      date1 = e.values[63],
      purpose1 = e.values[64],
      amount1 = e.values[65],
      notes1 = e.values[66]; 

  //Request 2
  var cCode2 = e.values[68],
      meetings2 = e.values[69],
      office2 = e.values[70],
      contracting2 = e.values[71],
      events2 = e.values[72],
      other2 = e.values[73],
      date2 = e.values[74],
      purpose2 = e.values[75],
      amount2 = e.values[76],
      notes2 = e.values[77];

  //Request 3
  var cCode3 = e.values[79],
      meetings3 = e.values[80],
      office3 = e.values[81],
      contracting3 = e.values[82],
      events3 = e.values[83],
      other3 = e.values[84],
      date3 = e.values[85],
      purpose3 = e.values[86],
      amount3 = e.values[87],
      notes3 = e.values[88];

  //Request 4
  var cCode4 = e.values[90],
      meetings4 = e.values[91],
      office4 = e.values[92],
      contracting4 = e.values[93],
      events4 = e.values[94],
      other4 = e.values[95],
      date4 = e.values[96],
      purpose4 = e.values[97],
      amount4 = e.values[98],
      notes4 = e.values[99];

  //Request 5
  var cCode5 = e.values[101],
      meetings5 = e.values[102],
      office5 = e.values[103],
      contracting5 = e.values[104],
      events5 = e.values[105],
      other5 = e.values[106],
      date5 = e.values[107],
      purpose5 = e.values[108],
      amount5 = e.values[109],
      notes5 = e.values[110];

  //Receipts
  var receipts = e.values[112];


  /* ########## WORK WITH DATA ########## */

  //Date
  if (requestDate.split('/')[0].length == 1){
	    var month = '0' + requestDate.split('/')[0];
  } else {
	    var month = requestDate.split('/')[0];
  }

  if (requestDate.split('/')[1].length == 1){
	    var day = '0' + requestDate.split('/')[1];
  } else {
	    var day = requestDate.split('/')[1];
  }

  var year = requestDate.split('/')[2];

  var formattedDate = year + '-' + month + '-' + day;


  //Mileage calculations based on information in the spreadsheet
  var totalMiles1 = +endMiles1 - +startMiles1 - +personalMiles1,
      totalMiles2 = +endMiles2 - +startMiles2 - +personalMiles2,
      totalMiles3 = +endMiles3 - +startMiles3 - +personalMiles3,
      totalMiles4 = +endMiles4 - +startMiles4 - +personalMiles4,
      totalMiles5 = +endMiles5 - +startMiles5 - +personalMiles5,
      totalTotalMiles = +totalMiles1 +
                        +totalMiles2 +
                        +totalMiles3 +
                        +totalMiles4 +
                        +totalMiles5 ;

  if (+totalMiles1 <= 0 ||
      +totalMiles2 <= 0 ||
      +totalMiles3 <= 0 ||
      +totalMiles4 <= 0 ||
      +totalMiles5 <= 0){
        var mileageWarning = 'Check mileage';
      } else {
        var mileageWarning = ' ';
      }

  var milesCost1 = totalMiles1 * .625,
      milesCost2 = totalMiles2 * .625,
      milesCost3 = totalMiles3 * .625,
      milesCost4 = totalMiles4 * .625,
      milesCost5 = totalMiles5 * .625;

  var milesCostTotal = +totalTotalMiles * .625;
  var amountTotal = +amount1 + +amount2 + +amount3 + +amount4 + +amount5;

  //Process account names
  //Check for double entry, the slice the account info
  var arr1 = [meetings1, office1, contracting1, events1, other1].filter(n =>n), count1 = 0, i = arr1.length;
    while (i--) { if (arr[i].length > 0) count1++; }

    if (count1 > 1) {
      var accountWarning1 = '1 ';
      var account1 = 'Error';
      var accounta = 'Error'
    } else {
      var accountWarning1 = ''
      var account1 = arr1.toString().slice(0,5);
      var accounta = arr1.toString().slice(5);
    }

  var arr2 = [meetings2, office2, contracting2, events2, other2].filter(n =>n), count2 = 0, i = arr2.length;
    while (i--) { if (arr2[i].length > 0) count2++; }

    if (count2 > 1) {
      var accountWarning2 = '2 ';
      var account2 = 'Error';
      var accountb = 'Error'
    } else {
      var accountWarning2 = ''
      var account2 = arr2.toString().slice(0,5);
      var accountb = arr2.toString().slice(5);
    }

  var arr3 = [meetings3, office3, contracting3, events3, other3].filter(n =>n), count3 = 0, i = arr3.length;
    while (i--) { if (arr3[i].length > 0) count3++; }

    if (count3 > 1) {
      var accountWarning3 = '3 ';
      var account3 = 'Error';
      var accountc = 'Error'
    } else {
      var accountWarning3 = ''
      var account3 = arr3.toString().slice(0,5);
      var accountc = arr3.toString().slice(5);

  var arr4 = [meetings4, office4, contracting4, events4, other4].filter(n =>n), count4 = 0, i = arr4.length;
    while (i--) { if (arr4[i].length > 0) count4++; }

    if (count4 > 1) {
      var accountWarning4 = '4 ';
      var account4 = 'Error';
      var accountd = 'Error'
    } else {
      var accountWarning4 = ''
      var account4 = arr4.toString().slice(0,5);
      var accountd = arr4.toString().slice(5);

  var arr5 = [meetings5, office5, contracting5, events5, other5].filter(n =>n), count5 = 0, i = arr5.length;
    while (i--) { if (arr5[i].length > 0) count5++; }

    if (count5 > 1) {
      var accountWarning5 = '5 ';
      var account5 = 'Error';
      var accounte = 'Error'
    } else {
      var accountWarning5 = ''
      var account5 = arr5.toString().slice(0,5);
      var accounte = arr5.toString().slice(5);

  var accountWarning = [accountWarning1, accountWarning2, accountWarning3, accountWarning4, accountWarning5];
  
  //Check for possible duplicates
  const ss = SpreadsheetApp.getActiveSpreadsheet(),
        s = ss.getSheetByName('Dupe Check'),
        lastRow = s.getLastRow(),
        almostLastRow = +lastRow-1,
        allbutLastNames = s.getRange('F2:F' + almostLastRow).getValues(),
        lastName = s.getRange('F' + lastRow + ':F' + lastRow).getValue(),
        flatNames = allbutLastNames.flat();

        if(flatNames.includes(lastName)){
           var dupeWarning = 'POSSIBLE DUPLICATE';
        } else {
           var dupeWarning = ' ';
        }


  /* ########## CREATE VOUCHER ########## */

  //Make a copy of the template, name it, and file it
  var file = DriveApp.getFileById('googleFoo');

  if (cCode1 == '340' || mileage == 'Mileage') {
    var folder = DriveApp.getFolderById('googleBar'); 
  } else if (cCode1 == '350') {
    var folder = DriveApp.getFolderById('googleBat');
  };
  
  var copy = file.makeCopy('EXP ' + formattedDate + ' ' + vName, folder);

  // get ID and Url of new File
  var newFileID = copy.getId();
  var newFileUrl = copy.getUrl();

  //Once we've got the new file created, we need to open it as a document by using its ID
  var doc = SpreadsheetApp.openById(newFileID); 
  
  //write to topmost spreadsheet
  var sheet = doc.getSheets()[0];

  //fill in lines 4-10
    sheet.getRange('e4').setValue(requestDate);
    sheet.getRange('b6').setValue(vName);
    sheet.getRange('b7').setValue(vStreet);
    sheet.getRange('b8').setValue(vCity);
    sheet.getRange('b9').setValue(memo);

  //fill in lines 13-17
  if (mileage == 'Mileage'){
    if (start1.length !== 0){
    sheet.getRange('a13').setValue('340-7000');
    sheet.getRange('b13').setValue(milesDate1);
    sheet.getRange('c13').setValue('DHS Phase Two: Travel');
    sheet.getRange('d13').setValue('See mileage log');
    sheet.getRange('e13').setValue(milesCost1);}

    if (start2.length !== 0){
    sheet.getRange('a14').setValue('340-7000');
    sheet.getRange('b14').setValue(milesDate2);
    sheet.getRange('c14').setValue('DHS Phase Two: Travel');
    sheet.getRange('d14').setValue('See mileage log');
    sheet.getRange('e14').setValue(milesCost2);}

    if (start3.length !== 0){
    sheet.getRange('a15').setValue('340-7000');
    sheet.getRange('b15').setValue(milesDate3);
    sheet.getRange('c15').setValue('DHS Phase Two: Travel');
    sheet.getRange('d15').setValue('See mileage log');
    sheet.getRange('e15').setValue(milesCost3);}

    if (start4.length !== 0){
    sheet.getRange('a16').setValue('340-7000');
    sheet.getRange('b16').setValue(milesDate4);
    sheet.getRange('c16').setValue('DHS Phase Two: Travel');
    sheet.getRange('d16').setValue('See mileage log');
    sheet.getRange('e16').setValue(milesCost4);}

    if (start5.length !== 0){
    sheet.getRange('a17').setValue('340-7000');
    sheet.getRange('b17').setValue(milesDate5);
    sheet.getRange('c17').setValue('DHS Phase Two: Travel');
    sheet.getRange('d17').setValue('See mileage log');
    sheet.getRange('e17').setValue(milesCost5);}


  /* ########## Create Mileage Log ########## */
  //Place in separate document
    var mileageFile = DriveApp.getFileById('mileageFoo'); 
    var mileageFolder = DriveApp.getFolderById('mileageBar');
    var mileageCopy = mileageFile.makeCopy('Mileage Log ' + formattedDate + ' ' + vName, mileageFolder);

  //Get ID and Url of new File
    var newMileageFileID = mileageCopy.getId();
    var newMileageFileUrl = mileageCopy.getUrl();

  //Once we've got the new file created, we need to open it as a document by using its ID
    var mileageDoc = SpreadsheetApp.openById(newMileageFileID); 

  //write to topmost spreadsheet
    var mileageSheet = mileageDoc.getSheets()[0];

  //Top of sheet
    mileageSheet.getRange('f4').setValue(requestDate);
    mileageSheet.getRange('b6').setValue(vName);
    mileageSheet.getRange('b7').setValue(vStreet);
    mileageSheet.getRange('b8').setValue(vCity);

  //Mileage Log 1
    mileageSheet.getRange('a13').setValue('340-7000');
    mileageSheet.getRange('b13').setValue(milesDate1);
    mileageSheet.getRange('c13').setValue(milesPurpose1);
    mileageSheet.getRange('a15').setValue(start1);
    mileageSheet.getRange('d15').setValue(end1);
    mileageSheet.getRange('d17').setValue(startMiles1);
    mileageSheet.getRange('c17').setValue(endMiles1);
    mileageSheet.getRange('e17').setValue(personalMiles1);
    mileageSheet.getRange('f17').setValue(totalMiles1);

  //Mileage Log 2
    mileageSheet.getRange('a19').setValue('340-7000');
    mileageSheet.getRange('b19').setValue(milesDate2);
    mileageSheet.getRange('c19').setValue(milesPurpose2);
    mileageSheet.getRange('a21').setValue(start2);
    mileageSheet.getRange('d21').setValue(end2);
    mileageSheet.getRange('a23').setValue(startMiles2);
    mileageSheet.getRange('c23').setValue(endMiles2);
    mileageSheet.getRange('e23').setValue(personalMiles2);
    mileageSheet.getRange('f23').setValue(totalMiles2);

  //Mileage Log 3
    mileageSheet.getRange('a25').setValue('340-7000');
    mileageSheet.getRange('b25').setValue(milesDate3);
    mileageSheet.getRange('c25').setValue(milesPurpose3);
    mileageSheet.getRange('a27').setValue(start3);
    mileageSheet.getRange('d27').setValue(end3);
    mileageSheet.getRange('a29').setValue(startMiles3);
    mileageSheet.getRange('c29').setValue(endMiles3);
    mileageSheet.getRange('e29').setValue(personalMiles3);
    mileageSheet.getRange('f29').setValue(totalMiles3);

  //Mileage Log 4
    mileageSheet.getRange('a31').setValue('340-7000');
    mileageSheet.getRange('b31').setValue(milesDate4);
    mileageSheet.getRange('c31').setValue(milesPurpose4);
    mileageSheet.getRange('a33').setValue(start4);
    mileageSheet.getRange('d33').setValue(end4);
    mileageSheet.getRange('a35').setValue(startMiles4);
    mileageSheet.getRange('c35').setValue(endMiles4);
    mileageSheet.getRange('e35').setValue(personalMiles4);
    mileageSheet.getRange('f35').setValue(totalMiles4);

  //Mileage Log 5
    mileageSheet.getRange('a37').setValue('340-7000');
    mileageSheet.getRange('b37').setValue(milesDate5);
    mileageSheet.getRange('c37').setValue(milesPurpose5);
    mileageSheet.getRange('a39').setValue(start5);
    mileageSheet.getRange('d39').setValue(end5);
    mileageSheet.getRange('a41').setValue(startMiles5);
    mileageSheet.getRange('c41').setValue(endMiles5);
    mileageSheet.getRange('e41').setValue(personalMiles5);
    mileageSheet.getRange('f41').setValue(totalMiles5);

  //Mileage Log Notes
    mileageSheet.getRange('a43').setValue (
      milesNotes1 + '\n' +
      milesNotes2 + '\n' +
      milesNotes3 + '\n' +
      milesNotes4 + '\n' +
      milesNotes5
      );

  sheet.getRange('e45').setValue(
    totalTotalMiles + ' @ $.625 per mile' + '\n' +
    'Please see attached mileage log for more details'
    );

  //if not mileage
  } else {
    sheet.getRange('a13').setValue(cCode1+'-'+account1);
    sheet.getRange('b13').setValue(date1);
    sheet.getRange('c13').setValue(accounta);
    sheet.getRange('d13').setValue(purpose1);
    sheet.getRange('e13').setValue(amount1);

  if (cCode2.length !==0){
    sheet.getRange('a14').setValue(cCode2+'-'+account2);
    sheet.getRange('b14').setValue(date2);
    sheet.getRange('c14').setValue(accountb);
    sheet.getRange('d14').setValue(purpose2);
    sheet.getRange('e14').setValue(amount2);
  }
  if (cCode3.length !== 0){
    sheet.getRange('a15').setValue(cCode3+'-'+account3);
    sheet.getRange('b15').setValue(date3);
    sheet.getRange('c15').setValue(accountc);
    sheet.getRange('d15').setValue(purpose3);
    sheet.getRange('e15').setValue(amount3);
  }
  if (cCode4.length !== 0){
    sheet.getRange('a16').setValue(cCode4+'-'+account4);
    sheet.getRange('b16').setValue(date4);
    sheet.getRange('c16').setValue(accountd);
    sheet.getRange('d16').setValue(purpose4);
    sheet.getRange('e16').setValue(amount4);
  }
  if (cCode5.length !== 0){
    sheet.getRange('a17').setValue(cCode5+'-'+account5);
    sheet.getRange('b17').setValue(date5);
    sheet.getRange('c17').setValue(accounte);
    sheet.getRange('d17').setValue(purpose5);
    sheet.getRange('e17').setValue(amount5);
  }}

  //fill in lines 36-45
  if (mileage == 'Mileage'){
    sheet.getRange('a36').setValue(milesNotes1 + '\n' + 
                                   milesNotes2 + '\n' + 
                                   milesNotes3 + '\n' + 
                                   milesNotes4 + '\n' + 
                                   milesNotes5);
  } else {
    sheet.getRange('a36').setValue(notes1 + '\n' + notes2 + '\n' + notes3 + '\n' + notes4 + '\n' + notes5);
    sheet.getRange('e45').setValue(amountTotal);
  }

  sheet.getRange('C47').setValue(completedBy);

  // Log vouchers for expense tracker
  const s3 = ss.getSheetByName('Voucher List');
  var logLink = '=HYPERLINK(\"' + newMileageFileUrl + '\",\"See mileage log\")';
  if (mileage == 'Mileage'){
      s3.appendRow([milesDate1, vName, '340-7000', logLink, milesCost1]);
    } else {
      s3.appendRow([date1, vName, cCode1 + '-' + account1 + ' ' + accounta, purpose1, amountTotal]); 
    }

  
  //Send email about new payment request
    var subject = 'EXP ' + formattedDate + ' ' + vName;
    if (receipts.length !== 0){
        var content = 'New voucher for review: ' + newFileUrl + '\n' +
                      ' ' + 'Receipts: ' + receipts + '\n' +
                      ' ' + 'Check accounts for items: ' + accountWarning + '\n',
                      ' ' + mileageWarning + ' ' + dupeWarning;
    } else {
        var content = 'New voucher for review: ' + newFileUrl + '\n' +
                       ' ' + 'Mileage Voucher: ' + newMileageFileUrl + '\n' +
                       ' ' + mileageWarning + ' ' + dupeWarning;
    }
    
  
  //Recipients
  switch (completedBy)
    {
      case 0: 'G'
        var completeMail = 'g@xxxx.org';
        break;
      case 1: 'K'
        var completeMail = 'k@xxxx.org';
        break;
      case 2: 'P'
       var completeMail = 'p@xxxx.org';
       break;
      case 3: 'J'
        var completeMail = 'j@xxxx.org';
        break;
      case 4: 'L'
        var completeMail = 'l@xxxx.org';
        break;
      case 5: 'B'
        var completeMail = 'b@xxxx.org';
        break;
      case 6: 'D'
        var completeMail = 'd@xxxx.org';
        break;
    }

 //email itself
 //   MailApp.sendEmail('completeMail', subject, content, {
 //     htmlBody: content,
 //     cc: completeMail
 //     });     
}