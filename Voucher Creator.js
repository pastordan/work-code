function autoFillGoogleDocFromForm(e) {
    /* ########## BRING DATA IN FROM FORM ########## */
//
    //e.values is an array of form values
    //Begin with general information about who's being paid
    var requestDate = e.values[0].slice(0, 9);
    var completedBy = e.values[1];
    var vendorName = e.values[2];
    var vendorStreet = e.values[3];
    var vendorCity = e.values[4];
    var memo = e.values[5];

    /* ########## MILEAGE ########## */
    //Next, collect data on requests for mileage reimbursements
    var mileage = e.values[6];

    //Mileage 1
    var mileageDate1 = e.values[7],
        mileagePurpose1 = e.values[8],
        mileageClassCode1 = e.values[9].slice(0, 3),
        startingLocation1 = e.values[10],
        endingLocation1 = e.values[11],
        startingOdometer1 = e.values[12],
        endingOdometer1 = e.values[13],
        personalMiles1 = e.values[14],
        mileageNotes1 = e.values[15];

    //Mileage 2
    var mileageDate2 = e.values[17],
        mileagePurpose2 = e.values[18],
        mileageClassCode2 = e.values[19].slice(0, 3),
        startingLocation2 = e.values[20],
        endingLocation2 = e.values[21],
        startingOdometer2 = e.values[22],
        endingOdometer2 = e.values[23],
        personalMiles2 = e.values[24],
        mileageNotes2 = e.values[25];

    //Mileage 3
    var mileageDate3 = e.values[27],
        mileagePurpose3 = e.values[28],
        mileageClassCode3 = e.values[29].slice(0, 3),
        startingLocation3 = e.values[30],
        endingLocation3 = e.values[31],
        startingOdometer3 = e.values[32],
        endingOdometer3 = e.values[33],
        personalMiles3 = e.values[34],
        mileageNotes3 = e.values[35];

    //Mileage 4
    var mileageDate4 = e.values[37],
        mileagePurpose4 = e.values[38],
        mileageClassCode4 = e.values[39].slice(0, 3),
        startingLocation4 = e.values[40],
        endingLocation4 = e.values[41],
        startingOdometer4 = e.values[42],
        endingOdometer4 = e.values[43],
        personalMiles4 = e.values[44],
        mileageNotes4 = e.values[45];

    //Mileage 5
    var mileageDate5 = e.values[47],
        mileagePurpose5 = e.values[48],
        mileageClassCode5 = e.values[49].slice(0, 3),
        startingLocation5 = e.values[50],
        endingLocation5 = e.values[51],
        startingOdometer5 = e.values[52],
        endingOdometer5 = e.values[53],
        personalMiles5 = e.values[54],
        mileageNotes5 = e.values[55];


    /* ########## OTHER CHARGES ########## */
    //Data on non mileage-related reimbursement requests
    //Class Code = 3 digit departmental code,
    //sliced because we don't need identifying label
    //meetings, office, contracting, events, other = expense categories
    //broken down into line items

    //Request 1
    var classCode1 = e.values[57].slice(0, 3),
        meetings1 = e.values[58],
        office1 = e.values[59],
        contracting1 = e.values[60],
        events1 = e.values[61],
        other1 = e.values[62],
        date1 = e.values[63],
        purpose1 = e.values[64],
        amount1 = e.values[65],
        notes1 = e.values[66];

    //Request 2
    var classCode2 = e.values[68].slice(0, 3),
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
    var classCode3 = e.values[79].slice(0, 3),
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
    var classCode4 = e.values[90].slice(0, 3),
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
    var classCode5 = e.values[101].slice(0, 3),
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
    //links to any uploaded receipts
    var receipts = e.values[112];


    /* ########## WORK WITH DATA ########## */

    //Turn timestamp into date string
    if (requestDate.split("/")[0].length == 1) {
        var month = "0" + requestDate.split("/")[0];
    } else {
        month = requestDate.split("/")[0];
    }

    if (requestDate.split("/")[1].length == 1) {
        var day = "0" + requestDate.split("/")[1];
    } else {
        day = requestDate.split("/")[1];
    }

    var year = requestDate.split("/")[2];

    var formattedDate = year + "-" + month + "-" + day;


    //Mileage calculations based on information in the spreadsheet
    var totalMiles1 = +endingOdometer1 - +startingOdometer1 - +personalMiles1,
        totalMiles2 = +endingOdometer2 - +startingOdometer2 - +personalMiles2,
        totalMiles3 = +endingOdometer3 - +startingOdometer3 - +personalMiles3,
        totalMiles4 = +endingOdometer4 - +startingOdometer4 - +personalMiles4,
        totalMiles5 = +endingOdometer5 - +startingOdometer5 - +personalMiles5,
        totalTotalMiles = totalMiles1 +
        totalMiles2 +
        totalMiles3 +
        totalMiles4 +
        totalMiles5;

    //Warning if someone's math is off
    if (+totalMiles1 <= 0 ||
        +totalMiles2 <= 0 ||
        +totalMiles3 <= 0 ||
        +totalMiles4 <= 0 ||
        +totalMiles5 <= 0) {
        var mileageWarning = "Check mileage!";
    } else {
        mileageWarning = " ";
    }

    //Calculate reimbursement due based on IRS rate
    var mileageCost1 = totalMiles1 * 0.625,
        mileageCost2 = totalMiles2 * 0.625,
        mileageCost3 = totalMiles3 * 0.625,
        mileageCost4 = totalMiles4 * 0.625,
        mileageCost5 = totalMiles5 * 0.625;

    var mileageCostTotal = +totalTotalMiles * 0.625;

    //Non-mileage calculations
    //Oh hey, it's just a total amount
    var amountTotal = +amount1 + amount2 + amount3 + amount4 + amount5;

    //Process account names
    //Check for double entry by counting how many account codes are filled in
    //then split account codes & names for further use
    var arr1 = [meetings1, office1, contracting1, events1, other1].filter(n => n),
        count1 = 0,
        i = arr1.length;
    while (i--) {
        if (arr1[i].length > 0) count1++;
    }

    if (count1 > 1) {
        var accountWarning1 = "1";
        var account1 = "Error";
        var accounta = "Error";
    } else {
        accountWarning1 = "";
        account1 = arr1.toString().slice(0, 5);
        accounta = arr1.toString().slice(5);
    }

    var arr2 = [meetings2, office2, contracting2, events2, other2].filter(n => n),
        count2 = 0,
        i2 = arr2.length;
    while (i2--) {
        if (arr2[i2].length > 0) count2++;
    }

    if (count2 > 1) {
        var accountWarning2 = "2";
        var account2 = "Error";
        var accountb = "Error";
    } else {
        accountWarning2 = "";
        account2 = arr2.toString().slice(0, 5);
        accountb = arr2.toString().slice(5);
    }

    var arr3 = [meetings3, office3, contracting3, events3, other3].filter(n => n),
        count3 = 0,
        i3 = arr3.length;
    while (i3--) {
        if (arr3[i3].length > 0) count3++;
    }

    if (count3 > 1) {
        var accountWarning3 = "3";
        var account3 = "Error";
        var accountc = "Error";
    } else {
        vaccountWarning3 = "";
        account3 = arr3.toString().slice(0, 5);
        accountc = arr3.toString().slice(5);
    }

    var arr4 = [meetings4, office4, contracting4, events4, other4].filter(n => n),
        count4 = 0,
        i4 = arr4.length;
    while (i4--) {
        if (arr4[i4].length > 0) count4++;
    }

    if (count4 > 1) {
        var accountWarning4 = "4";
        var account4 = "Error";
        var accountd = "Error";
    } else {
        accountWarning4 = "";
        account4 = arr4.toString().slice(0, 5);
        accountd = arr4.toString().slice(5);
    }
    var arr5 = [meetings5, office5, contracting5, events5, other5].filter(n => n),
        count5 = 0,
        i5 = arr5.length;
    while (i5--) {
        if (arr5[i5].length > 0) count5++;
    }

    if (count5 > 1) {
        var accountWarning5 = "5";
        var account5 = "Error";
        var accounte = "Error";
    } else {
        accountWarning5 = "";
        account5 = arr5.toString().slice(0, 5);
        accounte = arr5.toString().slice(5);
    }

    //Put a heads-up on the sheet and in email if there are too many accounts
    var accountWarning = [accountWarning1,
        accountWarning2,
        accountWarning3,
        accountWarning4,
        accountWarning5
    ].filter(n => n);

    var accountWarningLead = "Check accounts for items: ";
    if (count1 >= 1 && classCode1.length < 3 ||
        count2 >= 1 && classCode2.length < 3 ||
        count3 >= 1 && classCode3.length < 3 ||
        count4 >= 1 && classCode4.length < 3 ||
        count5 >= 1 && classCode5.length < 3) {
        //Just in case somebody muffs the "Other" category
        accountWarningLead = "Check class codes and accounts for items: ";
    }

    //Check to see if this might be a duplicate payment to a vendor
    const ss = SpreadsheetApp.getActiveSpreadsheet(),
        s = ss.getSheetByName("Dupe Check"),
        lastRow = s.getLastRow(),
        almostLastRow = +lastRow - 1,
        allbutLastNames = s.getRange("F2:F" + almostLastRow).getValues(),
        lastName = s.getRange("F" + lastRow + ":F" + lastRow).getValue(),
        flatNames = allbutLastNames.flat();

    if (flatNames.includes(lastName)) {
        var dupeWarning = "Possible Duplicate";
    } else {
        dupeWarning = " ";
    }


    /* ########## CREATE VOUCHER ########## */

    //Make a copy of the template, name it, and file it
    var file = DriveApp.getFileById("1ZnU3TcAhjdKHwM5HNbvIzNA8iMcOivU1MbR58lVek0k");
    var folder = DriveApp.getFolderById("16sVtf1v5pTAZHtIBN7IV00trYpHERl07");

    var copy = file.makeCopy("EXP " + formattedDate + " " + vendorName, folder);

    // get ID and Url of new File
    var newFileID = copy.getId();
    var newFileUrl = copy.getUrl();

    //Once we've got the new file created, we need to open it as a document by using its ID
    var doc = SpreadsheetApp.openById(newFileID);

    //write to topmost spreadsheet
    var sheet = doc.getSheets()[0];

    //fill in lines 4-10
    //mostly who's getting paid
    sheet.getRange("e4").setValue(requestDate);
    sheet.getRange("b6").setValue(vendorName);
    sheet.getRange("b7").setValue(vendorStreet);
    sheet.getRange("b8").setValue(vendorCity);
    sheet.getRange("b9").setValue(memo);

    //fill in lines 13-17, if it's a mileage reimbursement
    if (mileage == "Mileage") {
        if (startingLocation1.length !== 0) {
            sheet.getRange("a13").setValue(mileageClassCode1 + "-7000");
            sheet.getRange("b13").setValue(mileageDate1);
            sheet.getRange("c13").setValue("Travel");
            sheet.getRange("d13").setValue("See mileage log");
            sheet.getRange("e13").setValue(mileageCost1);
        }

        if (startingLocation2.length !== 0) {
            sheet.getRange("a14").setValue(mileageClassCode2 + "-7000");
            sheet.getRange("b14").setValue(mileageDate2);
            sheet.getRange("c14").setValue("Travel");
            sheet.getRange("d14").setValue("See mileage log");
            sheet.getRange("e14").setValue(mileageCost2);
        }

        if (startingLocation3.length !== 0) {
            sheet.getRange("a15").setValue(mileageClassCode3 + "-7000");
            sheet.getRange("b15").setValue(mileageDate3);
            sheet.getRange("c15").setValue("Travel");
            sheet.getRange("d15").setValue("See mileage log");
            sheet.getRange("e15").setValue(mileageCost3);
        }

        if (startingLocation4.length !== 0) {
            sheet.getRange("a16").setValue(mileageClassCode1 + "-7000");
            sheet.getRange("b16").setValue(mileageDate4);
            sheet.getRange("c16").setValue("Travel");
            sheet.getRange("d16").setValue("See mileage log");
            sheet.getRange("e16").setValue(mileageCost4);
        }

        if (startingLocation5.length !== 0) {
            sheet.getRange("a17").setValue(mileageClassCode1 + "-7000");
            sheet.getRange("b17").setValue(mileageDate5);
            sheet.getRange("c17").setValue("Travel");
            sheet.getRange("d17").setValue("See mileage log");
            sheet.getRange("e17").setValue(mileageCost5);
        }


        /* ########## Create Mileage Log ########## */
        //Place in separate document
        var mileageFile = DriveApp.getFileById("1hU6bfWErwTwqn5gON16SiJLB2qXJ9oUUjsF_U_mFJpU");
        var mileageFolder = DriveApp.getFolderById("1zTyhhtxLM4_rLPQHud2TyIE_Up8qigqT");
        var mileageCopy = mileageFile.makeCopy("Mileage Log " + formattedDate + " " + vendorName, mileageFolder);

        //Get ID and Url of new File
        var newMileageFileID = mileageCopy.getId();
        var newMileageFileUrl = mileageCopy.getUrl();

        //Open the new document by using its ID
        var mileageDoc = SpreadsheetApp.openById(newMileageFileID);

        //Write to topmost spreadsheet
        var mileageSheet = mileageDoc.getSheets()[0];

        //Top of sheet
        //assumes vendor is employee
        mileageSheet.getRange("f4").setValue(requestDate);
        mileageSheet.getRange("b6").setValue(vendorName);
        mileageSheet.getRange("b7").setValue(vendorStreet);
        mileageSheet.getRange("b8").setValue(vendorCity);

        //Mileage Log 1
        //each log records travel information required by auditors
        if (mileageClassCode1.length !== 0) {
            mileageSheet.getRange("a13").setValue(mileageClassCode1 + "-7000");
        } else {
            mileageSheet.getRange("a13").setValue("");
        }
        mileageSheet.getRange("b13").setValue(mileageDate1);
        mileageSheet.getRange("c13").setValue(mileagePurpose1);
        mileageSheet.getRange("a15").setValue(startingLocation1);
        mileageSheet.getRange("d15").setValue(endingLocation1);
        mileageSheet.getRange("a17").setValue(startingOdometer1);
        mileageSheet.getRange("c17").setValue(endingOdometer1);
        mileageSheet.getRange("e17").setValue(personalMiles1);
        mileageSheet.getRange("f17").setValue(totalMiles1);

        //Mileage Log 2
        if (mileageClassCode2.length !== 0) {
            mileageSheet.getRange("a19").setValue(mileageClassCode2 + "-7000");
        } else {
            mileageSheet.getRange("a19").setValue("");
        }
        mileageSheet.getRange("b19").setValue(mileageDate2);
        mileageSheet.getRange("c19").setValue(mileagePurpose2);
        mileageSheet.getRange("a21").setValue(startingLocation2);
        mileageSheet.getRange("d21").setValue(endingLocation2);
        mileageSheet.getRange("a23").setValue(startingOdometer2);
        mileageSheet.getRange("c23").setValue(endingOdometer2);
        mileageSheet.getRange("e23").setValue(personalMiles2);
        mileageSheet.getRange("f23").setValue(totalMiles2);

        //Mileage Log 3
        if (mileageClassCode3.length !== 0) {
            mileageSheet.getRange("a25").setValue(mileageClassCode1 + "-7000");
        } else {
            mileageSheet.getRange("a25").setValue("");
        }
        mileageSheet.getRange("b25").setValue(mileageDate3);
        mileageSheet.getRange("c25").setValue(mileagePurpose3);
        mileageSheet.getRange("a27").setValue(startingLocation3);
        mileageSheet.getRange("d27").setValue(endingLocation3);
        mileageSheet.getRange("a29").setValue(startingOdometer3);
        mileageSheet.getRange("c29").setValue(endingOdometer3);
        mileageSheet.getRange("e29").setValue(personalMiles3);
        mileageSheet.getRange("f29").setValue(totalMiles3);

        //Mileage Log 4
        if (mileageClassCode4.length !== 0) {
            mileageSheet.getRange("a31").setValue(mileageClassCode4 + "-7000");
        } else {
            mileageSheet.getRange("a31").setValue("");
        }
        mileageSheet.getRange("b31").setValue(mileageDate4);
        mileageSheet.getRange("c31").setValue(mileagePurpose4);
        mileageSheet.getRange("a33").setValue(startingLocation4);
        mileageSheet.getRange("d33").setValue(endingLocation4);
        mileageSheet.getRange("a35").setValue(startingOdometer4);
        mileageSheet.getRange("c35").setValue(endingOdometer4);
        mileageSheet.getRange("e35").setValue(personalMiles4);
        mileageSheet.getRange("f35").setValue(totalMiles4);

        //Mileage Log 5
        if (mileageClassCode5.length !== 0) {
            mileageSheet.getRange("a37").setValue(mileageClassCode5 + "-7000");
        } else {
            mileageSheet.getRange("a37").setValue("");
        }
        mileageSheet.getRange("b37").setValue(mileageDate5);
        mileageSheet.getRange("c37").setValue(mileagePurpose5);
        mileageSheet.getRange("a39").setValue(startingLocation5);
        mileageSheet.getRange("d39").setValue(endingLocation5);
        mileageSheet.getRange("a41").setValue(startingOdometer5);
        mileageSheet.getRange("c41").setValue(endingOdometer5);
        mileageSheet.getRange("e41").setValue(personalMiles5);
        mileageSheet.getRange("f41").setValue(totalMiles5);

        //Mileage Log Notes + warning, if any
        mileageSheet.getRange("a43").setValue(
            mileageWarning + "\n" +
            mileageNotes1 + "\n" +
            mileageNotes2 + "\n" +
            mileageNotes3 + "\n" +
            mileageNotes4 + "\n" +
            mileageNotes5 + "\n"
        );

        //Total requested reimbursement
        mileageSheet.getRange("e45").setValue(
            totalTotalMiles + " @ $0.625 per mile" + "\n" +
            "Please see attached mileage log for more details"
        );

        //fill in details about non-mileage requests
    } else {
        if (classCode2.length !== 0) {
            sheet.getRange("a13").setValue(classCode1 + "-" + account1);
            sheet.getRange("b13").setValue(date1);
            sheet.getRange("c13").setValue(accounta);
            sheet.getRange("d13").setValue(purpose1);
            sheet.getRange("e13").setValue(amount1);
        }
        if (classCode2.length !== 0) {
            sheet.getRange("a14").setValue(classCode2 + "-" + account2);
            sheet.getRange("b14").setValue(date2);
            sheet.getRange("c14").setValue(accountb);
            sheet.getRange("d14").setValue(purpose2);
            sheet.getRange("e14").setValue(amount2);
        }
        if (classCode3.length !== 0) {
            sheet.getRange("a15").setValue(classCode3 + "-" + account3);
            sheet.getRange("b15").setValue(date3);
            sheet.getRange("c15").setValue(accountc);
            sheet.getRange("d15").setValue(purpose3);
            sheet.getRange("e15").setValue(amount3);
        }
        if (classCode4.length !== 0) {
            sheet.getRange("a16").setValue(classCode4 + "-" + account4);
            sheet.getRange("b16").setValue(date4);
            sheet.getRange("c16").setValue(accountd);
            sheet.getRange("d16").setValue(purpose4);
            sheet.getRange("e16").setValue(amount4);
        }
        if (classCode5.length !== 0) {
            sheet.getRange("a17").setValue(classCode5 + "-" + account5);
            sheet.getRange("b17").setValue(date5);
            sheet.getRange("c17").setValue(accounte);
            sheet.getRange("d17").setValue(purpose5);
            sheet.getRange("e17").setValue(amount5);
        }
    }

    //fill in lines 36-45 (notes area)
    if (mileage == "Mileage") {
        sheet.getRange("a36").setValue(mileageNotes1 + "\n" +
            mileageWarning + "\n" +
            mileageNotes1 + "\n" +
            mileageNotes2 + "\n" +
            mileageNotes3 + "\n" +
            mileageNotes4 + "\n" +
            mileageNotes5 + "\n");
    } else {
        sheet.getRange("a36").setValue(notes1 + "\n" +
                accountWarningLead +
                accountWarning.join(", ")) + "\n" +
            notes2 + "\n" +
            notes3 + "\n" +
            notes4 + "\n" +
            notes5;
    }

    //Who filled this voucher out?
    sheet.getRange("C47").setValue(completedBy);

    // Log vouchers for expense tracker
    const s3 = ss.getSheetByName("Voucher List");
    var logLink = "=HYPERLINK(\"" + newMileageFileUrl + "\",\"See mileage log\")";
    if (mileage == "Mileage") {
        s3.appendRow([mileageDate1, vendorName, mileageClassCode1 + "-7000", logLink, mileageCost1]);
    } else {
        s3.appendRow([date1, vendorName, classCode1 + "-" + account1 + " " + accounta, purpose1, amountTotal]);
    }

    //Send email about new payment request
    //allows requestors and reviewers to see copy of voucher,
    //along with warnings about potential problems
    var subject = "EXP " + formattedDate + " " + vendorName;
    if (receipts.length !== 0) {
        var content = "New voucher for review: " + newFileUrl + "\n" +
            " " + "Receipts: " + receipts + "\n" +
            " " + accountWarningLead " " + accountWarning.join(", ") + "\n" +
            " " + mileageWarning + " " + dupeWarning;
    } else {
        var content = "New voucher for review: " + newFileUrl + "\n" +
            " " + "Mileage Voucher: " + newMileageFileUrl + "\n" +
            " " + mileageWarning + " " + dupeWarning;
    }


    //email Recipients
  switch (completedBy)
    {
      case 0: "G"
        var completeMail = "g@xxxx.org";
        break;
      case 1: "K"
        var completeMail = "k@xxxx.org";
        break;
      case 2: "P"
       var completeMail = "p@xxxx.org";
       break;
      case 3: "J"
        var completeMail = "j@xxxx.org";
        break;
      case 4: "L"
        var completeMail = "l@xxxx.org";
        break;
      case 5: "B"
        var completeMail = "b@xxxx.org";
        break;
      case 6: "D"
        var completeMail = "d@xxxx.org";
        break;
    }

    //send email itself
    //   MailApp.sendEmail("completeMail", subject, content, {
    //     htmlBody: content,
    //     cc: completeMail
    //     });
}
