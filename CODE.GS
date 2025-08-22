// 26 oct
function doGet() {
  var htmlEvaluate = HtmlService.createTemplateFromFile("index")
  return htmlEvaluate.evaluate()
    .setTitle('Web App')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME).addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setRetDecValueInVahicleSh(obj) {
  try {
    obj = JSON.parse(obj);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var verhicle_sh = ss.getSheetByName('VEHCILE NUMBER');
    var branches_arr = obj.branches_arr;

    var dataValidation_sh_dt = ss.getSheetByName('DATAVALIDATION').getRange("X2:AB").getValues().filter(function (rw) {
      return rw[0] != "" && (rw[1] != "" || rw[3] != "") && (rw[2].toString().trim() === "YES" || rw[4].toString().trim() === "YES") && rw[0].toString().trim() == branches_arr.toString().trim();
    });

    var verhicle_dt = verhicle_sh.getDataRange().getValues();
    var wt_content = ss.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();
    var mas_apiKey = wt_content[17][1];
    var mas_username = wt_content[15][1];
    var mas_password = wt_content[16][1];

    var vehicle_num = obj.vehicle_num;
    var veri_ret_dec = obj.veri_ret_dec ? obj.veri_ret_dec : "Not available";
    var verify_perc_tds = obj.verify_perc_tds ? obj.verify_perc_tds : "Not available";
    var comment_verified = obj.comment_verified;

    var number = [];
    var content = "*Vehicle number:* " + vehicle_num + "\n*RET Dec:* " + veri_ret_dec + "\n% *TDS:* " + verify_perc_tds + "\n*Comment:* " + comment_verified;

    for (var i = 1; i < verhicle_dt.length; i++) {
      if (verhicle_dt[i][0].toString().trim() === vehicle_num.toString().trim()) {

        verhicle_sh.getRange('N' + (i + 1)).setValue('');
        verhicle_sh.getRange('O' + (i + 1)).setValue('');
        verhicle_sh.getRange('P' + (i + 1)).setValue('');

        if (dataValidation_sh_dt.length > 0) {
          if (dataValidation_sh_dt[0][2].toString().trim() === "YES") {
            number.push(dataValidation_sh_dt[0][1]);
          }
          if (dataValidation_sh_dt[0][4].toString().trim() === "YES") {
            number.push(dataValidation_sh_dt[0][3]);
          }
        }

        if (number.length > 0) {
          sendMessage(number, content, [], mas_apiKey, mas_username, mas_password);
        }

        return JSON.stringify({
          status: "success"
        });
      }
    }

    if (i === verhicle_dt.length) {
      return JSON.stringify({
        status: "fail"
      });
    }
  }
  catch (err) {
    return JSON.stringify({
      status: "fail"
    });
  }
}

function setRCValueInVahicleSh(obj) {
  try {
    obj = JSON.parse(obj);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var verhicle_sh = ss.getSheetByName('VEHCILE NUMBER');
    var dataValidation_sh_dt = ss.getSheetByName('DATAVALIDATION').getRange("X2:AB").getValues().filter(function (rw) {
      return rw[0] != "" && (rw[1] != "" || rw[3] != "") && (rw[2].toString().trim() === "YES" || rw[4].toString().trim() === "YES");
    });
    var verhicle_dt = verhicle_sh.getDataRange().getValues();
    var wt_content = ss.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();
    var mas_apiKey = wt_content[17][1];
    var mas_username = wt_content[15][1];
    var mas_password = wt_content[16][1];

    var vehicle_num = obj.vehicle_num;

    var veri_rc_copy = obj.veri_rc_copy ? obj.veri_rc_copy : "Not available";
    var veri_address = obj.veri_address ? obj.veri_address : "Not available";

    var comment_verified = obj.comment_verified;
    var branches_arr = obj.branches_arr;
    var number = [];
    var content = "*Vehicle number:* " + vehicle_num + "\n*RC Copy:* " + veri_rc_copy + "\n% *Address:* " + veri_address + "\n*Comment:* " + comment_verified;

    for (var i = 1; i < verhicle_dt.length; i++) {
      if (verhicle_dt[i][0].toString().trim() === vehicle_num.toString().trim()) {

        verhicle_sh.getRange('L' + (i + 1)).setValue('');
        verhicle_sh.getRange('M' + (i + 1)).setValue('');

        if (dataValidation_sh_dt.length > 0) {
          for (var k = 0; k < dataValidation_sh_dt.length; k++) {
            if (branches_arr.includes(dataValidation_sh_dt[k][0].toString().trim())) {
              if (dataValidation_sh_dt[k][2].toString().trim() === "YES") {
                number.push(dataValidation_sh_dt[k][1]);
              }
              if (dataValidation_sh_dt[k][4].toString().trim() === "YES") {
                number.push(dataValidation_sh_dt[k][3]);
              }
            }
          }
        }

        if (number.length > 0) {
          sendMessage(number, content, [], mas_apiKey, mas_username, mas_password);
        }


        return JSON.stringify({
          status: "success"
        });
      }
    }

    if (i === verhicle_dt.length) {
      return JSON.stringify({
        status: "fail"
      });
    }
  }
  catch (err) {
    return JSON.stringify({
      status: "fail"
    });
  }
}

function setTdsValueInVahicleSh(obj) {
  try {
    obj = JSON.parse(obj);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var verhicle_sh = ss.getSheetByName('VEHCILE NUMBER');
    var dataValidation_sh_dt = ss.getSheetByName('DATAVALIDATION').getRange("X2:AB").getValues().filter(function (rw) {
      return rw[0] != "" && (rw[1] != "" || rw[3] != "") && (rw[2].toString().trim() === "YES" || rw[4].toString().trim() === "YES");
    });
    var verhicle_dt = verhicle_sh.getDataRange().getValues();
    var wt_content = ss.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();
    var mas_apiKey = wt_content[17][1];
    var mas_username = wt_content[15][1];
    var mas_password = wt_content[16][1];

    var vehicle_num = obj.vehicle_num;
    var veri_tds_form = obj.veri_tds_form ? obj.veri_tds_form : "Not available";
    var veri_tds = obj.veri_tds ? obj.veri_tds : "Not available";
    var comment_verified = obj.comment_verified;
    var branches_arr = obj.branches_arr;
    var number = [];
    var content = "*Vehicle number:* " + vehicle_num + "\n*TDS:* " + veri_tds + "\n% *TDS FORM:* " + veri_tds_form + "\n*Comment:* " + comment_verified;

    for (var i = 1; i < verhicle_dt.length; i++) {
      if (verhicle_dt[i][0].toString().trim() === vehicle_num.toString().trim()) {

        verhicle_sh.getRange('J' + (i + 1)).setValue('');
        verhicle_sh.getRange('K' + (i + 1)).setValue('');

        if (dataValidation_sh_dt.length > 0) {
          for (var k = 0; k < dataValidation_sh_dt.length; k++) {
            if (branches_arr.includes(dataValidation_sh_dt[k][0].toString().trim())) {
              if (dataValidation_sh_dt[k][2].toString().trim() === "YES") {
                number.push(dataValidation_sh_dt[k][1]);
              }
              if (dataValidation_sh_dt[k][4].toString().trim() === "YES") {
                number.push(dataValidation_sh_dt[k][3]);
              }
            }
          }
        }

        if (number.length > 0) {
          sendMessage(number, content, [], mas_apiKey, mas_username, mas_password);
        }

        return JSON.stringify({
          status: "success"
        });
      }
    }

    if (i === verhicle_dt.length) {
      return JSON.stringify({
        status: "fail"
      });
    }
  }
  catch (err) {
    return JSON.stringify({
      status: "fail"
    });
  }
}

function setPanValueInVahicleSh(obj) {
  // try{
  obj = JSON.parse(obj);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var verhicle_sh = ss.getSheetByName('VEHCILE NUMBER');
  var dataValidation_sh_dt = ss.getSheetByName('DATAVALIDATION').getRange("X2:AB").getValues().filter(function (rw) {
    return rw[0] != "" && (rw[1] != "" || rw[3] != "") && (rw[2].toString().trim() === "YES" || rw[4].toString().trim() === "YES");
  });
  var verhicle_dt = verhicle_sh.getDataRange().getValues();
  var wt_content = ss.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();
  var mas_apiKey = wt_content[17][1];
  var mas_username = wt_content[15][1];
  var mas_password = wt_content[16][1];

  var vehicle_num = obj.vehicle_num;

  var veri_pan_img = obj.veri_pan_img ? obj.veri_pan_img : "Not available";
  var veri_name_on_pan = obj.veri_name_on_pan ? obj.veri_name_on_pan : "Not available";
  var veri_pan_no = obj.veri_pan_no ? obj.veri_pan_no : "Not available";

  var comment_verified = obj.comment_verified;
  var branches_arr = obj.branches_arr;
  var number = [];
  var content = "*Vehicle number:* " + vehicle_num + "\n*PAN NUMBER:* " + veri_pan_no + "\n*Name on PAN:* " + veri_name_on_pan + "\n% *PAN Image:* " + veri_pan_img + "\n*Comment:* " + comment_verified;

  for (var i = 1; i < verhicle_dt.length; i++) {
    if (verhicle_dt[i][0].toString().trim() === vehicle_num.toString().trim()) {

      verhicle_sh.getRange('G' + (i + 1)).setValue('');
      verhicle_sh.getRange('H' + (i + 1)).setValue('');
      verhicle_sh.getRange('I' + (i + 1)).setValue('');

      if (dataValidation_sh_dt.length > 0) {
        for (var k = 0; k < dataValidation_sh_dt.length; k++) {
          if (branches_arr.includes(dataValidation_sh_dt[k][0].toString().trim())) {
            if (dataValidation_sh_dt[k][2].toString().trim() === "YES") {
              number.push(dataValidation_sh_dt[k][1]);
            }
            if (dataValidation_sh_dt[k][4].toString().trim() === "YES") {
              number.push(dataValidation_sh_dt[k][3]);
            }
          }
        }
      }

      if (number.length > 0) {
        sendMessage(number, content, [], mas_apiKey, mas_username, mas_password);
      }

      return JSON.stringify({
        status: "success"
      });
    }
  }

  if (i === verhicle_dt.length) {
    return JSON.stringify({
      status: "fail"
    });
  }
  // }
  // catch(err){
  //   return JSON.stringify({
  //     status: "fail"
  //   });
  // }
}

function setVehiclesNumbers(obj) {
  try {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var vehicle_sh = ss.getSheetByName("VEHCILE NUMBER");
    var vehicle_data = vehicle_sh.getDataRange().getValues();

    var data_ = [];
    vehicle_data.forEach(function (row) {
      data_.push(row[0].toString().trim());
    });

    var obj = JSON.parse(obj);
    var vehicles = obj['vehicles'];
    var main_vehicle = obj['main_vehicle'];

    for (var i = 1; i < vehicle_data.length; i++) {
      var vehicleData = vehicle_data[i][0].toString().trim();
      if (vehicleData === main_vehicle.toString().trim()) {
        for (var j = 0; j < vehicles.length; j++) {
          if (!data_.includes(vehicles[j].toString().trim())) {
            vehicle_sh.appendRow([
              vehicles[j], "", "", "", "", "",
              vehicle_data[i][6], vehicle_data[i][7], vehicle_data[i][8], vehicle_data[i][9], vehicle_data[i][10]
            ])
          }

        }
        return JSON.stringify({
          status: "success"
        });
      }
    }

    if (i === vehicle_data.length) {
      return JSON.stringify({
        status: "fail"
      });
    }
  }
  catch (err) {
    return JSON.stringify({
      status: "fail"
    });
  }
}

function setVehicleNumberData(pan_no) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_ = ss.getSheetByName("VEHCILE NUMBER");
  var sheet_dt = sheet_.getDataRange().getValues();
  sheet_dt.forEach(function (row, ii) {
    if (row[6].toString().trim().toUpperCase() == pan_no.toString().trim().toUpperCase()) {
      sheet_.getRange(ii + 1, 10).setValue("YES");
      sheet_.getRange(ii + 1, 14, 1, 2).setValues([["", ""]]);
      sheet_.getRange(ii + 1, 18).setValue("");
    }
  });

  SpreadsheetApp.flush();
  var vehicle_no_dt = ss.getSheetByName("VEHCILE NUMBER").getDataRange().getValues();
  return JSON.stringify({
    'vehicle_no_dt': vehicle_no_dt
  });

}

function setAcknlmntData(data) {
  data = JSON.parse(data);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var folder_url = spreadsheet.getSheetByName("WHATSAPP CONTENT").getRange('G2').getValue();
  var wt_content = spreadsheet.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();

  var data_validation = spreadsheet.getSheetByName("DATAVALIDATION").getRange("X2:Z").getValues().filter(function (row) {
    return row[0] != "";
  });

  var upload_company_ack_arr = data[0], upload_back_of_lr_ack_ar = data[1], upload_front_of_lr_Ack_ar = data[2];
  var wt_RESPONSES = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();

  if (upload_company_ack_arr.length > 0) {
    var blob = Utilities.newBlob(Utilities.base64Decode(upload_company_ack_arr[0].data), upload_company_ack_arr[0].mimeType, upload_company_ack_arr[0].fileName);
    var fld = DriveApp.getFolderById(getIdFromUrl(folder_url)[0]);
    var upload_company_ack_arr_file = fld.createFile(blob).getUrl();
  }
  else {
    var upload_company_ack_arr_file = "";
  }

  if (upload_back_of_lr_ack_ar.length > 0) {
    var blob = Utilities.newBlob(Utilities.base64Decode(upload_back_of_lr_ack_ar[0].data), upload_back_of_lr_ack_ar[0].mimeType, upload_back_of_lr_ack_ar[0].fileName);
    var fld = DriveApp.getFolderById(getIdFromUrl(folder_url)[0]);
    var upload_back_of_lr_ack_ar_file = fld.createFile(blob).getUrl();
  }
  else {
    var upload_back_of_lr_ack_ar_file = "";
  }

  if (upload_front_of_lr_Ack_ar.length > 0) {
    var blob = Utilities.newBlob(Utilities.base64Decode(upload_front_of_lr_Ack_ar[0].data), upload_front_of_lr_Ack_ar[0].mimeType, upload_front_of_lr_Ack_ar[0].fileName);
    var fld = DriveApp.getFolderById(getIdFromUrl(folder_url)[0]);
    var upload_front_of_lr_Ack_ar_file = fld.createFile(blob).getUrl();
  }
  else {
    var upload_front_of_lr_Ack_ar_file = "";
  }

  var welcom_ = data[3];
  var ind = data[4];

  for (var i = 0; i < wt_RESPONSES.length; i++) {
    if (wt_RESPONSES[i][3] == ind) {
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 51, 1, 5).setValues([[new Date(), welcom_, upload_front_of_lr_Ack_ar_file, upload_back_of_lr_ack_ar_file, upload_company_ack_arr_file]]);

      var data_ = data_validation.filter(function (rw_) {
        return rw_[0] == wt_RESPONSES[i][17]
      });

      if (data_.length > 0) {
        var number = data_[0][2];
        var content = wt_content[30][1];
        var mas_apiKey = wt_content[17][1];
        var mas_username = wt_content[15][1];
        var mas_password = wt_content[16][1];

        content = content.replace(/{LR NO}/, data_[0][41]).replace(/{LR DATE}/, Utilities.formatDate(new Date(data_[0][42]), 'GMT+05:30', 'dd-MM-yyyy')).replace(/{VEHICLE NO}/, data_[0][36]).replace(/{CLIENT}/, data_[0][4]).replace(/{SOURCE}/, data_[0][5]).replace(/{DESTINATION}/, data_[0][6]);

        sendMessage(number, content, [], mas_apiKey, mas_username, mas_password);
      }


      break;
    }
  }

  SpreadsheetApp.flush();
  var wt_RESPONSES1 = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  return JSON.stringify({
    'responses': wt_RESPONSES1
  });

}

function setAcknlmntVerifyData(data) {
  data = JSON.parse(data);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var wt_RESPONSES = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  var wt_content = spreadsheet.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();
  var data_validation = spreadsheet.getSheetByName("DATAVALIDATION").getRange("X2:Z").getValues().filter(function (row) {
    return row[0] != "";
  });

  var folder_url = spreadsheet.getSheetByName("WHATSAPP CONTENT").getRange('G2').getValue();

  let ind = data[0], somoa = data[1], case_shortage = data[2], unit_shortage = data[3], arrival_date_time = data[4], release_date_time = data[5], welcom_ = data[6], case_excess = data[7], unit_excess = data[8], case_damage = data[9], unit_damage = data[10], labour_amount = data[11];

  var upload_company_ack_arr = data[12];

  if (upload_company_ack_arr.length > 0) {
    var blob = Utilities.newBlob(Utilities.base64Decode(upload_company_ack_arr[0].data), upload_company_ack_arr[0].mimeType, upload_company_ack_arr[0].fileName);
    var fld = DriveApp.getFolderById(getIdFromUrl(folder_url)[0]);
    var upload_company_ack_arr_file = fld.createFile(blob).getUrl();
  }
  else {
    var upload_company_ack_arr_file = "";
  }

  var upload_back_of_lr_ack_arr = data[13];

  if (upload_back_of_lr_ack_arr.length > 0) {
    var blob = Utilities.newBlob(Utilities.base64Decode(upload_back_of_lr_ack_arr[0].data), upload_back_of_lr_ack_arr[0].mimeType, upload_back_of_lr_ack_arr[0].fileName);
    var fld = DriveApp.getFolderById(getIdFromUrl(folder_url)[0]);
    var upload_back_of_lr_ack_ar_file = fld.createFile(blob).getUrl();
  }
  else {
    var upload_back_of_lr_ack_ar_file = "";
  }

  var upload_front_of_lr_Ack_arr = data[14];

  if (upload_front_of_lr_Ack_arr.length > 0) {
    var blob = Utilities.newBlob(Utilities.base64Decode(upload_front_of_lr_Ack_arr[0].data), upload_front_of_lr_Ack_arr[0].mimeType, upload_front_of_lr_Ack_arr[0].fileName);
    var fld = DriveApp.getFolderById(getIdFromUrl(folder_url)[0]);
    var upload_front_of_lr_Ack_ar_file = fld.createFile(blob).getUrl();
  }
  else {
    var upload_front_of_lr_Ack_ar_file = "";
  }



  for (var i = 0; i < wt_RESPONSES.length; i++) {
    if (wt_RESPONSES[i][3] == ind) {
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 53, 1, 3).setValues([[upload_company_ack_arr_file, upload_back_of_lr_ack_ar_file, upload_front_of_lr_Ack_ar_file]]);

      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 57, 1, 13).setValues([[new Date(), welcom_, arrival_date_time, release_date_time, case_shortage, unit_shortage, case_excess, unit_excess, case_damage, unit_damage, labour_amount, somoa, data[15]]]);

      var data_ = data_validation.filter(function (rw_) {
        return rw_[0].toString().trim() == wt_RESPONSES[i][17].toString().trim();
      });

      if (data_.length > 0 && (case_shortage > 0 || unit_shortage > 0 || case_excess > 0 || unit_excess > 0 || case_damage > 0 || unit_damage > 0 || labour_amount > 0)) {
        var number = data_[0][2];
        var content = wt_content[40][1];
        var mas_apiKey = wt_content[17][1];
        var mas_username = wt_content[15][1];
        var mas_password = wt_content[16][1];

        var date1 = new Date(data_[0][41]);
        var date2 = new Date();
        var diffTime = Math.abs(date2 - date1);
        var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        if (diffDays <= 60) {
          content = content.replace(/{REASON}/, "Acknowledgement sucessfully submitted with in 60 days");
        }
        else {
          content = content.replace(/{REASON}/, "Acknowledgement sucessfully but since submitted after 60 days Balance is not Payable");
        }

        content = content.replace(/{LR NO}/, data_[0][41]).replace(/{LR DATE}/, Utilities.formatDate(new Date(data_[0][42]), 'GMT+05:30', 'dd-MM-yyyy')).replace(/{Case Shortage}/, case_shortage).replace(/{Unit Shortage}/, unit_shortage).replace(/{VEHICLE NO}/, data_[0][36]).replace(/{CLIENT}/, data_[0][4]).replace(/{SOURCE}/, data_[0][5]).replace(/{DESTINATION}/, data_[0][6]).replace(/{Case Excess}/, case_excess).replace(/{Unit Excess}/, unit_excess).replace(/{Case Damage}/, case_damage).replace(/{Unit Damage}/, unit_damage).replace(/{Labour Amount}/, labour_amount);

        sendMessage(number, content, [], mas_apiKey, mas_username, mas_password);
      }


      break;
    }
  }

  SpreadsheetApp.flush();
  var wt_RESPONSES1 = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  return JSON.stringify({
    'responses': wt_RESPONSES1
  });
}

function submitVerifyData(data) {
  data = JSON.parse(data);
  var vehicle_no = data[0], veri_pan_no = data[1], veri_name_on_pan = data[3], veri_address = data[7];
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var vehicle_sh = spreadsheet.getSheetByName("VEHCILE NUMBER")
  var vehicle_no_dt = vehicle_sh.getDataRange().getValues();

  for (var i = 1; i < vehicle_no_dt.length; i++) {
    if (vehicle_no_dt[i][0] != "" && vehicle_no_dt[i][0].toString().trim() == vehicle_no.toString().trim()) {
      vehicle_sh.getRange(i + 1, 7).setValue(veri_pan_no);
      vehicle_sh.getRange(i + 1, 9).setValue(veri_name_on_pan);
      vehicle_sh.getRange(i + 1, 13).setValue(veri_address);
      vehicle_sh.getRange(i + 1, 18).setValue(new Date());
      break;
    }
  }

  SpreadsheetApp.flush();
  var vehicle_no_dt_ = vehicle_sh.getDataRange().getValues()
  return JSON.stringify({
    'vehicle_no_dt': vehicle_no_dt_
  });
}


function saveMarketEmptyMovement(formDataVal, ActionType) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName(ActionType);
  var idVal = [];
  var currentId = "0000";
  var maxNumber = 1;
  sourceSheet.getDataRange().getValues().forEach(function (item, i) {
    if (i > 0 && item[0] != "") {
      idVal.push(Number(item[0].split('-')[1]));
    }
  });
  if (idVal.length == 0) {
    if (ActionType == "Empty Movement") {
      currentId = "EMP-0001"
    } else {
      currentId = "MKT-0001"
    }
  } else {
    maxNumber = Math.max(...idVal);
    maxNumber++;
    var str = String(maxNumber);
    while (str.length < 4) str = "0" + str;
    currentId = str;
    if (ActionType == "Empty Movement") {
      currentId = "EMP-" + currentId;
    } else {
      currentId = "MKT-" + currentId;
    }
  }
  formDataVal[0] = currentId;
  formDataVal.push(new Date())
  sourceSheet.appendRow(formDataVal);
}



function setPanData(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var vehicle_sh = spreadsheet.getSheetByName("VEHCILE NUMBER")
  var vehicle_no_dt = vehicle_sh.getDataRange().getValues();
  data = JSON.parse(data);
  var pan_no_ = data[0];
  var name_on_pan = data[1];
  var files = data[2];
  var vehicle_no_ = data[3];
  var img_url = data[4];
  var tm_rec_dec = data[5];
  var tm_rec_dec_ = data[6];

  if (img_url) {
    var file = img_url
  }
  else {
    var folder_url = spreadsheet.getSheetByName("WHATSAPP CONTENT").getRange('G2').getValue();
    var blob = Utilities.newBlob(Utilities.base64Decode(files[0].data), files[0].mimeType, files[0].fileName);
    var fld = DriveApp.getFolderById(getIdFromUrl(folder_url)[0]);
    var file = fld.createFile(blob).getUrl();
  }


  for (var i = 0; i < vehicle_no_dt.length; i++) {
    if (vehicle_no_dt[i][0] == vehicle_no_) {
      vehicle_sh.getRange(i + 1, 7, 1, 3).setValues([[pan_no_, file, name_on_pan]]);
      vehicle_sh.getRange(i + 1, 14, 1, 2).setValues([[tm_rec_dec, tm_rec_dec_]]);
      break;
    }
  }
  SpreadsheetApp.flush();
  var vehicle_no_dt_ = vehicle_sh.getDataRange().getValues();

  var c = 0;

  if (vehicle_no_dt_[i][6] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][9] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][11] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][13] != "") {
    c++;
  }

  if (c == 4) {
    vehicle_sh.getRange(i + 1, 6).setValue(new Date());
  }

  return JSON.stringify({
    'vehicle_no_dt': vehicle_no_dt_
  });
}

function setTdsData(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var vehicle_sh = spreadsheet.getSheetByName("VEHCILE NUMBER")
  var vehicle_no_dt = vehicle_sh.getDataRange().getValues();
  data = JSON.parse(data);
  var tds = data[0];
  var files = data[1];
  var vehicle_no_ = data[2];

  var folder_url = spreadsheet.getSheetByName("WHATSAPP CONTENT").getRange('G2').getValue();

  if (files.length > 0) {
    var blob = Utilities.newBlob(Utilities.base64Decode(files[0].data), files[0].mimeType, files[0].fileName);
    var fld = DriveApp.getFolderById(getIdFromUrl(folder_url)[0]);
    var file = fld.createFile(blob).getUrl();
  }
  else {
    var file = "";
  }

  for (var i = 0; i < vehicle_no_dt.length; i++) {
    if (vehicle_no_dt[i][0] == vehicle_no_) {
      vehicle_sh.getRange(i + 1, 10, 1, 2).setValues([[tds, file]]);
      break;
    }
  }
  SpreadsheetApp.flush();
  var vehicle_no_dt_ = vehicle_sh.getDataRange().getValues();

  var c = 0;

  if (vehicle_no_dt_[i][6] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][9] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][11] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][13] != "") {
    c++;
  }

  if (c == 4) {
    vehicle_sh.getRange(i + 1, 6).setValue(new Date());
  }

  return JSON.stringify({
    'vehicle_no_dt': vehicle_no_dt_
  });
}

function setRecDecData(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var vehicle_sh = spreadsheet.getSheetByName("VEHCILE NUMBER")
  var vehicle_no_dt = vehicle_sh.getDataRange().getValues();
  data = JSON.parse(data);
  var RET_DEC_AVAI = data[0];
  var tds = data[1];
  var files = data[2];
  var vehicle_no_ = data[3];

  var folder_url = spreadsheet.getSheetByName("WHATSAPP CONTENT").getRange('G2').getValue();

  if (files.length > 0) {
    var blob = Utilities.newBlob(Utilities.base64Decode(files[0].data), files[0].mimeType, files[0].fileName);
    var fld = DriveApp.getFolderById(getIdFromUrl(folder_url)[0]);
    var file = fld.createFile(blob).getUrl();
  }
  else {
    var file = "";
  }

  for (var i = 0; i < vehicle_no_dt.length; i++) {
    if (vehicle_no_dt[i][0] == vehicle_no_) {
      vehicle_sh.getRange(i + 1, 14, 1, 3).setValues([[RET_DEC_AVAI, file, tds]]);
      break;
    }
  }

  SpreadsheetApp.flush();
  var vehicle_no_dt_ = vehicle_sh.getDataRange().getValues();
  var c = 0;

  if (vehicle_no_dt_[i][6] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][9] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][11] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][13] != "") {
    c++;
  }

  if (c == 4) {
    vehicle_sh.getRange(i + 1, 6).setValue(new Date());
  }

  return JSON.stringify({
    'vehicle_no_dt': vehicle_no_dt_
  });
}

function setRCData(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var vehicle_sh = spreadsheet.getSheetByName("VEHCILE NUMBER")
  var vehicle_no_dt = vehicle_sh.getDataRange().getValues();
  data = JSON.parse(data);
  var address = data[0];
  var files = data[1];
  var vehicle_no_ = data[2];

  var folder_url = spreadsheet.getSheetByName("WHATSAPP CONTENT").getRange('G2').getValue();

  var blob = Utilities.newBlob(Utilities.base64Decode(files[0].data), files[0].mimeType, files[0].fileName);
  var fld = DriveApp.getFolderById(getIdFromUrl(folder_url)[0]);
  var file = fld.createFile(blob).getUrl();

  for (var i = 0; i < vehicle_no_dt.length; i++) {
    if (vehicle_no_dt[i][0] == vehicle_no_) {
      vehicle_sh.getRange(i + 1, 12, 1, 2).setValues([[file, address]]);
      break;
    }
  }
  SpreadsheetApp.flush();
  var vehicle_no_dt_ = vehicle_sh.getDataRange().getValues();

  var c = 0;

  if (vehicle_no_dt_[i][6] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][9] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][11] != "") {
    c++;
  }
  if (vehicle_no_dt_[i][13] != "") {
    c++;
  }

  if (c == 4) {
    vehicle_sh.getRange(i + 1, 6).setValue(new Date());
  }

  return JSON.stringify({
    'vehicle_no_dt': vehicle_no_dt_
  });
}

function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }

function setUploadReport(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var wt_RESPONSES = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  data = JSON.parse(data);
  var date = data[0];
  var indend_no = data[1];
  for (var i = 0; i < wt_RESPONSES.length; i++) {
    if (wt_RESPONSES[i][3] == indend_no) {
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 45, 1, 2).setValues([[new Date(), date]]);
      break;
    }
  }

  SpreadsheetApp.flush();
  var wt_RESPONSES1 = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  return JSON.stringify({
    'responses': wt_RESPONSES1
  });
}

function setUploadDone(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var wt_RESPONSES = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  data = JSON.parse(data);
  var date = data[0];
  var indend_no = data[1];
  for (var i = 0; i < wt_RESPONSES.length; i++) {
    if (wt_RESPONSES[i][3] == indend_no) {
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 48, 1, 2).setValues([[new Date(), date]]);
      break;
    }
  }

  SpreadsheetApp.flush();
  var wt_RESPONSES1 = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  return JSON.stringify({
    'responses': wt_RESPONSES1
  });
}

function setReleaseNl(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var wt_sheet = spreadsheet.getSheetByName("RESPONSES");
  var wt_RESPONSES = wt_sheet.getDataRange().getValues();
  data = JSON.parse(data);
  var reale_shipment_type = data[0];
  var reale_vehicle_no = data[1];
  var reale_vehicle_type = data[2];
  var invoices = data[3];
  var reale_no_of_cases = data[4];
  var reale_lr_no = data[5];
  var reale_lr_date = data[6];
  var indend_no = data[7];

  var temp = wt_RESPONSES.filter(function (rw) {
    return rw[36] == reale_vehicle_no;
  }).length;

  for (var i = 0; i < wt_RESPONSES.length; i++) {
    if (wt_RESPONSES[i][3] == indend_no) {
      reale_lr_date = Utilities.formatDate(new Date(reale_lr_date), 'GMT+05:30', 'dd/MMM/yyyy')
      if (temp == 0) {
        spreadsheet.getSheetByName("VEHCILE NUMBER").appendRow([reale_vehicle_no, reale_lr_no, reale_lr_date, wt_sheet.getRange(i + 1, 18).getValue()]);
      }

      wt_sheet.getRange(i + 1, 36, 1, 8).setValues([[new Date(), reale_vehicle_no, reale_vehicle_type, reale_shipment_type, invoices, reale_no_of_cases, reale_lr_no, reale_lr_date]]);

      break;
    }
  }

  SpreadsheetApp.flush();
  var wt_RESPONSES1 = wt_sheet.getDataRange().getValues();
  return JSON.stringify({
    'responses': wt_RESPONSES1
  });
}

function setOwnedData(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var wt_RESPONSES = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  data = JSON.parse(data);
  var aa_da = data[0], vehicle_number = data[1], indend_no = data[2], edit = data[3];
  for (var i = 0; i < wt_RESPONSES.length; i++) {
    if (wt_RESPONSES[i][3] == indend_no) {
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 20).setValue("OWNED");
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 22, 1, 3).setValues([[new Date(), vehicle_number, aa_da]]);
      if (edit) {
        spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 26, 1, 3).setValues([["", "", ""]]);
        spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 30, 1, 2).setValues([["", ""]]);
      }
      break;
    }
  }
  SpreadsheetApp.flush();
  var wt_RESPONSES1 = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  return JSON.stringify({
    'responses': wt_RESPONSES1
  });

}

function setMarketData(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var wt_RESPONSES = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  data = JSON.parse(data);
  var bn = data[1], rate = data[2], vehicle_no = data[3], advance_amt = data[4], indend_no = data[5], edit = data[6];
  for (var i = 0; i < wt_RESPONSES.length; i++) {
    if (wt_RESPONSES[i][3] == indend_no) {
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 20).setValue("MARKET");
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 26, 1, 3).setValues([[new Date(), bn, rate]]);
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 23, 1, 2).setValues([[vehicle_no, advance_amt]]);
      if (edit) {
        spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 30, 1, 2).setValues([["", ""]]);
        spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 22).setValue("");
      }
      break;
    }
  }
  SpreadsheetApp.flush();
  var wt_RESPONSES1 = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  return JSON.stringify({
    'responses': wt_RESPONSES1
  });

}

function setNonPlacementData(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var wt_RESPONSES = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  var wt_content = spreadsheet.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();
  data = JSON.parse(data);
  var none_placement = data[0], indend_no = data[1], dt, edit = data[2];
  for (var i = 0; i < wt_RESPONSES.length; i++) {
    if (wt_RESPONSES[i][3] == indend_no) {
      dt = wt_RESPONSES[i];
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 20).setValue("NON-PLACEMENT");
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 30, 1, 2).setValues([[new Date(), none_placement]]);
      if (edit) {
        spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 22, 1, 3).setValues([["", "", ""]]);
        spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 26, 1, 3).setValues([["", "", ""]]);
      }
      break;
    }
  }
  var number = wt_content[3][1];
  var content = wt_content[4][1];
  var mas_apiKey = wt_content[17][1];
  var mas_username = wt_content[15][1];
  var mas_password = wt_content[16][1];
  content = content.replace(/{DATE OF PLACEMENT}/, Utilities.formatDate(new Date(dt[7]), 'GMT+05:30', 'dd-MM-yyyy')).replace(/{REASON}/, none_placement).replace(/{DESTINATION}/, dt[6]).replace(/{SOURCE}/, dt[5]).replace(/{CLIENT NAME}/, dt[4]).replace(/{INDENT NO}/, indend_no);
  sendMessage(number, content, [], mas_apiKey, mas_username, mas_password);
  SpreadsheetApp.flush();
  var wt_RESPONSES1 = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  return JSON.stringify({
    'responses': wt_RESPONSES1
  });

}

function sendMessage(receivers, textMessages, filesUrls, mas_apiKey, mas_username, mas_password) {
  var messages = [].concat(textMessages || []);
  var urls = [].concat(filesUrls || []);
  var rawReceivers = [].concat(receivers || []);

  var receiverIds = rawReceivers.filter((item) => String(item).endsWith('@c.us') || String(item).endsWith('@g.us'));
  var receiverNumbers = rawReceivers.filter((item) => !receiverIds.includes(item));

  var driveFiles = urls.filter((filePath) => filePath.indexOf('drive.google.com') !== -1);
  var nonDriveFiles = urls.filter((filePath) => !driveFiles.includes(filePath));

  var messageBody = {
    username: mas_username,
    password: mas_password,
    receiverMobileNo: receiverNumbers.join(","),
    recipientIds: receiverIds,
    message: messages,
    filePathUrl: nonDriveFiles,
    base64File: driveFiles.map((url) => getGoogleFileAsBase64(url)),
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      ...mas_password && { 'Authorization': 'Basic ' + Utilities.base64Encode(mas_username + ':' + mas_password, Utilities.Charset.UTF_8) },
      ...mas_apiKey && { 'x-api-key': mas_apiKey }
    },
    payload: JSON.stringify(messageBody)
  };

  try {
    var sendNow = UrlFetchApp.fetch("https://app.messageautosender.com/api/v1/message/create", options);
    Logger.log(sendNow)
  } catch (e) {
    Logger.log(e)
  }
}

function getDriveFileIdFromUrl(url) {
  return url.match(/[-\w]{25,}/)
};

function getGoogleFileAsBase64(url) {
  var returnValue = {}
  if (url.indexOf('docs.google.com') !== -1) {
    var file = DocumentApp.openById(getDriveFileIdFromUrl(url));
    returnValue = {
      name: file.getName().replace(/\.[^/.]+$/, '') + '.pdf',
      body: Utilities.base64Encode(file.getAs('application/pdf').getBytes())
    };
  } else if (url.indexOf('drive.google.com') !== -1) {
    var file = DriveApp.getFileById(getDriveFileIdFromUrl(url));
    returnValue = {
      name: file.getName(),
      body: Utilities.base64Encode(file.getBlob().getBytes())
    }
  }
  return returnValue;
}

function getDatabaseData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var wt_RESPONSES = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  var doer_name = spreadsheet.getSheetByName("DOER NAME").getDataRange().getValues();
  var data_validation = spreadsheet.getSheetByName("DATAVALIDATION").getDataRange().getValues();
  var wt_content = spreadsheet.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();
  var vehicle_no_dt = spreadsheet.getSheetByName("VEHCILE NUMBER").getDataRange().getValues();
  var marketLoadVN = {};
  var doerNameBranch = {};
  data_validation.forEach(function (item) {
    if (item[29] != "") {
      if (!marketLoadVN[item[30]]) {
        marketLoadVN[item[30]] = [];
      }
      marketLoadVN[item[30]].push(item[29]);
    }
    if (item[20] != "") {
      if (!doerNameBranch[item[20]]) {
        doerNameBranch[item[20]] = [];
      }
      doerNameBranch[item[20]].push(item[21]);
    }
  })

  return JSON.stringify(
    {
      'doer_name': doer_name,
      'data_validation': data_validation,
      'responses': wt_RESPONSES,
      'company_name': wt_content[7][5],
      'vehicle_no_dt': vehicle_no_dt,
      marketLoadVN,
      doerNameBranch
    }
  );
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }

function saveIndentOrder(date, branch, itemsToSave) {
  console.log(date, branch, itemsToSave)
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("RESPONSES");
  var data_validation_sh = spreadsheet.getSheetByName("DATAVALIDATION");
  var data_validation_dt = data_validation_sh.getDataRange().getValues();

  var dataValidation_sh_dt = spreadsheet.getSheetByName('DATAVALIDATION').getRange("X2:AB").getValues().filter(function (rw) {
    return rw[0] != "" && (rw[1] != "" || rw[3] != "") && (rw[2].toString().trim() === "YES" || rw[4].toString().trim() === "YES") && rw[0].toString().trim() == branch.toString().trim();
  });


  var wt_content = spreadsheet.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();
  var responses_dt = sheet.getDataRange().getValues()[sheet.getLastRow() - 1];

  var uni_code = wt_content[2][6];
  var indent = Number(responses_dt[3]) + 1;
  var folder_id = getIdFromUrl(wt_content[1][6]);
  var urls = generateFileUrl(date, branch, itemsToSave, folder_id, uni_code);
  console.log(urls)
  var arr = [];
  for (var i = 0; i < itemsToSave.length; i++) {
    if (itemsToSave[i].row_ != "") {

      var fd = data_validation_dt.filter(function (rw) {
        return itemsToSave[i].client_name.toString().trim() === rw[5].toString().trim() && itemsToSave[i].source.toString().trim() === rw[6].toString().trim() && rw[7].toString().trim() === itemsToSave[i].destination.toString().trim();
      })[0][12];

      for (var k = 0; k < itemsToSave[i].no_truck; k++) {

        arr.push([
          uni_code, new Date(), "", indent, itemsToSave[i].client_name, itemsToSave[i].source,
          itemsToSave[i].destination, itemsToSave[i].date_palc, 1, itemsToSave[i].track_type,
          urls[1], urls[0], "", itemsToSave[i].CONSIGNOR, itemsToSave[i].CONSIGNEE, itemsToSave[i].TRANSIT_DAYS, date, branch, fd
        ]);
        indent++;
      }
    }
  }
  var uni_code_ = Number(uni_code.toString().slice(5,)) + 1;
  uni_code = uni_code.toString().slice(0, 5) + uni_code_;

  spreadsheet.getSheetByName("WHATSAPP CONTENT").getRange("G3").setValue(uni_code);

  sheet.getRange(sheet.getLastRow() + 1, 1, arr.length, arr[0].length).setValues(arr);

  // var number = wt_content[21][1];

  var mas_apiKey = wt_content[17][1];
  var mas_username = wt_content[15][1];
  var mas_password = wt_content[16][1];
  var content = wt_content[22][1];
  // var att = wt_content[23][1];
  content = content.replace(/{file}/, urls[0]);
  var number = [];

  if (dataValidation_sh_dt.length > 0) {
    if (dataValidation_sh_dt[0][2].toString().trim() === "YES") {
      number.push(dataValidation_sh_dt[0][1]);
    }
    if (dataValidation_sh_dt[0][4].toString().trim() === "YES") {
      number.push(dataValidation_sh_dt[0][3]);
    }
  }

  //TestsendMessage(number, content, [urls[0]], mas_apiKey, mas_username, mas_password, urls[2], uni_code);

  // if (att == "YES") {
  //   if (number.length > 0) {
  //     //sendMessage(number, '', [urls[0]], mas_apiKey, mas_username, mas_password);
  //   }
  //   else {
  //     //sendBranchMessage(data_validation_dt, branch, [urls[0]], content, mas_apiKey, mas_username, mas_password);
  //   }
  // }
  // else {
  //   if (number.length > 0) {
  //     //sendMessage(number, content, [], mas_apiKey, mas_username, mas_password);
  //   }
  //   else {
  //     //sendBranchMessage(data_validation_dt, branch, [], content, mas_apiKey, mas_username, mas_password);
  //   }
  // }

  return true;
}

function sendBranchMessage(data_validation_dt, branch, url_, content, mas_apiKey, mas_username, mas_password) {
  for (var i = 0; i < data_validation_dt.length; i++) {
    if (data_validation_dt[i][23].toString().trim() === branch.toString().trim() && data_validation_dt[i][24].toString().trim() != "") {

      if (data_validation_dt[i][25].toString().trim() == "YES" || data_validation_dt[i][27].toString().trim() == "YES") {
        if (data_validation_dt[i][25].toString().trim() == "YES") {
          var number = data_validation_dt[i][24].toString().trim();
          sendMessage(number, content, url_, mas_apiKey, mas_username, mas_password);
        }

        if (data_validation_dt[i][27].toString().trim() == "YES") {
          var number = data_validation_dt[i][26].toString().trim();
          sendMessage(number, content, url_, mas_apiKey, mas_username, mas_password);
        }

        break;
      }

    }
  }
}


// function sendPdfToWhatsappGroup(pdfFile, mas_apiKey, mas_username, mas_password) {
//   // Your Twilio Account SID and Auth Token from twilio.com/console
//   const accountSid = 'YOUR_ACCOUNT_SID';
//   const authToken = 'YOUR_AUTH_TOKEN';
//   const twilioNumber = 'whatsapp:+14155238886';
//   const whatsappGroupNumber = 'whatsapp:GROUP_PHONE_NUMBER';

//   // Create a new Twilio client instance
//   // const client = Twilio(accountSid, authToken);

//   const pdfContent = pdfFile.getBlob().getBytes();

//   //  var nonDriveFiles = urls.filter((filePath) => !driveFiles.includes(filePath));
//   // Create a new media message
//   // const message = {
//   //       username: mas_username,
//   //   password: mas_password,
//   //   // from: twilioNumber,
//   //   // to: whatsappGroupNumber,
//   //       receiverMobileNo: 9811710810,
//   // //   recipientIds: receiverIds,
//   //   body: 'Here is the PDF file you requested:',
//   //   mediaUrl: 'data:application/pdf;base64,' + Utilities.base64Encode(pdfContent),
//   // };
//   const messageBody = {
//     username: mas_username,
//     password: mas_password,
//     // from: twilioNumber,
//     // to: whatsappGroupNumber,
//     receiverMobileNo: 8883406121,
//     //   recipientIds: receiverIds,
//     message: 'Here is the PDF file you requested:',
//     // base64File: 'data:application/pdf;base64,' + Utilities.base64Encode(pdfContent),
//   };
//   // var filesUrls=[]
//   //  var urls = [].concat(filesUrls || []);
//   //  var nonDriveFiles = urls.filter((filePath) => !driveFiles.includes(filePath));
//   //   var messageBody = {
//   //   username: mas_username,
//   //   password: mas_password,
//   //   receiverMobileNo: 8883406121,
//   //   recipientIds: 8883406121,
//   //   message: "messages",
//   //   filePathUrl: nonDriveFiles,
//   //   base64File: pdfContent,
//   // };
//   var options = {
//     method: 'post',
//     contentType: 'application/json',
//     headers: {
//       ...mas_password && { 'Authorization': 'Basic ' + Utilities.base64Encode(mas_username + ':' + mas_password, Utilities.Charset.UTF_8) },
//       ...mas_apiKey && { 'x-api-key': mas_apiKey }
//     },
//     payload: JSON.stringify(messageBody)
//   };


//   // UrlFetchApp.fetch("https://app.messageautosender.com/api/v1/message/create", options);
//   try {
//     var sendNow = UrlFetchApp.fetch("https://app.messageautosender.com/api/v1/message/create", options);
//     Logger.log(sendNow)
//   } catch (e) {
//     Logger.log(e)
//   }
//   // Send the message
//   // client.messages
//   //   .create(message)
//   //   .then(message => console.log(message.sid))
//   //   .done();
// }
// function sendWhatsAppMessage(receivers, textMessages, filesUrls, mas_apiKey, mas_username, mas_password) {
//   var messages = [].concat(textMessages || []);
//   var urls = [].concat(filesUrls || []);
//   var rawReceivers = [].concat(receivers || []);

//   var receiverIds = rawReceivers.filter((item) => String(item).endsWith('@c.us') || String(item).endsWith('@g.us'));
//   var receiverNumbers = rawReceivers.filter((item) => !receiverIds.includes(item));

//   var driveFiles = urls.filter((filePath) => filePath.indexOf('drive.google.com') !== -1);
//   var nonDriveFiles = urls.filter((filePath) => !driveFiles.includes(filePath));

// var messageBody = {
//   username: mas_username,
//   password: mas_password,
//   receiverMobileNo: receiverNumbers.join(","),
//   recipientIds: receiverIds,
//   message: messages,
//   filePathUrl: nonDriveFiles,
//   base64File: driveFiles.map((url) => getGoogleFileAsBase64(url)),
// };

//   var options = {
//     method: 'post',
//     contentType: 'application/json',
//     headers: {
//       ...mas_password && { 'Authorization': 'Basic ' + Utilities.base64Encode(mas_username + ':' + mas_password, Utilities.Charset.UTF_8) },
//       ...mas_apiKey && { 'x-api-key': mas_apiKey }
//     },
//     payload: JSON.stringify(messageBody)
//   };

//   try {
//     var sendNow = UrlFetchApp.fetch("https://app.messageautosender.com/api/v1/message/create", options);
//     Logger.log(sendNow)
//   } catch (e) {
//     Logger.log(e)
//   }
// }


function setVehicleData(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var wt_RESPONSES = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();
  var wt_content = spreadsheet.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();
  data = JSON.parse(data);
  var indend_no = data[2];
  var ivnc = data[0];
  var cvn = data[1], dt = [];
  for (var i = 0; i < wt_RESPONSES.length; i++) {
    if (wt_RESPONSES[i][3] == indend_no) {
      dt = wt_RESPONSES[i];
      spreadsheet.getSheetByName("RESPONSES").getRange(i + 1, 32, 1, 3).setValues([[new Date(), ivnc, cvn]]);
      break;
    }
  }
  if (dt.length > 0) {
    try {
      if (ivnc == "NO") {
        var content = wt_content[27][1];
        var mas_apiKey = wt_content[17][1];
        var mas_username = wt_content[15][1];
        var mas_password = wt_content[16][1];
        var branch = dt[17];
        var data_validation_ = spreadsheet.getSheetByName("DATAVALIDATION").getDataRange().getValues().filter(function (rw) {
          return rw[23] == branch;
        })[0];

        var number = data_validation_[24];
        content = content.replace(/{CLIENT NAME}/, dt[4]).replace(/{DESTINATION}/, dt[6]).replace(/{SOURCE}/, dt[5]).replace(/{CORRECT VEHICLE NUMBER}/, cvn).replace(/{INDENT NO}/, indend_no);
        sendMessage(number, content, [], mas_apiKey, mas_username, mas_password);
      }
      var wt_RESPONSES1 = spreadsheet.getSheetByName("RESPONSES").getDataRange().getValues();

      SpreadsheetApp.flush();
      return JSON.stringify({
        'responses': wt_RESPONSES1
      });
    }
    catch (e) {
      return false;
    }
  }
  else {
    return false;
  }
}

function generateFileUrl(date, branch, itemsToSave, folder_id, uni_code) {
  // var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // var wt_content = spreadsheet.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();
  // var mas_apiKey = wt_content[17][1];
  // var mas_username = wt_content[15][1];
  // var mas_password = wt_content[16][1];
  // date = "25/2/2023"
  // branch = "KOLKATA"
  // folder_id = getIdFromUrl(wt_content[1][6]);
  folder_id = getIdFromUrl('https://drive.google.com/drive/folders/1wCmswf8ZFUBigMdHCl0jQYStl9dPaovF');
  // uni_code = wt_content[2][6];
  // itemsToSave = [{
  //   client_name: "ANKUR",
  //   source: "KOLKATA",
  //   destination: "PATNA",
  //   date_palc: "25/2/2023",
  //   no_truck: "432151",
  //   track_type: "new",
  //   "row_": 2,
  //   "TRANSIT_DAYS": 24,
  //   "CONSIGNOR": 'test',
  //   "CONSIGNEE": 'best',
  // }, {
  //   client_name: "ANKUR",
  //   source: "KOLKATA",
  //   destination: "PATNA",
  //   date_palc: "25/2/2023",
  //   no_truck: "432151",
  //   track_type: "new",
  //   "row_": 2,
  //   "TRANSIT_DAYS": 24,
  //   "CONSIGNOR": 'test',
  //   "CONSIGNEE": 'best',
  // }]
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dt = Utilities.formatDate(new Date(date), "GMT+05:30", "dd-MMM-yy");
  var folder = DriveApp.getFolderById(folder_id);
  var new_ss = SpreadsheetApp.create(uni_code);
  var file_sheet = spreadsheet.getSheetByName("EXCEL");
  var newSheet = file_sheet.copyTo(new_ss);
  new_ss.deleteSheet(new_ss.getSheetByName("Sheet1"));

  var new_ss_id = new_ss.getId();
  var new_s = DriveApp.getFileById(new_ss_id);
  new_s.moveTo(folder);
  var pgFrmtData = newSheet.getRange(1, 1, 57, 10).getValues();

  pgFrmtData[3][4] = dt;
  pgFrmtData[4][4] = branch;

  var count = 0, c = 0;
  for (var i = 0; i < itemsToSave.length; i++) {
    if (itemsToSave[i].row_ != "") {
      pgFrmtData[7 + count][3] = itemsToSave[i].client_name;
      pgFrmtData[7 + count][4] = itemsToSave[i].source;
      pgFrmtData[7 + count][5] = itemsToSave[i].destination;
      pgFrmtData[7 + count][6] = itemsToSave[i].date_palc;
      pgFrmtData[7 + count][7] = itemsToSave[i].no_truck;
      pgFrmtData[7 + count][8] = itemsToSave[i].track_type;
      count++;
      c = 7 + count;
    }
  }

  newSheet.getRange(1, 1, 57, 10).setValues(pgFrmtData);

  newSheet.hideRows(c + 3, 59 - (c + 7));

  SpreadsheetApp.flush();

  const blob = newSheet.getParent().getBlob().getAs('application/pdf');
  const pdfFile = folder.createFile(blob);
  // sendPdfToWhatsappGroup(pdfFile, mas_apiKey, mas_username, mas_password)
  // receivers = "919811710810-1469615550@g.us"
  // textMessages = ""
  // TestsendMessage(receivers, textMessages, [], mas_apiKey, mas_username, mas_password, pdfFile)
  return [pdfFile.getUrl(), "", pdfFile]
}


function TestsendMessage(receivers, textMessages, filesUrls, mas_apiKey, mas_username, mas_password, pdfFile, uni_code) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var wt_content = spreadsheet.getSheetByName("WHATSAPP CONTENT").getDataRange().getValues();
  var mas_apiKey = wt_content[17][1];
  var mas_username = wt_content[15][1];
  var mas_password = wt_content[16][1];
  filesUrls = []
  // pdfFile = DriveApp.getFileById('1ilr7oQUGKONtjPx67OPhrrUc1yAw7lsU');
  textMessages = ``
  console.log(textMessages)
  // receivers = ["919811710810-1469615550@g.us"]
  // receivers = ["919811710810-1469615550@g.us"]
  // receivers = ["919871081012-1467549753@g.us"]

  filesUrls = pdfFile.getUrl()
  var messages = [].concat(textMessages || []);
  var urls = [].concat(filesUrls || []);
  var rawReceivers = [].concat(receivers || []);

  var receiverIds = rawReceivers.filter((item) => String(item).endsWith('@c.us') || String(item).endsWith('@g.us'));
  var receiverNumbers = rawReceivers.filter((item) => !receiverIds.includes(item));

  const pdfContent = pdfFile.getBlob().getBytes();
  var pdfBase64 = Utilities.base64Encode(pdfContent);
  var messageBody = {
    username: mas_username,
    password: mas_password,
    receiverMobileNo: receiverNumbers.join(","),
    recipientIds: receiverIds,
    // message: messages,
    base64File: [{
      name: uni_code + '.pdf',
      body: pdfBase64
    }]
  };
  // console.log(messageBody)
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      ...mas_password && { 'Authorization': 'Basic ' + Utilities.base64Encode(mas_username + ':' + mas_password, Utilities.Charset.UTF_8) },
      ...mas_apiKey && { 'x-api-key': mas_apiKey }
    },
    payload: JSON.stringify(messageBody)
  };

  try {
    var sendNow = UrlFetchApp.fetch("https://app.messageautosender.com/api/v1/message/create", options);
    Logger.log(sendNow)
  } catch (e) {
    Logger.log(e)
  }
}
