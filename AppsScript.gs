function doGet(e) {
  try {
    var code = e && e.parameter && e.parameter.code ? e.parameter.code.trim().toUpperCase() : "";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var guestSheet = ss.getSheetByName("GuestList");
    if (!guestSheet) return json({found: false, error: "Tab 'GuestList' not found."});

    var guestData = guestSheet.getDataRange().getValues();
    var guestRow = null;
    for (var i = 1; i < guestData.length; i++) {
      if (guestData[i][0] && guestData[i][0].toString().trim().toUpperCase() === code) {
        guestRow = guestData[i];
        break;
      }
    }
    if (!guestRow) return json({found: false, error: "Passcode not found."});

    var result = {
      found: true,
      code: guestRow[0],
      name: guestRow[1],
      haldi: guestRow[2].toString().trim().toUpperCase() === "YES",
      mehndi: guestRow[3].toString().trim().toUpperCase() === "YES",
      wedding: guestRow[4].toString().trim().toUpperCase() === "YES",
      reception: guestRow[5].toString().trim().toUpperCase() === "YES",
      previousResponse: null
    };

    // Look up any RSVP they already submitted, so the site can pre-fill it
    var rsvpSheet = ss.getSheetByName("RSVP_Responses");
    if (rsvpSheet) {
      var rsvpData = rsvpSheet.getDataRange().getValues();
      for (var j = 1; j < rsvpData.length; j++) {
        if (rsvpData[j][1] && rsvpData[j][1].toString().trim().toUpperCase() === code) {
          result.previousResponse = {
            haldi: rsvpData[j][3],
            wedding: rsvpData[j][5],
            reception: rsvpData[j][6],
            dietary: rsvpData[j][7]
          };
          break;
        }
      }
    }

    return json(result);
  } catch (err) {
    return json({found: false, error: err.toString()});
  }
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RSVP_Responses");
    var data = JSON.parse(e.postData.contents);
    var codeUpper = data.code.toString().trim().toUpperCase();

    var values = sheet.getDataRange().getValues();
    var existingRow = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][1] && values[i][1].toString().trim().toUpperCase() === codeUpper) {
        existingRow = i + 1; // sheet rows are 1-indexed
        break;
      }
    }

    var rowValues = [
      new Date(),
      data.code,
      data.name,
      data.haldi || "N/A",
      data.mehndi || "N/A",
      data.wedding || "N/A",
      data.reception || "N/A",
      data.dietary || "None"
    ];

    if (existingRow > -1) {
      // Update their existing row instead of creating a duplicate
      sheet.getRange(existingRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return json({result: "success"});
  } catch (err) {
    return json({result: "error", error: err.toString()});
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
