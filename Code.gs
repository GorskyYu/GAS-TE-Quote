function onEdit(e) {
  const sheet = e.range.getSheet();
  const cell = e.range;

  // ✅ New: Auto-run quote if B11 is changed to "Yes"
  if (sheet.getName() === "Sheet1" && cell.getA1Notation() === 'B11' && cell.getValue() === 'Yes') {
    getTripleEagleQuoteFromSheet();
    cell.clearContent(); // clear B11
    return;
  }

  // Feature 1: Clear and set values when C20 is edited to "Yes"
  if (cell.getA1Notation() === 'C20' && cell.getValue() === 'Yes') {
    clearAndSetValues(sheet);
    sheet.getRange("C20").clearContent(); // ✅ clears the trigger cell
  }

  // Feature 2: Set L2:L6 to 0 when B10 is edited to "加境內" or "加台海運"
  if (cell.getA1Notation() === 'B10') {
    const value = cell.getValue();
    
    // Set I10 based on B10 value
    const i10Cell = sheet.getRange('I10');
    if (value === '加境內') {
      setValuesToZero(sheet);
      i10Cell.setValue(0);
    } else if (value === '加台空運') {
      setValuesToOne(sheet);
      i10Cell.setValue(10);
    } else if (value === '加台海運') {
      setValuesToZero(sheet);
      i10Cell.setValue(5);
    }
  }
}

function clearAndSetValues(sheet) {
  const rangesToClear = ['B2:C6', 'B8:B9','B10', 'P2', 'Q4', 'Q2'];
  const rangeToSet = sheet.getRange('L2:L6');

  // Clear content of specified ranges
  rangesToClear.forEach(r => {
    sheet.getRange(r).clearContent();
  });

  // Set L2:L6 to 1
  const ones = Array(rangeToSet.getNumRows()).fill([1]);
  rangeToSet.setValues(ones);

  // Finally, clear content in C17
  sheet.getRange('C17').clearContent();
}

function setValuesToZero(sheet) {
  const rangeToSet = sheet.getRange('L2:L6');

  // Set L2:L6 to 0
  const zeros = Array(rangeToSet.getNumRows()).fill([0]);
  rangeToSet.setValues(zeros);
}

function setValuesToOne(sheet) {
  const rangeToSet = sheet.getRange('L2:L6');

  // Set L2:L6 to 1
  const ones = Array(rangeToSet.getNumRows()).fill([1]);
  rangeToSet.setValues(ones);
}

function getTripleEagleQuoteFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Sheet1");
  const outputSheet = ss.getSheetByName("Quote Test");

  const appId = '584';
  const secret = 'da5f3bec04ed5ec082d8a0a7f04e12e7';
  const timestamp = Math.floor(Date.now() / 1000);
  const action = 'shipment/quote';
  const format = 'json';

  const getParams = {
    id: appId,
    timestamp: timestamp.toString(),
    action: action,
    format: format
  };
  const sign = generateSignature(getParams, secret);
  getParams.sign = sign;

  const url = 'https://eship.tripleeaglelogistics.com/api?' + toQueryString(getParams);

  const fromPostal = inputSheet.getRange("B8").getValue().toString().trim();
  const toPostal = inputSheet.getRange("B9").getValue().toString().trim();
  const boxData = inputSheet.getRange("A2:C6").getValues(); // 5 boxes, 3 columns

  const packages = [];

  for (const [id, size, weight] of boxData) {
    if (!weight || !size || typeof size !== "string") continue;

    const parts = size.split("*").map(s => parseFloat(s));
    if (parts.length !== 3 || parts.some(n => isNaN(n))) {
      Logger.log(`⚠️ Skipped box "${id}" due to malformed size: "${size}"`);
      continue;
    }

    const [lengthCm, widthCm, heightCm] = parts;
    const toInches = cm => Math.round((cm / 2.54) * 100) / 100;

    packages.push({
      weight: parseFloat(weight),
      dimension: {
        length: toInches(lengthCm),
        width: toInches(widthCm),
        height: toInches(heightCm)
      },
      insurance: 100
    });
  }

  const postPayload = {
    initiation: {
      region_id: 'CA',
      postalcode: fromPostal
    },
    destination: {
      region_id: 'CA',
      postalcode: toPostal
    },
    package: {
      type: 'parcel',
      packages: packages
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(postPayload),
    headers: { 'Accept-Language': 'en-US' },
    muteHttpExceptions: true
  };

  Logger.log("📤 Post payload being sent:");
  Logger.log(JSON.stringify(postPayload, null, 2));
  Logger.log("📤 Full GET URL:");
  Logger.log(url);

  const response = UrlFetchApp.fetch(url, options);

  const rawText = response.getContentText();
  Logger.log("📦 Raw API response:");
  Logger.log(rawText);

  const json = JSON.parse(rawText);

  outputSheet.getRange("H1:N1").setValues([[
    "Carrier / Service",
    "Freight Base (CAD)",
    "Tax (CAD)",
    "Total Charge (CAD)",
    "ETA",
    "Surcharges (CAD)",
    "Surcharge Details"
  ]]);

  if (json.status === 1 && json.response) {
    const allServices = [];

    json.response.forEach(carrier => {
      const vendor = carrier.name;
      carrier.services.forEach(service => {
        allServices.push({
          carrier: vendor,
          name: service.name,
          charge: parseFloat(service.charge),
          eta: service.eta || 'N/A'
        });
      });
    });

    if (allServices.length === 0) {
      outputSheet.getRange("H2:N3").setValues([["No services found", "", "", "", "", "", ""], ["", "", "", "", "", "", ""]]);
      inputSheet.getRange("P2").setValue("—");
      return;
    }

    allServices.sort((a, b) => a.charge - b.charge);
    const top2 = allServices.slice(0, 2);

    // Attach full service detail
    top2.forEach(q => {
      q._service = json.response
        .flatMap(r => r.services)
        .find(s => s.name === q.name && parseFloat(s.charge) === q.charge);
    });

    const rows = top2.map(q => {
      const s = q._service;

      const base = parseFloat(s.freight || 0);
      const total = parseFloat(s.charge || 0);
      const eta = s.eta || 'N/A';

      const tax = (s.tax_details || []).reduce((sum, t) => sum + parseFloat(t.price), 0);

      let surchargeSum = 0;
      const surchargeDetails = [];

      (s.charge_details || []).forEach(d => {
        const price = parseFloat(d.price || 0);
        surchargeSum += price;
        surchargeDetails.push(`${d.name}: $${price.toFixed(2)}`);
      });

      return [
        `${q.carrier} - ${q.name}`,
        base,
        tax,
        total,
        eta,
        surchargeSum,
        surchargeDetails.join("; ")
      ];
    });

    while (rows.length < 2) rows.push(["—", "", "", "", "", "", ""]);

    outputSheet.getRange(2, 8, 2, 7).setValues(rows); // H2:N3

    // Write cheapest Total Charge to Sheet1!P2
    inputSheet.getRange("P2").setValue(rows[0][3]); // column 3 = Total Charge
  } else {
    outputSheet.getRange("H2:N3").setValues([
      ["❌ " + (json.message || "API error"), "", "", "", "", "", ""],
      ["", "", "", "", "", "", ""]
    ]);
    inputSheet.getRange("P2").setValue("❌");
  }
}

function generateSignature(params, secretKey) {
  const sortedKeys = Object.keys(params)
    .map(k => k.toLowerCase())
    .filter(k => k !== 'sign')
    .sort();

  const encodedParams = sortedKeys.map(k => {
    const val = encodeURIComponent(params[k]).replace(/%7E/g, '~');
    return `${k}=${val}`;
  });

  const queryString = encodedParams.join('&');
  const rawSignature = Utilities.computeHmacSha256Signature(queryString, secretKey);
  return Utilities.base64Encode(rawSignature);
}

function toQueryString(params) {
  return Object.entries(params)
    .map(([k, v]) => encodeURIComponent(k) + '=' + encodeURIComponent(v))
    .join('&');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("TripleEagle API")
    .addItem("Get Shipping Quote", "getTripleEagleQuoteFromSheet")
    .addToUi();
}