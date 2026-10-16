/**
 * Contribution Group System - V1 Backend
 * One group -> payment items -> people pay with a TXN reference tag ->
 * system shows who has paid and who hasn't.
 *
 * Kept from the old system: reference tags, bank alert logging + auto-matcher,
 * manual match, offline payments, collection account settings, admin login.
 * Removed: batches, slots, slot requests, payouts, frequency/pace tracking.
 */

const SS = SpreadsheetApp.openById("1y2g8ro8SKV-eeh5trpWFCsGy3ZM8b2Vw0biBP5ZaDYE");
const SHEETS = {
  MEMBERS: "Members",
  ITEMS: "Payment_Items",
  CONTRIBUTIONS: "Contributions",
  ALERTS: "Bank_Alerts_Log",
  ADMINS: "Admins",
  SETTINGS: "Settings"
};

/*
 * Contributions columns (0-based):
 * 0 Transaction_ID | 1 Item_ID | 2 Payer_Name | 3 Payer_Phone | 4 Amount
 * 5 Payment_Method | 6 Status | 7 Timestamp | 8 Verified_By
 */

/**
 * Safe setup: only creates sheets that don't exist yet. It never clears existing data.
 * NOTE: the old "Contributions" sheet has a different column layout. If you already have one,
 * rename it (e.g. "Contributions_OLD") before running this so a fresh V1 sheet is created.
 */
function setupDatabase() {
  const sheetConfigs = [
    { name: SHEETS.MEMBERS, headers: ["Member_ID", "Full_Name", "Phone_Number", "Date_Joined", "Status", "Notes"] },
    { name: SHEETS.ITEMS, headers: ["Item_ID", "Item_Name", "Amount", "Deadline", "Status", "Created_At", "Notes"] },
    { name: SHEETS.CONTRIBUTIONS, headers: ["Transaction_ID", "Item_ID", "Payer_Name", "Payer_Phone", "Amount", "Payment_Method", "Status", "Timestamp", "Verified_By"] },
    { name: SHEETS.ALERTS, headers: ["Alert_ID", "Timestamp", "Bank_Name", "Raw_Amount", "Sender_Name", "Sender_Details", "Match_Status", "Matched_Transaction_ID"] },
    { name: SHEETS.SETTINGS, headers: ["Setting_Key", "Setting_Value", "Updated_At", "Updated_By"] },
    { name: SHEETS.ADMINS, headers: ["Admin_ID", "Full_Name", "Phone_Number", "Bank_Name", "Pin_Code", "Role"] }
  ];

  sheetConfigs.forEach(config => {
    let sheet = SS.getSheetByName(config.name);
    if (sheet) return; // never wipe existing data
    sheet = SS.insertSheet(config.name);
    sheet.appendRow(config.headers);
    sheet.getRange(1, 1, 1, config.headers.length)
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  });

  try {
    SS.getSheetByName(SHEETS.MEMBERS).getRange("C:C").setNumberFormat('@');
    SS.getSheetByName(SHEETS.CONTRIBUTIONS).getRange("D:D").setNumberFormat('@');
    SS.getSheetByName(SHEETS.ADMINS).getRange("C:C").setNumberFormat('@');
  } catch (e) {}

  try {
    SpreadsheetApp.getUi().alert("Database initialized successfully!");
  } catch (err) {
    Logger.log("UI alert skipped (non-UI context).");
  }
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (sheet.getName() === SHEETS.MEMBERS && col === 2 && row > 1) {
    const idCell = sheet.getRange(row, 1);
    if (!idCell.getValue()) idCell.setValue(nextId_(sheet, "MEM-", 1001));
    if (!sheet.getRange(row, 4).getValue()) sheet.getRange(row, 4).setValue(new Date());
    if (!sheet.getRange(row, 5).getValue()) sheet.getRange(row, 5).setValue("Active");
  }
}

/* ===================== ROUTING ===================== */

function doGet(e) {
  if (e.parameter && e.parameter.action) {
    const action = e.parameter.action;
    try {
      if (action === "getPublicInfo") return jsonResponse(getPublicInfo());
      if (action === "getOpenItems") return jsonResponse(getItems(true));
      if (action === "getAllItems") return jsonResponse(getItems(false));
      if (action === "getAdminData") return jsonResponse(getAdminDashboardData());
      if (action === "getUnmatchedAlerts") return jsonResponse(getUnmatchedAlerts());
      if (action === "getCollectionAccount") return jsonResponse({ status: "success", collectionAccount: getCollectionAccount() });
      if (action === "getItemSummary") return jsonResponse(getItemSummary(e.parameter.itemId));
      if (action === "getItemCsv") return jsonResponse(getItemCsv(e.parameter.itemId));
      if (action === "getMyPayments") return jsonResponse(getMyPayments(e.parameter.memberId, e.parameter.phone));
      return jsonResponse({ status: "error", message: "Invalid GET action." });
    } catch (err) {
      return jsonResponse({ status: "error", message: err.toString() });
    }
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Collectly')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    let data = {};
    if (e.parameter && Object.keys(e.parameter).length > 0) {
      data = e.parameter;
    } else if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        let params = {};
        e.postData.contents.split("&").forEach(part => {
          let item = part.split("=");
          if (item.length === 2) params[decodeURIComponent(item[0])] = decodeURIComponent(item[1]);
        });
        data = params;
      }
    }

    const action = data.action;
    Logger.log("Incoming Action: " + action);

    if (action === "submitPaymentIntent") return jsonResponse(recordPaymentIntent(data));
    if (action === "confirmPayment") return jsonResponse(updatePaymentStatus(data.transactionId, "Confirmed", data.adminName));
    if (action === "rejectPayment") return jsonResponse(updatePaymentStatus(data.transactionId, "Failed", data.adminName));
    if (action === "adminLogin") return jsonResponse(verifyAdminLogin(data.phone, data.pin));
    if (action === "addMember") return jsonResponse(addMember(data));
    if (action === "logBankAlert") return jsonResponse(logBankAlert(data));
    if (action === "createItem") return jsonResponse(createItem(data));
    if (action === "setItemStatus") return jsonResponse(setItemStatus(data.itemId, data.status));
    if (action === "recordOfflinePayment") return jsonResponse(recordOfflinePayment(data));
    if (action === "updateSetting") return jsonResponse(updateSetting(data.key, data.value, data.adminName));
    if (action === "manualMatchAlert") return jsonResponse(manualMatchAlert(data.alertId, data.transactionId, data.adminName));

    return jsonResponse({ status: "error", message: "Invalid POST action: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===================== HELPERS ===================== */

// Next ID like "MEM-1005" based on the highest existing number (safe if rows are deleted).
function nextId_(sheet, prefix, start) {
  const ids = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
  let max = start - 1;
  ids.forEach(r => {
    const s = String(r[0] || "");
    if (s.indexOf(prefix) === 0) {
      const n = Number(s.substring(prefix.length));
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return prefix + (max + 1);
}

function cleanPhone_(p) {
  return String(p || "").replace(/^'/, "").trim();
}

// Last 10 digits so 0803..., +234803... and 234803... all compare equal.
function phoneKey_(p) {
  const d = String(p || "").replace(/\D/g, "");
  return d.length >= 10 ? d.slice(-10) : "";
}

function nameKey_(n) {
  return String(n || "").toLowerCase().replace(/\s+/g, " ").trim();
}

function readMembers_() {
  const data = SS.getSheetByName(SHEETS.MEMBERS).getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const phone = cleanPhone_(data[i][2]);
    list.push({
      id: data[i][0],
      name: data[i][1],
      phone: phone,
      phoneKey: phoneKey_(phone),
      status: data[i][4] || "Active"
    });
  }
  return list;
}

function findItem_(itemId) {
  const data = SS.getSheetByName(SHEETS.ITEMS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === itemId) {
      return {
        row: i + 1,
        itemId: data[i][0],
        itemName: data[i][1],
        amount: Number(data[i][2]) || 0,
        deadline: data[i][3] || "",
        status: data[i][4] || "Open"
      };
    }
  }
  return null;
}

/* ===================== MEMBERS (optional, admin-managed) ===================== */

function addMember(data) {
  try {
    const sheet = SS.getSheetByName(SHEETS.MEMBERS);
    const fullName = String(data.fullName || "").trim();
    const phone = cleanPhone_(data.phone);
    if (!fullName || !phone) return { status: "error", message: "Full name and phone number are required." };

    const rows = sheet.getDataRange().getValues();
    const key = phoneKey_(phone);
    for (let i = 1; i < rows.length; i++) {
      if (key && phoneKey_(cleanPhone_(rows[i][2])) === key) {
        return { status: "error", message: "This phone number is already on the member list." };
      }
    }

    const memberId = nextId_(sheet, "MEM-", 1001);
    const targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 3).setNumberFormat('@');
    sheet.getRange(targetRow, 1, 1, 6).setValues([[
      memberId, fullName, phone, new Date(), "Active", data.notes || "Added by admin"
    ]]);

    return { status: "success", message: `Member added. Member ID: ${memberId}`, memberId: memberId };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

/* ===================== PAYMENT ITEMS ===================== */

function createItem(data) {
  const sheet = SS.getSheetByName(SHEETS.ITEMS);
  const name = String(data.itemName || "").trim();
  const amount = Number(data.amount);
  if (!name) return { status: "error", message: "Item name is required." };
  if (isNaN(amount) || amount < 0) return { status: "error", message: "Enter a valid amount (use 0 for open-ended items)." };

  const itemId = nextId_(sheet, "ITM-", 1001);
  sheet.appendRow([itemId, name, amount, data.deadline || "", "Open", new Date(), data.notes || ""]);
  return { status: "success", message: `Payment item "${name}" created (${itemId}).`, itemId: itemId };
}

function setItemStatus(itemId, status) {
  const item = findItem_(itemId);
  if (!item) return { status: "error", message: "Item not found." };
  if (status !== "Open" && status !== "Closed") return { status: "error", message: "Invalid status." };
  SS.getSheetByName(SHEETS.ITEMS).getRange(item.row, 5).setValue(status);
  return { status: "success", message: `${item.itemName} is now ${status}.` };
}

function getItems(openOnly) {
  const data = SS.getSheetByName(SHEETS.ITEMS).getDataRange().getValues();
  const items = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const status = data[i][4] || "Open";
    if (openOnly && String(status).toLowerCase() !== "open") continue;
    items.push({
      itemId: data[i][0],
      itemName: data[i][1],
      amount: Number(data[i][2]) || 0,
      deadline: data[i][3] || "",
      status: status
    });
  }
  return { status: "success", items: items };
}

function getPublicInfo() {
  return {
    status: "success",
    groupName: getSetting("Group_Name") || "Contribution Group",
    collectionAccount: getCollectionAccount()
  };
}

/* ===================== PAYMENTS ===================== */

function recordPaymentIntent(data) {
  const item = findItem_(data.itemId);
  if (!item) return { status: "error", message: "Payment item not found." };
  if (String(item.status).toLowerCase() !== "open") return { status: "error", message: "This payment item is closed." };

  const amount = Number(data.amount);
  if (isNaN(amount) || amount <= 0) return { status: "error", message: "Enter a valid amount." };

  const contribSheet = SS.getSheetByName(SHEETS.CONTRIBUTIONS);
  const txId = "TXN-" + Math.floor(100000 + Math.random() * 900000);
  const phone = cleanPhone_(data.phone);

  const targetRow = contribSheet.getLastRow() + 1;
  contribSheet.getRange(targetRow, 4).setNumberFormat('@');
  contribSheet.getRange(targetRow, 1, 1, 9).setValues([[
    txId, item.itemId, String(data.payerName || "").trim() || "Anonymous", phone, amount,
    "Bank Transfer", "Pending", new Date(), ""
  ]]);

  return {
    status: "success",
    message: `Reference Tag generated: ${txId}. Include this tag in your bank transfer narration for ₦${amount.toLocaleString()}.`,
    transactionId: txId,
    itemName: item.itemName,
    amount: amount,
    collectionAccount: getCollectionAccount()
  };
}

function updatePaymentStatus(transactionId, newStatus, adminName) {
  const sheet = SS.getSheetByName(SHEETS.CONTRIBUTIONS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === transactionId) {
      if (data[i][6] === "Confirmed" && newStatus !== "Confirmed") {
        return { status: "error", message: "A confirmed payment can't be rejected here." };
      }
      sheet.getRange(i + 1, 7).setValue(newStatus);
      sheet.getRange(i + 1, 9).setValue(adminName || "Admin");
      return { status: "success", message: `Transaction ${transactionId} marked as ${newStatus}.` };
    }
  }
  return { status: "error", message: "Transaction ID not found." };
}

function recordOfflinePayment(data) {
  const item = findItem_(data.itemId);
  if (!item) return { status: "error", message: "Select a valid payment item." };
  const amount = Number(data.amount) || 0;
  if (amount <= 0) return { status: "error", message: "A valid amount is required." };

  const sheet = SS.getSheetByName(SHEETS.CONTRIBUTIONS);
  const txId = "OFF-" + Math.floor(100000 + Math.random() * 900000);
  const targetRow = sheet.getLastRow() + 1;
  sheet.getRange(targetRow, 4).setNumberFormat('@');
  sheet.getRange(targetRow, 1, 1, 9).setValues([[
    txId, item.itemId, String(data.payerName || "").trim() || "Offline Payer", cleanPhone_(data.phone), amount,
    data.paymentMethod || "Offline / Cash / Direct Transfer", "Confirmed", new Date(), data.adminName || "Admin"
  ]]);

  return {
    status: "success",
    message: `Offline payment recorded as ${txId}. ₦${amount.toLocaleString()} applied to ${item.itemName}.`,
    transactionId: txId
  };
}

/* ===================== BANK ALERTS ===================== */

function runAutoMatcher(alertId, bankName, rawAmount, senderName, rawNarration) {
  const targetBank = String(bankName || "").toUpperCase();
  if (!targetBank.includes("KUDA") && !targetBank.includes("MONIEPOINT") && !targetBank.includes("ALAT")) {
    updateAlertStatus(alertId, "Ignored (Non-Target Bank)", "");
    return { status: "success", matched: false, message: "Alert ignored. Only supported partner banks are accepted." };
  }

  const sheet = SS.getSheetByName(SHEETS.CONTRIBUTIONS);
  const rows = sheet.getDataRange().getValues();
  const alertAmount = Number(rawAmount);
  const narration = String(rawNarration || "").toUpperCase();

  for (let i = 1; i < rows.length; i++) {
    const txId = String(rows[i][0] || "");
    if (rows[i][6] === "Pending" && alertAmount === Number(rows[i][4]) && narration.includes(txId.toUpperCase())) {
      sheet.getRange(i + 1, 7).setValue("Confirmed");
      sheet.getRange(i + 1, 9).setValue(`Auto-Matched (${bankName})`);
      updateAlertStatus(alertId, "Matched", txId);
      return { status: "success", matched: true, transactionId: txId };
    }
  }

  updateAlertStatus(alertId, "Unmatched", "");
  return { status: "success", matched: false, message: "No matching reference code found. Available for manual review." };
}

function logBankAlert(data) {
  const alertSheet = SS.getSheetByName(SHEETS.ALERTS);
  const alertId = "ALT-" + Math.floor(100000 + Math.random() * 900000);

  const rawText = data.senderDetails || data.text || data.body || JSON.stringify(data);
  const bankName = data.bankName || "ALAT";

  const amountMatch = rawText.match(/(?:NGN|₦)\s*([0-9,]+\.[0-9]+|[0-9,]+)/i) ||
                      rawText.match(/Amount[:\s]*([0-9,.]+)/i) ||
                      rawText.match(/([0-9,]+\.[0-9]{2})/);

  let parsedAmount = amountMatch ? Number(amountMatch[1].replace(/,/g, '')) : (Number(data.amount) || 0);

  const txMatch = rawText.match(/TXN-[A-Za-z0-9]+/i);
  const extractedTxId = txMatch ? txMatch[0].toUpperCase() : null;

  // Sender extraction: prioritise text after "NIP:"
  let senderName = data.senderName || "Bank User";
  const nipMatch = rawText.match(/NIP[:\s]*([A-Za-z0-9\s\.\-]+)/i);
  if (nipMatch) {
    senderName = nipMatch[1].trim().split(/\s+/).slice(0, 3).join(" ");
  } else {
    const senderMatch = rawText.match(/Sender[:\s]*([^\r\n]+)/i) || rawText.match(/From[:\s]*([^\r\n]+)/i);
    if (senderMatch) senderName = senderMatch[1].trim();
  }

  const noteMatch = rawText.match(/Note[:\s]*([^\r\n]+)/i) || rawText.match(/Narration[:\s]*([^\r\n]+)/i);
  const parsedNarration = noteMatch ? noteMatch[1].trim() : rawText;

  const initialStatus = extractedTxId ? "Unmatched" : "No Reference (Manual Review)";

  alertSheet.appendRow([
    alertId, new Date(), bankName, parsedAmount, senderName,
    parsedNarration.substring(0, 500), initialStatus, ""
  ]);

  if (extractedTxId) {
    return runAutoMatcher(alertId, bankName, parsedAmount, senderName, parsedNarration);
  }
  return {
    status: "success",
    matched: false,
    alertId: alertId,
    message: "Logged without reference tag. Available for manual matching."
  };
}

function updateAlertStatus(alertId, matchStatus, matchedTxId) {
  const alertSheet = SS.getSheetByName(SHEETS.ALERTS);
  const data = alertSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === alertId) {
      alertSheet.getRange(i + 1, 7).setValue(matchStatus);
      alertSheet.getRange(i + 1, 8).setValue(matchedTxId || "");
      break;
    }
  }
}

function manualMatchAlert(alertId, transactionId, adminName) {
  if (!alertId || !transactionId) {
    return { status: "error", message: "Both alertId and transactionId are required." };
  }

  const alertSheet = SS.getSheetByName(SHEETS.ALERTS);
  const contribSheet = SS.getSheetByName(SHEETS.CONTRIBUTIONS);

  const alertData = alertSheet.getDataRange().getValues();
  let alertFound = false;
  for (let i = 1; i < alertData.length; i++) {
    if (alertData[i][0] === alertId) {
      alertFound = true;
      if (String(alertData[i][6]).toLowerCase().includes("matched") && String(alertData[i][6]).toLowerCase() !== "unmatched") {
        return { status: "error", message: `Alert ${alertId} is already matched.` };
      }
      break;
    }
  }
  if (!alertFound) return { status: "error", message: "Alert ID not found." };

  const contribData = contribSheet.getDataRange().getValues();
  for (let i = 1; i < contribData.length; i++) {
    if (contribData[i][0] === transactionId) {
      if (contribData[i][6] !== "Pending") {
        return { status: "error", message: `Transaction ${transactionId} is not in Pending status.` };
      }
      contribSheet.getRange(i + 1, 7).setValue("Confirmed");
      contribSheet.getRange(i + 1, 9).setValue(adminName || "Manual Match");
      updateAlertStatus(alertId, "Manually Matched", transactionId);
      return { status: "success", message: `Alert ${alertId} manually matched to ${transactionId}. Payment confirmed.` };
    }
  }
  return { status: "error", message: "Transaction ID not found." };
}

function getUnmatchedAlerts() {
  const data = SS.getSheetByName(SHEETS.ALERTS).getDataRange().getValues();
  const unmatched = [];
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][6] || "");
    if (status === "No Reference (Manual Review)" || status === "Unmatched") {
      unmatched.push({
        alertId: data[i][0],
        timestamp: data[i][1],
        bankName: data[i][2],
        amount: data[i][3],
        senderName: data[i][4],
        details: data[i][5],
        matchStatus: status
      });
    }
  }
  return { status: "success", unmatchedAlerts: unmatched };
}

/* ===================== ADMIN ===================== */

function getAdminDashboardData() {
  const data = SS.getSheetByName(SHEETS.CONTRIBUTIONS).getDataRange().getValues();
  const itemNames = {};
  getItems(false).items.forEach(it => itemNames[it.itemId] = it.itemName);

  const pendingPayments = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === "Pending") {
      pendingPayments.push({
        transactionId: data[i][0],
        itemId: data[i][1],
        itemName: itemNames[data[i][1]] || data[i][1],
        payerName: data[i][2],
        payerPhone: cleanPhone_(data[i][3]),
        amount: data[i][4],
        method: data[i][5],
        timestamp: data[i][7]
      });
    }
  }
  return { status: "success", pendingPayments: pendingPayments };
}

function verifyAdminLogin(phone, pin) {
  const adminSheet = SS.getSheetByName(SHEETS.ADMINS);
  if (!adminSheet) return { status: "error", message: "Admins sheet missing." };

  const data = adminSheet.getDataRange().getValues();
  const cleanInputPhone = String(phone || "").replace(/\D/g, "").trim();
  const cleanInputPin = String(pin || "").trim();

  for (let i = 1; i < data.length; i++) {
    const sheetPhone = String(data[i][2] || "").replace(/^'/, "").replace(/\D/g, "").trim();
    const sheetPin = String(data[i][4] || "").trim();
    if (sheetPhone === cleanInputPhone && sheetPin === cleanInputPin) {
      return { status: "success", admin: { name: data[i][1], role: data[i][5] || "Admin" } };
    }
  }
  return { status: "error", message: "Invalid Admin Credentials." };
}

/* ===================== SETTINGS ===================== */

function getSetting(key) {
  const sheet = SS.getSheetByName(SHEETS.SETTINGS);
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return String(data[i][1] || "").trim();
  }
  return "";
}

function getCollectionAccount() {
  return {
    bankName: getSetting("Collection_Bank_Name") || "Not set",
    accountNumber: getSetting("Collection_Account_Number") || "Not set",
    accountName: getSetting("Collection_Account_Name") || "Not set"
  };
}

function updateSetting(key, value, updatedBy) {
  const sheet = SS.getSheetByName(SHEETS.SETTINGS);
  if (!sheet) return { status: "error", message: "Settings sheet missing." };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      sheet.getRange(i + 1, 3).setValue(new Date());
      sheet.getRange(i + 1, 4).setValue(updatedBy || "Admin");
      return { status: "success", message: `${key} updated.` };
    }
  }
  sheet.appendRow([key, value, new Date(), updatedBy || "Admin"]);
  return { status: "success", message: `${key} created.` };
}

/* ===================== REPORTS: PAID vs YET TO PAY ===================== */

/**
 * For one payment item: who has paid in full, who has paid part, and which
 * listed members have not paid at all. Payers are matched to members by phone
 * (last 10 digits) first, then by name. Non-members who paid still appear.
 */
function getItemSummary(itemId) {
  const item = findItem_(itemId);
  if (!item) return { status: "error", message: "Payment item not found." };

  const members = readMembers_();
  const byPhone = {}, byName = {};
  members.forEach(m => {
    if (m.phoneKey) byPhone[m.phoneKey] = m;
    const nk = nameKey_(m.name);
    if (nk) byName[nk] = m;
  });

  const rows = SS.getSheetByName(SHEETS.CONTRIBUTIONS).getDataRange().getValues();
  const groups = {};
  let collected = 0;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] !== itemId || rows[i][6] !== "Confirmed") continue;
    const amt = Number(rows[i][4]) || 0;
    collected += amt;

    const pk = phoneKey_(rows[i][3]);
    const nk = nameKey_(rows[i][2]);
    const m = (pk && byPhone[pk]) || (nk && byName[nk]) || null;
    const key = m ? "m:" + m.id : (pk ? "p:" + pk : "n:" + nk);

    if (!groups[key]) {
      groups[key] = {
        name: m ? m.name : (rows[i][2] || "Anonymous"),
        memberId: m ? m.id : "",
        isMember: !!m,
        total: 0,
        payments: 0,
        lastPaid: null
      };
    }
    groups[key].total += amt;
    groups[key].payments += 1;
    groups[key].lastPaid = rows[i][7];
  }

  const paidList = [], partialList = [];
  Object.keys(groups).forEach(k => {
    const g = groups[k];
    const full = g.total > 0 && g.total >= item.amount;
    if (full) paidList.push(g);
    else partialList.push(Object.assign({}, g, { owing: Math.max(0, item.amount - g.total) }));
  });

  const notPaidList = members
    .filter(m => String(m.status).toLowerCase() !== "inactive" && !groups["m:" + m.id])
    .map(m => ({ memberId: m.id, name: m.name, phone: m.phone, owing: item.amount }));

  return {
    status: "success",
    summary: {
      itemId: item.itemId,
      itemName: item.itemName,
      amount: item.amount,
      deadline: item.deadline,
      itemStatus: item.status,
      collected: collected,
      paidList: paidList,
      partialList: partialList,
      notPaidList: notPaidList
    }
  };
}

function getItemCsv(itemId) {
  const res = getItemSummary(itemId);
  if (res.status !== "success") return res;
  const s = res.summary;

  const esc = v => '"' + String(v === undefined || v === null ? "" : v).replace(/"/g, '""') + '"';
  const lines = [["Item", "Name", "Member_ID", "Status", "Amount_Paid", "Amount_Owing"].map(esc).join(",")];

  s.paidList.forEach(p => lines.push([s.itemName, p.name, p.memberId, "Paid", p.total, 0].map(esc).join(",")));
  s.partialList.forEach(p => lines.push([s.itemName, p.name, p.memberId, "Partial", p.total, p.owing].map(esc).join(",")));
  s.notPaidList.forEach(p => lines.push([s.itemName, p.name, p.memberId, "Not Paid", 0, p.owing].map(esc).join(",")));

  return {
    status: "success",
    filename: `${s.itemId}_${s.itemName.replace(/[^A-Za-z0-9]+/g, "_")}.csv`,
    csv: lines.join("\n")
  };
}

/* ===================== MEMBER "MY PAYMENTS" (optional) ===================== */

function getMyPayments(memberId, phone) {
  const members = readMembers_();
  const pk = phoneKey_(phone);
  const member = members.filter(m => m.id === memberId && pk && m.phoneKey === pk)[0];
  if (!member) return { status: "error", message: "Invalid Member ID or Phone." };

  const rows = SS.getSheetByName(SHEETS.CONTRIBUTIONS).getDataRange().getValues();
  const nk = nameKey_(member.name);
  const paidByItem = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][6] !== "Confirmed") continue;
    const rowPk = phoneKey_(rows[i][3]);
    const mine = (rowPk && rowPk === member.phoneKey) || (!rowPk && nameKey_(rows[i][2]) === nk);
    if (mine) paidByItem[rows[i][1]] = (paidByItem[rows[i][1]] || 0) + (Number(rows[i][4]) || 0);
  }

  const items = getItems(false).items.map(it => {
    const paid = paidByItem[it.itemId] || 0;
    const status = (paid > 0 && paid >= it.amount) ? "Paid" : (paid > 0 ? "Partial" : "Not Paid");
    return {
      itemId: it.itemId, itemName: it.itemName, amount: it.amount,
      deadline: it.deadline, itemStatus: it.status, paid: paid, status: status
    };
  });

  return { status: "success", member: { memberId: member.id, fullName: member.name }, items: items };
}
