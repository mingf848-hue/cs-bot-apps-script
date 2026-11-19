function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);

    // Run before workload month/sheet selection: monitoring can target another workbook.
    if (params.action === "get_sheet_snapshot") {
      return handleGetSheetSnapshot(params);
    }

    if (params.action === "sync_a69_workload") {
      return handleSyncA69Workload(params);
    }

    var sheetInfo = getSmartSheetInfo(params.day, params.month, params.year);
    var sheet = sheetInfo.sheet;

    if (params.action === "get_keywords") {
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      }
      return handleGetKeywords(sheet);
    }

    if (!sheet) {
      return responseJSON({
        success: false,
        msg: "找不到工作表: [" + sheetInfo.name + "]。请确认表格已创建且名字正确。"
      });
    }

    if (params.action === "sync_manual_workload") {
      return handleWriteManualWorkload(sheet, params, sheetInfo.year, sheetInfo.month);
    }

    if (params.day && params.stats) {
      return handleWriteStats(sheet, params.day, params.stats, sheetInfo.year, sheetInfo.month);
    }

    return responseJSON({ success: false, msg: "未知请求类型" });
  } catch (err) {
    return responseJSON({ success: false, msg: "脚本报错: " + err.toString() });
  }
}

function handleSyncA69Workload(params) {
  var year = parseInt(params.year, 10);
  var month = parseInt(params.month, 10);
  var day = parseInt(params.day, 10);
  if (year !== 2026 || month !== 11 || !day || day < 1 || day > 30) {
    return responseJSON({success: false, msg: "A69 目标表仅对应 2026 年 11 月"});
  }

  var sourceInfo = getSmartSheetInfo(day, month, year);
  var source = sourceInfo.sheet;
  if (!source) return responseJSON({success: false, msg: "找不到来源表: " + sourceInfo.name});
  var sourceColumn = findDateColumn(source, day, year, month);
  if (!sourceColumn) return responseJSON({success: false, msg: "来源表未找到 " + month + " 月 " + day + " 日"});

  var target = SpreadsheetApp.openById("1W55B3jsXN7KDz8Op0QEIjrgYeyMgnPQhUS1NfNoKCOo").getSheetByName("A69");
  if (!target) return responseJSON({success: false, msg: "目标表未找到 A69 页签"});
  var targetDays = target.getRange(2, 4, 1, 30).getValues()[0];
  var targetColumn = 0;
  for (var c = 0; c < targetDays.length; c++) {
    if (Number(targetDays[c]) === day) {
      targetColumn = c + 4;
      break;
    }
  }
  if (!targetColumn) return responseJSON({success: false, msg: "A69 页签未找到 " + day + " 日列"});

  var targetPeople = target.getRange(3, 1, Math.max(1, target.getLastRow() - 2), 2).getDisplayValues();
  var summaryIndex = targetPeople.findIndex(function(person) {
    return person[0].trim() === "汇总";
  });
  if (summaryIndex < 0) return responseJSON({success: false, msg: "A69 未找到人员区下方的汇总行"});
  targetPeople = targetPeople.slice(0, summaryIndex);
  var sourceRows = [3, 24, 45];
  var sections = sourceRows.map(function(start) {
    var people = source.getRange(start, 2, 17, 2).getDisplayValues();
    var counts = source.getRange(start, sourceColumn, 17, 1).getValues();
    var byEmployeeId = Object.create(null);
    people.forEach(function(person, i) {
      var id = person[1].trim();
      if (id) {
        byEmployeeId[id] = Object.prototype.hasOwnProperty.call(byEmployeeId, id)
          ? null
          : {name: person[0].trim().toUpperCase(), count: counts[i][0], row: start + i};
      }
    });
    return byEmployeeId;
  });
  var batches = [];
  var seen = Object.create(null);
  var updated = 0;
  for (var i = 0; i < targetPeople.length; i++) {
    var row = i + 3;
    var name = targetPeople[i][0].trim().toUpperCase();
    var employeeId = targetPeople[i][1].trim();
    if (!name && !employeeId) continue;
    if (!name || !employeeId || seen[employeeId]) {
      return responseJSON({success: false, msg: "A69 第 " + row + " 行缺少姓名、工号或工号重复"});
    }
    seen[employeeId] = true;
    var total = 0;
    for (var j = 0; j < sections.length; j++) {
      var person = sections[j][employeeId];
      if (!person || person.name !== name) {
        return responseJSON({success: false, msg: "A69 第 " + row + " 行工号 " + employeeId + " 在来源表第 " + (j + 1) + " 分区找不到唯一且姓名一致的记录"});
      }
      var count = person.count;
      if (count !== "" && (typeof count !== "number" || !isFinite(count))) {
        return responseJSON({success: false, msg: "来源表第 " + person.row + " 行工作量不是数字"});
      }
      total += count === "" ? 0 : count;
    }
    var batch = batches[batches.length - 1];
    if (!batch || batch.row + batch.values.length !== row) {
      batch = {row: row, values: []};
      batches.push(batch);
    }
    batch.values.push([total]);
    updated++;
  }
  if (!updated) return responseJSON({success: false, msg: "A69 没有可同步的人员"});
  batches.forEach(function(batch) {
    target.getRange(batch.row, targetColumn, batch.values.length, 1).setValues(batch.values);
  });
  return responseJSON({success: true, msg: "A69 " + month + " 月 " + day + " 日已更新 " + updated + " 人"});
}

function getSmartSheetInfo(dayInput, monthInput, yearInput) {
  var now = new Date();
  var fullYear = yearInput ? parseInt(yearInput, 10) : now.getFullYear();
  var month = monthInput ? parseInt(monthInput, 10) : (now.getMonth() + 1);

  if (!monthInput && dayInput) {
    var currentDay = now.getDate();
    var inputDay = parseInt(dayInput, 10);
    if (currentDay <= 5 && inputDay >= 20) {
      now.setMonth(now.getMonth() - 1);
      fullYear = now.getFullYear();
      month = now.getMonth() + 1;
    }
  }

  var yearShort = fullYear.toString().slice(-2);
  var sheetName = "工作量统计（" + yearShort + "年" + month + "月份）";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  return {
    sheet: sheet,
    name: sheetName,
    year: fullYear,
    month: month
  };
}

function handleGetKeywords(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return responseJSON({ success: true, keywords: [] });

  // C 列的处理量小计标记第一段结束；人员增减时行号会变化。
  var range = sheet.getRange(3, 3, lastRow - 2, 2);
  var values = range.getValues();
  var keywords = [];

  for (var i = 0; i < values.length; i++) {
    var sectionLabel = String(values[i][0] || "").trim();
    if (sectionLabel.indexOf("处理量") > -1 || sectionLabel.indexOf("总处理") > -1) break;
    var val = String(values[i][1] || "").trim();
    if (val === "姓名") break;
    if (val.indexOf("处理量") > -1 || val.indexOf("总处理") > -1) break;
    if (val === "稍等快捷") continue;

    if (val && keywords.indexOf(val) === -1) {
      keywords.push(val);
    }
  }
  return responseJSON({ success: true, keywords: keywords });
}

function handleWriteStats(sheet, day, stats, targetYear, targetMonth) {
  var colIndex = findDateColumn(sheet, day, targetYear, targetMonth);

  if (colIndex === 0) {
    return responseJSON({
      success: false,
      msg: "在表 [" + sheet.getName() + "] 中未找到日期 [" + targetMonth + "月" + parseInt(day, 10) + "日]"
    });
  }

  var keywordValues = sheet.getRange("D:D").getValues();
  var updates = 0;

  for (var key in stats) {
    var promoCount = stats[key].promo || 0;
    var assistCount = stats[key].assist || 0;
    var foundRows = [];
    for (var i = 0; i < keywordValues.length; i++) {
      if (keywordValues[i][0].toString().trim() === key.trim()) foundRows.push(i + 1);
    }
    if (foundRows.length >= 1) {
      var promoCell = sheet.getRange(foundRows[0], colIndex);
      promoCell.setValue(promoCount);
      fixCellStyle(promoCell);
      updates++;
    }
    if (foundRows.length >= 2) {
      var assistCell = sheet.getRange(foundRows[1], colIndex);
      assistCell.setValue(assistCount);
      fixCellStyle(assistCell);
      updates++;
    }
  }

  return responseJSON({ success: true, msg: "成功更新 " + updates + " 条数据" });
}

function handleWriteManualWorkload(sheet, params, targetYear, targetMonth) {
  var worker = (params.worker || params.name || "ARATAKITO").toString().trim();
  var occurrence = parseInt(params.row_occurrence || 3, 10);
  var total = params.total;
  if (total === undefined || total === null || total === "") total = params.quantity;

  if (!worker) {
    return responseJSON({ success: false, msg: "缺少 worker/name" });
  }
  if (!occurrence || occurrence < 1) {
    occurrence = 3;
  }
  if (total === undefined || total === null || total === "") {
    return responseJSON({ success: false, msg: "缺少 total/quantity" });
  }

  var day = params.day;
  if (!day && params.date) {
    var dateParts = params.date.toString().split("-");
    day = dateParts.length >= 3 ? dateParts[2] : "";
  }
  var colIndex = findDateColumn(sheet, day, targetYear, targetMonth);
  if (colIndex === 0) {
    return responseJSON({
      success: false,
      msg: "在表 [" + sheet.getName() + "] 中未找到日期 [" + targetMonth + "月" + parseInt(day, 10) + "日]"
    });
  }

  var lastRow = sheet.getLastRow();
  var nameValues = sheet.getRange(1, 4, lastRow, 1).getValues();
  var foundRows = [];
  var targetName = worker.toUpperCase();

  for (var i = 0; i < nameValues.length; i++) {
    var name = nameValues[i][0].toString().trim();
    if (name.toUpperCase() === targetName) {
      foundRows.push(i + 1);
    }
  }

  if (foundRows.length < occurrence) {
    return responseJSON({
      success: false,
      msg: "在 D 列只找到 " + foundRows.length + " 个 [" + worker + "]，不足第 " + occurrence + " 个"
    });
  }

  var rowIndex = foundRows[occurrence - 1];
  var cell = sheet.getRange(rowIndex, colIndex);
  cell.setValue(total);
  fixCellStyle(cell);

  return responseJSON({
    success: true,
    msg: worker + " 工作量已写入 " + cell.getA1Notation() + " = " + total
  });
}

function findDateColumn(sheet, day, targetYear, targetMonth) {
  var dateRowIndex = 2;
  var range = sheet.getRange(dateRowIndex, 1, 1, sheet.getLastColumn());
  var displayValues = range.getDisplayValues()[0];
  var rawValues = range.getValues()[0];
  var targetDay = parseInt(day, 10);
  if (!targetDay) return 0;

  var month = parseInt(targetMonth, 10);
  var year = parseInt(targetYear, 10);
  var paddedMonth = month < 10 ? "0" + month : "" + month;
  var paddedDay = targetDay < 10 ? "0" + targetDay : "" + targetDay;
  var targets = [
    month + "月" + targetDay + "日",
    year + "-" + month + "-" + targetDay,
    year + "-" + paddedMonth + "-" + paddedDay,
    year + "/" + month + "/" + targetDay,
    year + "/" + paddedMonth + "/" + paddedDay,
    month + "/" + targetDay,
    paddedMonth + "/" + paddedDay,
    month + "." + targetDay,
    paddedMonth + "." + paddedDay
  ];

  for (var i = 0; i < displayValues.length; i++) {
    var raw = rawValues[i];
    if (raw instanceof Date) {
      if (raw.getFullYear() === year && raw.getMonth() + 1 === month && raw.getDate() === targetDay) {
        return i + 1;
      }
    }
    var display = displayValues[i].toString().trim();
    if (targets.indexOf(display) > -1) {
      return i + 1;
    }
  }
  return 0;
}

function fixCellStyle(cell) {
  cell.setHorizontalAlignment("center");
  cell.setVerticalAlignment("middle");
  cell.setFontSize(9);
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}


// Set SHEET_MONITOR_SECRET in Script Properties and the bot environment to the
// same random value. Keep this secret out of source code and admin config exports.
function handleGetSheetSnapshot(params) {
  var secret = PropertiesService.getScriptProperties().getProperty("SHEET_MONITOR_SECRET");
  if (!secret || String(params.monitor_secret || "") !== secret) {
    return responseJSON({ success: false, code: "MONITOR_AUTH" });
  }
  var id = String(params.spreadsheet_id || "");
  if (!/^[A-Za-z0-9_-]+$/.test(id)) {
    return responseJSON({ success: false, code: "MONITOR_ACCESS" });
  }
  try {
    var book = SpreadsheetApp.openById(id);
    var snapshot = Object.create(null);
    var sheetIds = Object.create(null);
    book.getSheets().forEach(function(sheet) {
      var cells = Object.create(null);
      if (sheet.getLastRow() && sheet.getLastColumn()) {
        var range = sheet.getDataRange();
        var values = range.getDisplayValues();
        var formulas = range.getFormulas();
        for (var r = 0; r < values.length; r++) {
          for (var c = 0; c < values[r].length; c++) {
            var value = values[r][c];
            var formula = formulas[r][c];
            if (value !== "" || formula !== "") {
              var column = "";
              for (var n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) {
                column = String.fromCharCode(65 + (n - 1) % 26) + column;
              }
              cells[column + (r + 1)] = [formula || value, value];
            }
          }
        }
      }
      snapshot[sheet.getName()] = cells;
      sheetIds[sheet.getName()] = sheet.getSheetId();
    });
    return responseJSON({success: true, version: 1, spreadsheet_id: book.getId(), snapshot: snapshot, sheet_ids: sheetIds});
  } catch (err) {
    // Do not expose private document metadata through an anonymous web endpoint.
    return responseJSON({success: false, code: "MONITOR_ACCESS"});
  }
}
