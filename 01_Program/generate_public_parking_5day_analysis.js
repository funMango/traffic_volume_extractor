const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const ROOT_DIR = path.resolve(__dirname, "..");
const INPUT_FILE = path.join(ROOT_DIR, "00_Data", "공영주차장 데이터.xlsx");
const OUTPUT_DIR = path.join(ROOT_DIR, "02_Result", "09_기타");
const MAIN_OUTPUT_FILE = path.join(OUTPUT_DIR, "공영주차장_차량5부제_대상차량_이용분석.html");
const DETAIL_OUTPUT_FILE = path.join(OUTPUT_DIR, "공영주차장_차량5부제_상세내역.html");
const PARKING_TARGET_XLSX_FILE = path.join(
  OUTPUT_DIR,
  "공영주차장_차량5부제_주차장별_타겟차량현황.xlsx",
);
const SUMMARY_DETAIL_XLSX_FILE = path.join(
  OUTPUT_DIR,
  "공영주차장_차량5부제_일자별_주차장별_요약_및_상세이용기록.xlsx",
);
const MAIN_OUTPUT_BASENAME = path.basename(MAIN_OUTPUT_FILE);
const DETAIL_OUTPUT_BASENAME = path.basename(DETAIL_OUTPUT_FILE);

const SHEET_NAME = "Sheet1";
const OUTPUT_SHEET_NAME = "Sheet1";
const SOURCE_COLUMNS = ["입차시간", "출차시간", "차량번호", "주차장명"];
const ANALYSIS_START = "2026-04-08";
const ANALYSIS_END = "2026-04-21";

const PARKING_HOURS = {
  구버스터미널일원: { start: "09:00", end: "18:00" },
  로데오거리: { start: "09:00", end: "18:00" },
  롯데백화점일원: { start: "09:00", end: "18:00" },
  중동먹거리: { start: "09:00", end: "18:00" },
  중동먹자골목: { start: "09:00", end: "18:00" },
  도당어울마당: { start: "08:00", end: "15:00" },
  원미공원1: { start: "10:00", end: "18:00" },
  원미공원2: { start: "10:00", end: "18:00" },
  원미철골: { start: "10:00", end: "18:00" },
};

const TARGET_DIGITS_BY_DAY = {
  1: ["1", "6"],
  2: ["2", "7"],
  3: ["3", "8"],
  4: ["4", "9"],
  5: ["5", "0"],
};

const WEEKDAY_KO = ["일", "월", "화", "수", "목", "금", "토"];
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const MS_PER_MINUTE = 60 * 1000;

function pad2(value) {
  return String(value).padStart(2, "0");
}

function dateKeyFromUtcDate(date) {
  return `${date.getUTCFullYear()}-${pad2(date.getUTCMonth() + 1)}-${pad2(date.getUTCDate())}`;
}

function parseDateKey(dateKey) {
  const [year, month, day] = dateKey.split("-").map(Number);
  return { year, month, day };
}

function utcMsFromDateKey(dateKey) {
  const { year, month, day } = parseDateKey(dateKey);
  return Date.UTC(year, month - 1, day);
}

function getWeekday(dateKey) {
  return new Date(utcMsFromDateKey(dateKey)).getUTCDay();
}

function getAnalysisDates() {
  const dates = [];
  for (
    let ms = utcMsFromDateKey(ANALYSIS_START);
    ms <= utcMsFromDateKey(ANALYSIS_END);
    ms += MS_PER_DAY
  ) {
    const date = new Date(ms);
    const day = date.getUTCDay();
    if (day >= 1 && day <= 5) {
      dates.push(dateKeyFromUtcDate(date));
    }
  }
  return dates;
}

function timeTextToMinutes(timeText) {
  const [hour, minute] = timeText.split(":").map(Number);
  return hour * 60 + minute;
}

function getOperationWindow(dateKey, hours) {
  const startOfDay = utcMsFromDateKey(dateKey);
  return {
    startMs: startOfDay + timeTextToMinutes(hours.start) * MS_PER_MINUTE,
    endMs: startOfDay + timeTextToMinutes(hours.end) * MS_PER_MINUTE,
  };
}

function excelSerialToDateTimeParts(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return {
      year: value.getFullYear(),
      month: value.getMonth() + 1,
      day: value.getDate(),
      hour: value.getHours(),
      minute: value.getMinutes(),
      second: value.getSeconds(),
    };
  }

  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`엑셀 날짜 serial 값이 아닙니다: ${value}`);
  }

  let wholeDays = Math.floor(value);
  const fraction = value - wholeDays;
  let totalSeconds = Math.round(fraction * 24 * 60 * 60);

  if (totalSeconds >= 24 * 60 * 60) {
    wholeDays += Math.floor(totalSeconds / (24 * 60 * 60));
    totalSeconds %= 24 * 60 * 60;
  }

  const date = new Date(Date.UTC(1899, 11, 30 + wholeDays));
  const hour = Math.floor(totalSeconds / 3600);
  const minute = Math.floor((totalSeconds % 3600) / 60);
  const second = totalSeconds % 60;

  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hour,
    minute,
    second,
  };
}

function dateTimePartsToMs(parts) {
  return Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second);
}

function dateKeyFromParts(parts) {
  return `${parts.year}-${pad2(parts.month)}-${pad2(parts.day)}`;
}

function formatDateTime(parts) {
  const seconds = parts.second ? `:${pad2(parts.second)}` : "";
  return `${dateKeyFromParts(parts)} ${pad2(parts.hour)}:${pad2(parts.minute)}${seconds}`;
}

function formatOperation(hours) {
  return `${hours.start} ~ ${hours.end}`;
}

function hasValue(value) {
  return value !== null && value !== undefined && String(value).trim() !== "";
}

function getLastDigit(vehicleNumber) {
  const text = String(vehicleNumber || "");
  for (let index = text.length - 1; index >= 0; index -= 1) {
    const char = text[index];
    if (char >= "0" && char <= "9") {
      return char;
    }
  }
  return null;
}

function classifyVehicle(vehicleNumber, weekday) {
  const digit = getLastDigit(vehicleNumber);
  if (!digit) {
    return {
      digit: "판단불가",
      status: "판단불가",
      isTarget: false,
    };
  }

  const targetDigits = TARGET_DIGITS_BY_DAY[weekday] || [];
  const isTarget = targetDigits.includes(digit);
  return {
    digit,
    status: isTarget ? "대상" : "비대상",
    isTarget,
  };
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatNumber(value) {
  return new Intl.NumberFormat("ko-KR").format(value);
}

function formatAverage(value) {
  return Number(value || 0).toFixed(1);
}

function formatPercent(value) {
  return `${value.toFixed(2)}%`;
}

function serializeJsonForScript(value) {
  return JSON.stringify(value)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
}

function readSourceRows() {
  const workbook = XLSX.readFile(INPUT_FILE, { cellDates: false });
  if (!workbook.Sheets[SHEET_NAME]) {
    throw new Error(`${SHEET_NAME} 시트를 찾을 수 없습니다.`);
  }

  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[SHEET_NAME], {
    header: 1,
    raw: true,
    defval: null,
    blankrows: false,
  });

  if (rows.length < 2) {
    throw new Error(`${SHEET_NAME}에 분석할 데이터가 없습니다.`);
  }

  const headers = rows[0].slice(0, SOURCE_COLUMNS.length);
  SOURCE_COLUMNS.forEach((expected, index) => {
    if (headers[index] !== expected) {
      throw new Error(`원천 컬럼 ${index + 1}번이 '${expected}'가 아닙니다. 실제값: '${headers[index]}'`);
    }
  });

  return rows.slice(1);
}

function buildGroups(analysisDates) {
  const groups = new Map();
  const parkingNames = Object.keys(PARKING_HOURS);

  analysisDates.forEach((dateKey) => {
    parkingNames.forEach((parkingName) => {
      const key = `${dateKey}|${parkingName}`;
      groups.set(key, {
        key,
        dateKey,
        weekday: getWeekday(dateKey),
        parkingName,
        operation: formatOperation(PARKING_HOURS[parkingName]),
        vehicles: new Map(),
        detailRows: [],
        counts: {
          total: 0,
          target: 0,
          nonTarget: 0,
          unknown: 0,
        },
      });
    });
  });

  return groups;
}

function createDetailRow(sourceRowNumber, dateKey, parkingName, entryParts, exitParts, vehicleNumber, missingExit) {
  const weekday = getWeekday(dateKey);
  const classification = classifyVehicle(vehicleNumber, weekday);
  const hours = PARKING_HOURS[parkingName];

  return {
    sourceRowNumber,
    dateKey,
    weekdayLabel: WEEKDAY_KO[weekday],
    parkingName,
    operation: formatOperation(hours),
    entryTime: formatDateTime(entryParts),
    exitTime: exitParts ? formatDateTime(exitParts) : "출차시간 누락",
    vehicleNumber: String(vehicleNumber || "").trim(),
    lastDigit: classification.digit,
    inOperation: "예",
    targetStatus: classification.status,
    missingExit,
  };
}

function addRecordToGroup(group, detailRow) {
  const vehicleKey = detailRow.vehicleNumber || `빈차량번호:${detailRow.sourceRowNumber}`;

  group.detailRows.push(detailRow);
  if (!group.vehicles.has(vehicleKey)) {
    const weekday = getWeekday(group.dateKey);
    const classification = classifyVehicle(detailRow.vehicleNumber, weekday);

    group.vehicles.set(vehicleKey, classification.status);
    group.counts.total += 1;
    if (classification.status === "대상") {
      group.counts.target += 1;
    } else if (classification.status === "판단불가") {
      group.counts.unknown += 1;
    } else {
      group.counts.nonTarget += 1;
    }
  }
}

function overlapsOperation(entryMs, exitMs, opStartMs, opEndMs) {
  return entryMs < opEndMs && exitMs > opStartMs;
}

function analyzeRows(sourceRows, analysisDates) {
  const groups = buildGroups(analysisDates);
  const parkingNames = new Set(Object.keys(PARKING_HOURS));
  const unmatchedParkingNames = new Map();
  const analysisDateSet = new Set(analysisDates);
  const stats = {
    sourceRows: sourceRows.length,
    missingExitRows: 0,
    blankVehicleRows: 0,
    invalidExitOrderRows: 0,
    includedDetailRows: 0,
  };

  sourceRows.forEach((row, index) => {
    const sourceRowNumber = index + 2;
    const entryRaw = row[0];
    const exitRaw = row[1];
    const vehicleNumber = hasValue(row[2]) ? String(row[2]).trim() : "";
    const parkingName = hasValue(row[3]) ? String(row[3]).trim() : "";

    if (!parkingNames.has(parkingName)) {
      unmatchedParkingNames.set(parkingName || "(빈 주차장명)", (unmatchedParkingNames.get(parkingName || "(빈 주차장명)") || 0) + 1);
      return;
    }

    if (!vehicleNumber) {
      stats.blankVehicleRows += 1;
    }

    const entryParts = excelSerialToDateTimeParts(entryRaw);
    const entryMs = dateTimePartsToMs(entryParts);
    const exitExists = hasValue(exitRaw);
    let exitParts = null;
    let exitMs = null;

    if (exitExists) {
      exitParts = excelSerialToDateTimeParts(exitRaw);
      exitMs = dateTimePartsToMs(exitParts);
      if (exitMs <= entryMs) {
        stats.invalidExitOrderRows += 1;
        return;
      }
    } else {
      stats.missingExitRows += 1;
    }

    if (!exitExists) {
      const entryDateKey = dateKeyFromParts(entryParts);
      if (!analysisDateSet.has(entryDateKey)) {
        return;
      }

      const group = groups.get(`${entryDateKey}|${parkingName}`);
      const { startMs, endMs } = getOperationWindow(entryDateKey, PARKING_HOURS[parkingName]);
      const virtualExitMs = utcMsFromDateKey(entryDateKey) + MS_PER_DAY;

      if (overlapsOperation(entryMs, virtualExitMs, startMs, endMs)) {
        const detailRow = createDetailRow(sourceRowNumber, entryDateKey, parkingName, entryParts, null, vehicleNumber, true);
        addRecordToGroup(group, detailRow);
        stats.includedDetailRows += 1;
      }
      return;
    }

    analysisDates.forEach((dateKey) => {
      const { startMs, endMs } = getOperationWindow(dateKey, PARKING_HOURS[parkingName]);
      if (!overlapsOperation(entryMs, exitMs, startMs, endMs)) {
        return;
      }

      const group = groups.get(`${dateKey}|${parkingName}`);
      const detailRow = createDetailRow(sourceRowNumber, dateKey, parkingName, entryParts, exitParts, vehicleNumber, false);
      addRecordToGroup(group, detailRow);
      stats.includedDetailRows += 1;
    });
  });

  if (unmatchedParkingNames.size > 0) {
    const lines = Array.from(unmatchedParkingNames.entries())
      .map(([name, count]) => `${name}: ${count}`)
      .join(", ");
    throw new Error(`운영시간 기준표와 매칭되지 않는 주차장명이 있습니다. ${lines}`);
  }

  return { groups, stats };
}

function validateAnalysis(groups, analysisDates) {
  const expectedSummaryRows = analysisDates.length * Object.keys(PARKING_HOURS).length;
  if (groups.size !== expectedSummaryRows) {
    throw new Error(`요약 행 수가 예상과 다릅니다. 예상 ${expectedSummaryRows}, 실제 ${groups.size}`);
  }

  const failures = [];
  groups.forEach((group) => {
    const uniqueVehicles = new Set(group.detailRows.map((row) => row.vehicleNumber || `빈차량번호:${row.sourceRowNumber}`));
    if (uniqueVehicles.size !== group.counts.total) {
      failures.push(`${group.dateKey} ${group.parkingName}: 전체 차량 대수 ${group.counts.total}, 상세 고유 차량 ${uniqueVehicles.size}`);
    }

    const subtotal = group.counts.target + group.counts.nonTarget + group.counts.unknown;
    if (subtotal !== group.counts.total) {
      failures.push(`${group.dateKey} ${group.parkingName}: 대상+비대상+판단불가 ${subtotal}, 전체 ${group.counts.total}`);
    }
  });

  if (failures.length > 0) {
    throw new Error(`집계 검증 실패:\n${failures.slice(0, 10).join("\n")}`);
  }
}

function buildSummaryRows(groups) {
  return Array.from(groups.values()).map((group, index) => {
    const ratio = group.counts.total ? (group.counts.target / group.counts.total) * 100 : 0;
    const detailRows = group.detailRows.sort((a, b) => {
      if (a.entryTime !== b.entryTime) return a.entryTime.localeCompare(b.entryTime);
      return a.vehicleNumber.localeCompare(b.vehicleNumber, "ko-KR");
    });

    return {
      id: `summary-${index + 1}`,
      dateKey: group.dateKey,
      weekdayLabel: WEEKDAY_KO[group.weekday],
      parkingName: group.parkingName,
      operation: group.operation,
      total: group.counts.total,
      target: group.counts.target,
      nonTarget: group.counts.nonTarget,
      unknown: group.counts.unknown,
      ratio,
      detailUrl: `${DETAIL_OUTPUT_BASENAME}?date=${encodeURIComponent(group.dateKey)}&parking=${encodeURIComponent(group.parkingName)}`,
      detailRows,
    };
  });
}

function buildParkingRows(summaryRows, analysisDates) {
  const summaryByParkingDate = new Map(
    summaryRows.map((row) => [`${row.parkingName}|${row.dateKey}`, row]),
  );
  const analysisDateCount = analysisDates.length;

  return Object.keys(PARKING_HOURS).map((parkingName, index) => {
    const dailyRows = analysisDates.map((dateKey) => {
      const row = summaryByParkingDate.get(`${parkingName}|${dateKey}`);
      if (!row) {
        throw new Error(`주차장별 집계에 필요한 요약 행이 없습니다. ${dateKey} ${parkingName}`);
      }

      return {
        dateKey: row.dateKey,
        weekdayLabel: row.weekdayLabel,
        total: row.total,
        target: row.target,
        ratio: row.total ? (row.target / row.total) * 100 : 0,
      };
    });

    const total = dailyRows.reduce((sum, row) => sum + row.total, 0);
    const target = dailyRows.reduce((sum, row) => sum + row.target, 0);
    const averageTotal = analysisDateCount ? total / analysisDateCount : 0;
    const averageTarget = analysisDateCount ? target / analysisDateCount : 0;
    const averageRatio = averageTotal ? (averageTarget / averageTotal) * 100 : 0;

    return {
      id: `parking-${index + 1}`,
      parkingName,
      total,
      target,
      averageTotal,
      averageTarget,
      averageRatio,
      analysisDateCount,
      detailUrl: `${DETAIL_OUTPUT_BASENAME}?parking=${encodeURIComponent(parkingName)}`,
      dailyRows,
    };
  });
}

function buildMetadata(summaryRows, stats, analysisDates) {
  const totalVehicles = summaryRows.reduce((sum, row) => sum + row.total, 0);
  const targetVehicles = summaryRows.reduce((sum, row) => sum + row.target, 0);
  const unknownVehicles = summaryRows.reduce((sum, row) => sum + row.unknown, 0);
  const detailRows = summaryRows.reduce((sum, row) => sum + row.detailRows.length, 0);

  return {
    totalVehicles,
    targetVehicles,
    unknownVehicles,
    detailRows,
    targetRatio: totalVehicles ? (targetVehicles / totalVehicles) * 100 : 0,
    analysisDateCount: analysisDates.length,
    sourceRows: stats.sourceRows,
    missingExitRows: stats.missingExitRows,
    blankVehicleRows: stats.blankVehicleRows,
    invalidExitOrderRows: stats.invalidExitOrderRows,
  };
}

function renderParkingRows(parkingRows) {
  return parkingRows
    .map((row) => {
      return `
        <tr class="summary-row" tabindex="0" role="link" aria-label="${escapeHtml(`${row.parkingName} 상세 내역으로 이동`)}">
          <td><a class="summary-link" href="${escapeHtml(row.detailUrl)}">${escapeHtml(row.parkingName)}</a></td>
          <td class="numeric">${formatAverage(row.averageTotal)}</td>
          <td class="numeric emphasis">${formatAverage(row.averageTarget)}</td>
          <td class="numeric">${formatPercent(row.averageRatio)}</td>
        </tr>`;
    })
    .join("");
}

function renderSummaryRows(summaryRows) {
  return summaryRows
    .map((row) => {
      const detailCountText = row.detailRows.length > 0 ? `${formatNumber(row.detailRows.length)}건` : "0건";
      return `
        <tr class="summary-row" tabindex="0" role="link" aria-label="${escapeHtml(`${row.dateKey} ${row.parkingName} 상세 내역으로 이동`)}">
          <td><a class="summary-link" href="${escapeHtml(row.detailUrl)}">${escapeHtml(`${row.dateKey} (${row.weekdayLabel})`)}</a></td>
          <td>${escapeHtml(row.parkingName)}</td>
          <td>${escapeHtml(row.operation)}</td>
          <td class="numeric">${formatNumber(row.total)}</td>
          <td class="numeric emphasis">${formatNumber(row.target)}</td>
          <td class="numeric">${formatNumber(row.unknown)}</td>
          <td class="numeric">${formatPercent(row.ratio)}</td>
          <td class="numeric muted">${detailCountText}</td>
        </tr>`;
    })
    .join("");
}

function sanitizeWorksheetName(name) {
  const sanitized = String(name || "Sheet")
    .replace(/[\\/?*\[\]:]/g, "_")
    .slice(0, 31);
  return sanitized || "Sheet";
}

function getUniqueWorksheetName(name, usedNames) {
  const baseName = sanitizeWorksheetName(name);
  let worksheetName = baseName;
  let suffixNumber = 2;

  while (usedNames.has(worksheetName)) {
    const suffix = `_${suffixNumber}`;
    worksheetName = `${baseName.slice(0, 31 - suffix.length)}${suffix}`;
    suffixNumber += 1;
  }

  usedNames.add(worksheetName);
  return worksheetName;
}

function ratioToExcelPercent(value) {
  return Number(value || 0) / 100;
}

function applyColumnFormats(worksheet, rowCount, columnFormats) {
  columnFormats.forEach((numberFormat, columnIndex) => {
    if (!numberFormat) {
      return;
    }

    for (let rowIndex = 1; rowIndex <= rowCount; rowIndex += 1) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      if (worksheet[cellAddress]) {
        worksheet[cellAddress].z = numberFormat;
      }
    }
  });
}

function createWorksheet(headers, rows, columnFormats, columnWidths) {
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  worksheet["!cols"] = columnWidths.map((width) => ({ wch: width }));
  worksheet["!autofilter"] = {
    ref: XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: rows.length, c: headers.length - 1 },
    }),
  };

  applyColumnFormats(worksheet, rows.length, columnFormats);
  return worksheet;
}

function appendWorksheet(workbook, worksheet, requestedName, usedWorksheetNames) {
  const worksheetName = getUniqueWorksheetName(requestedName, usedWorksheetNames);
  XLSX.utils.book_append_sheet(workbook, worksheet, worksheetName);
}

function compareDetailRowsByDateAndEntry(a, b) {
  if (a.dateKey !== b.dateKey) {
    return a.dateKey.localeCompare(b.dateKey);
  }
  if (a.entryTime !== b.entryTime) {
    return a.entryTime.localeCompare(b.entryTime);
  }
  if (a.exitTime !== b.exitTime) {
    return a.exitTime.localeCompare(b.exitTime);
  }
  if (a.vehicleNumber !== b.vehicleNumber) {
    return a.vehicleNumber.localeCompare(b.vehicleNumber, "ko-KR");
  }
  return (a.sourceRowNumber || 0) - (b.sourceRowNumber || 0);
}

function buildParkingDailyWorksheet(parkingRow) {
  const headers = ["일자", "요일", "전체 차량 대수", "타겟 차량 대수", "타겟 차량 비율"];
  const rows = parkingRow.dailyRows.map((dailyRow) => [
    dailyRow.dateKey,
    dailyRow.weekdayLabel,
    dailyRow.total,
    dailyRow.target,
    ratioToExcelPercent(dailyRow.ratio),
  ]);

  return createWorksheet(headers, rows, [null, null, "#,##0", "#,##0", "0.00%"], [14, 8, 16, 16, 16]);
}

function writeParkingTargetWorkbook(parkingRows) {
  const workbook = XLSX.utils.book_new();
  const usedWorksheetNames = new Set();
  const headers = ["주차장명", "평균 전체 차량 대수", "평균 타겟 차량 대수", "평균 타겟 차량 비율"];
  const rows = parkingRows.map((row) => [
    row.parkingName,
    row.averageTotal,
    row.averageTarget,
    ratioToExcelPercent(row.averageRatio),
  ]);

  appendWorksheet(
    workbook,
    createWorksheet(headers, rows, [null, "#,##0.0", "#,##0.0", "0.00%"], [20, 20, 20, 20]),
    OUTPUT_SHEET_NAME,
    usedWorksheetNames,
  );

  parkingRows.forEach((parkingRow) => {
    appendWorksheet(workbook, buildParkingDailyWorksheet(parkingRow), parkingRow.parkingName, usedWorksheetNames);
  });

  XLSX.writeFile(workbook, PARKING_TARGET_XLSX_FILE);
}

function buildSummaryWorksheet(summaryRows) {
  const headers = [
    "일자",
    "요일",
    "주차장명",
    "운영시간",
    "전체 차량 대수",
    "타겟 차량 대수",
    "판단불가 차량 대수",
    "타겟 차량 비율",
    "상세 기록 수",
  ];
  const rows = summaryRows.map((row) => [
    row.dateKey,
    row.weekdayLabel,
    row.parkingName,
    row.operation,
    row.total,
    row.target,
    row.unknown,
    ratioToExcelPercent(row.ratio),
    row.detailRows.length,
  ]);

  return createWorksheet(
    headers,
    rows,
    [null, null, null, null, "#,##0", "#,##0", "#,##0", "0.00%", "#,##0"],
    [14, 8, 18, 16, 16, 16, 18, 16, 14],
  );
}

function buildParkingDetailWorksheet(summaryRows, parkingName) {
  const headers = [
    "일자",
    "주차장명",
    "운영시간",
    "입차시간",
    "출차시간",
    "차량번호",
    "차량번호 끝자리",
    "운영시간 내 이용 여부",
    "타겟 여부",
  ];
  const detailRows = summaryRows
    .filter((summaryRow) => summaryRow.parkingName === parkingName)
    .flatMap((summaryRow) => summaryRow.detailRows)
    .sort(compareDetailRowsByDateAndEntry);
  const rows = detailRows.map((detailRow) => [
    detailRow.dateKey,
    detailRow.parkingName,
    detailRow.operation,
    detailRow.entryTime,
    detailRow.exitTime,
    detailRow.vehicleNumber || "빈 차량번호",
    detailRow.lastDigit,
    detailRow.inOperation,
    detailRow.targetStatus,
  ]);

  return createWorksheet(headers, rows, [null, null, null, null, null, null, null, null, null], [14, 18, 16, 20, 20, 16, 16, 20, 12]);
}

function writeSummaryDetailWorkbook(summaryRows, parkingRows) {
  const workbook = XLSX.utils.book_new();
  const usedWorksheetNames = new Set();

  appendWorksheet(workbook, buildSummaryWorksheet(summaryRows), OUTPUT_SHEET_NAME, usedWorksheetNames);

  parkingRows.forEach((parkingRow) => {
    appendWorksheet(
      workbook,
      buildParkingDetailWorksheet(summaryRows, parkingRow.parkingName),
      parkingRow.parkingName,
      usedWorksheetNames,
    );
  });

  XLSX.writeFile(workbook, SUMMARY_DETAIL_XLSX_FILE);
}

function writeExcelOutputs(parkingRows, summaryRows) {
  writeParkingTargetWorkbook(parkingRows);
  writeSummaryDetailWorkbook(summaryRows, parkingRows);
}

function renderHtml(parkingRows, summaryRows, metadata, analysisDates) {
  const generatedAt = new Date().toLocaleString("ko-KR", { timeZone: "Asia/Seoul" });

  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>공영주차장 차량 5부제 대상차량 이용분석</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f8fb;
      --surface: #ffffff;
      --line: #d8dee8;
      --line-strong: #aeb9c9;
      --text: #18202c;
      --muted: #5b6678;
      --accent: #176b87;
      --accent-soft: #e7f3f7;
      --target: #b42318;
      --target-soft: #fff0ee;
      --unknown: #805300;
      --unknown-soft: #fff6d8;
      --ok: #146c43;
      --ok-soft: #e9f7ef;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: "Malgun Gothic", "Apple SD Gothic Neo", Arial, sans-serif;
      font-size: 14px;
      line-height: 1.5;
    }

    header {
      background: #102331;
      color: #ffffff;
      padding: 28px 32px 24px;
      border-bottom: 1px solid #07131c;
    }

    main {
      padding: 24px 32px 40px;
    }

    h1 {
      margin: 0 0 10px;
      font-size: 28px;
      font-weight: 700;
      letter-spacing: 0;
    }

    h2 {
      margin: 0 0 14px;
      font-size: 18px;
      font-weight: 700;
      letter-spacing: 0;
    }

    p {
      margin: 0;
    }

    .subhead {
      color: #d8e3ec;
      font-size: 15px;
    }

    .section {
      max-width: 1600px;
      margin: 0 auto 22px;
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
    }

    .section-body {
      padding: 18px 20px;
    }

    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(5, minmax(140px, 1fr));
      gap: 1px;
      background: var(--line);
      border-top: 1px solid var(--line);
    }

    .kpi {
      background: var(--surface);
      padding: 16px 18px;
    }

    .kpi-label {
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }

    .kpi-value {
      margin-top: 6px;
      font-size: 24px;
      font-weight: 700;
      letter-spacing: 0;
    }

    .info-grid {
      display: grid;
      grid-template-columns: 1.2fr 1fr;
      gap: 18px;
    }

    .info-block {
      min-width: 0;
    }

    .note {
      background: #fff8e6;
      border: 1px solid #f0d58a;
      color: #4d3a00;
      padding: 12px 14px;
      border-radius: 8px;
      margin-top: 12px;
    }

    .tag-list {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 10px;
    }

    .tag {
      display: inline-flex;
      align-items: center;
      min-height: 28px;
      padding: 4px 9px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #f9fbfd;
      font-weight: 700;
      white-space: nowrap;
    }

    table {
      width: 100%;
      border-collapse: collapse;
    }

    th,
    td {
      border-bottom: 1px solid var(--line);
      padding: 10px 12px;
      text-align: left;
      vertical-align: middle;
      word-break: keep-all;
    }

    thead th {
      background: #eef2f6;
      color: #263244;
      font-size: 12px;
      font-weight: 700;
      position: sticky;
      top: 0;
      z-index: 1;
    }

    .summary-table th,
    .summary-table td {
      white-space: nowrap;
    }

    .summary-row {
      cursor: pointer;
      background: #ffffff;
    }

    .summary-row:hover,
    .summary-row:focus,
    .summary-row:focus-within {
      background: var(--accent-soft);
      outline: none;
    }

    .summary-link {
      color: var(--accent);
      font-weight: 700;
      text-decoration: none;
    }

    .summary-link:hover,
    .summary-link:focus {
      text-decoration: underline;
    }

    .numeric {
      text-align: right;
      font-variant-numeric: tabular-nums;
    }

    .emphasis {
      color: var(--target);
      font-weight: 700;
    }

    .muted {
      color: var(--muted);
    }

    .status {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 54px;
      min-height: 24px;
      padding: 2px 8px;
      border-radius: 6px;
      font-weight: 700;
      white-space: nowrap;
    }

    .status-yes {
      background: var(--ok-soft);
      color: var(--ok);
    }

    .status-target {
      background: var(--target-soft);
      color: var(--target);
    }

    .status-unknown {
      background: var(--unknown-soft);
      color: var(--unknown);
    }

    .status-normal {
      background: #edf1f5;
      color: #394455;
    }

    .compact-table th,
    .compact-table td {
      padding: 8px 10px;
    }

    .empty {
      color: var(--muted);
      text-align: center;
      padding: 24px;
    }

    .table-wrap {
      overflow: auto;
    }

    @media (max-width: 980px) {
      header,
      main {
        padding-left: 16px;
        padding-right: 16px;
      }

      .kpi-grid,
      .info-grid {
        grid-template-columns: 1fr;
      }

      h1 {
        font-size: 23px;
      }
    }
  </style>
</head>
<body>
  <header>
    <h1>공영주차장 차량 5부제 대상차량 이용분석</h1>
    <p class="subhead">분석 기간: ${escapeHtml(ANALYSIS_START)} ~ ${escapeHtml(ANALYSIS_END)} 평일 ${metadata.analysisDateCount}일, 토요일·일요일 제외</p>
  </header>

  <main>
    <section class="section" aria-label="요약 지표">
      <div class="kpi-grid">
        <div class="kpi">
          <div class="kpi-label">요약 차량 대수</div>
          <div class="kpi-value">${formatNumber(metadata.totalVehicles)}</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">타겟 차량 대수</div>
          <div class="kpi-value">${formatNumber(metadata.targetVehicles)}</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">타겟 차량 비율</div>
          <div class="kpi-value">${formatPercent(metadata.targetRatio)}</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">판단불가 차량 대수</div>
          <div class="kpi-value">${formatNumber(metadata.unknownVehicles)}</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">상세 이용 기록</div>
          <div class="kpi-value">${formatNumber(metadata.detailRows)}</div>
        </div>
      </div>
    </section>

    <section class="section" aria-label="주차장별 타겟 차량 현황">
      <div class="section-body">
        <h2>주차장별 타겟 차량 현황</h2>
        <p class="muted">평균 차량 대수는 분석 평일 ${metadata.analysisDateCount}일을 분모로 산정했습니다.</p>
      </div>
      <div class="table-wrap">
        <table class="summary-table">
          <thead>
            <tr>
              <th>주차장명</th>
              <th class="numeric">평균 전체 차량 대수</th>
              <th class="numeric">평균 타겟 차량 대수</th>
              <th class="numeric">평균 타겟 차량 비율</th>
            </tr>
          </thead>
          <tbody>
            ${renderParkingRows(parkingRows)}
          </tbody>
        </table>
      </div>
    </section>

    <section class="section" aria-label="일자별 주차장별 요약">
      <div class="section-body">
        <h2>일자별·주차장별 요약</h2>
        <p class="muted">요약 차량 대수는 같은 일자·주차장·차량번호 조합을 1대로 집계했습니다.</p>
      </div>
      <div class="table-wrap">
        <table class="summary-table">
          <thead>
            <tr>
              <th>일자</th>
              <th>주차장명</th>
              <th>운영시간</th>
              <th class="numeric">전체 차량 대수</th>
              <th class="numeric">타겟 차량 대수</th>
              <th class="numeric">판단불가 차량 대수</th>
              <th class="numeric">타겟 차량 비율</th>
              <th class="numeric">상세 기록</th>
            </tr>
          </thead>
          <tbody>
            ${renderSummaryRows(summaryRows)}
          </tbody>
        </table>
      </div>
    </section>

    <section class="section" aria-label="데이터 품질 주석">
      <div class="section-body">
        <h2>데이터 품질 주석</h2>
        <p>원천 데이터 ${formatNumber(metadata.sourceRows)}행 중 출차시간 누락은 ${formatNumber(metadata.missingExitRows)}건, 차량번호 공백은 ${formatNumber(metadata.blankVehicleRows)}건, 출차시간이 입차시간 이하인 제외 행은 ${formatNumber(metadata.invalidExitOrderRows)}건입니다. 생성 시각: ${escapeHtml(generatedAt)}.</p>
        <div class="note">
          출차시간 누락 ${formatNumber(metadata.missingExitRows)}건은 입차일에만 반영했습니다. 누락 행은 입차시각부터 입차일 자정 전까지의 이용으로 보아, 해당 입차일 운영시간과 양의 시간 중첩이 있는 경우만 포함했고 이후 날짜로 확장하지 않았습니다.
        </div>
      </div>
    </section>
  </main>

  <script>
    document.querySelectorAll(".summary-row").forEach((row) => {
      const link = row.querySelector(".summary-link");
      const navigate = () => {
        if (link) {
          window.location.href = link.href;
        }
      };

      row.addEventListener("click", (event) => {
        if (event.target.closest("a")) {
          return;
        }
        navigate();
      });

      row.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          navigate();
        }
      });
    });
  </script>
</body>
</html>`;
}

function renderDetailHtml(parkingRows, summaryRows, metadata) {
  const generatedAt = new Date().toLocaleString("ko-KR", { timeZone: "Asia/Seoul" });
  const parkingDetailData = parkingRows.map((row) => ({
    parkingName: row.parkingName,
    total: row.total,
    target: row.target,
    averageTotal: row.averageTotal,
    averageTarget: row.averageTarget,
    averageRatio: row.averageRatio,
    analysisDateCount: row.analysisDateCount,
    dailyRows: row.dailyRows,
  }));
  const dateParkingDetailData = summaryRows.map((row) => ({
    dateKey: row.dateKey,
    weekdayLabel: row.weekdayLabel,
    parkingName: row.parkingName,
    operation: row.operation,
    total: row.total,
    target: row.target,
    nonTarget: row.nonTarget,
    unknown: row.unknown,
    ratio: row.ratio,
    detailRowCount: row.detailRows.length,
    detailRows: row.detailRows.map((detailRow) => ({
      dateKey: detailRow.dateKey,
      weekdayLabel: detailRow.weekdayLabel,
      parkingName: detailRow.parkingName,
      operation: detailRow.operation,
      entryTime: detailRow.entryTime,
      exitTime: detailRow.exitTime,
      vehicleNumber: detailRow.vehicleNumber,
      lastDigit: detailRow.lastDigit,
      inOperation: detailRow.inOperation,
      targetStatus: detailRow.targetStatus,
    })),
  }));

  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>공영주차장 차량 5부제 상세 내역</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f8fb;
      --surface: #ffffff;
      --line: #d8dee8;
      --line-strong: #aeb9c9;
      --text: #18202c;
      --muted: #5b6678;
      --accent: #176b87;
      --accent-soft: #e7f3f7;
      --target: #b42318;
      --target-soft: #fff0ee;
      --unknown: #805300;
      --unknown-soft: #fff6d8;
      --ok: #146c43;
      --ok-soft: #e9f7ef;
    }

    * {
      box-sizing: border-box;
    }

    [hidden] {
      display: none !important;
    }

    body {
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: "Malgun Gothic", "Apple SD Gothic Neo", Arial, sans-serif;
      font-size: 14px;
      line-height: 1.5;
    }

    header {
      background: #102331;
      color: #ffffff;
      padding: 24px 32px;
      border-bottom: 1px solid #07131c;
    }

    main {
      padding: 24px 32px 40px;
    }

    h1 {
      margin: 10px 0 10px;
      font-size: 28px;
      font-weight: 700;
      letter-spacing: 0;
    }

    h2 {
      margin: 0 0 14px;
      font-size: 18px;
      font-weight: 700;
      letter-spacing: 0;
    }

    p {
      margin: 0;
    }

    .subhead {
      color: #d8e3ec;
      font-size: 15px;
    }

    .back-link {
      color: #ffffff;
      display: inline-flex;
      font-weight: 700;
      text-decoration: none;
    }

    .back-link:hover,
    .back-link:focus {
      text-decoration: underline;
    }

    .section {
      max-width: 1600px;
      margin: 0 auto 22px;
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
    }

    .section-body {
      padding: 18px 20px;
    }

    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(5, minmax(140px, 1fr));
      gap: 1px;
      background: var(--line);
      border-top: 1px solid var(--line);
    }

    .kpi {
      background: var(--surface);
      padding: 16px 18px;
    }

    .kpi-label {
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }

    .kpi-value {
      margin-top: 6px;
      font-size: 24px;
      font-weight: 700;
      letter-spacing: 0;
    }

    table {
      width: 100%;
      border-collapse: collapse;
    }

    th,
    td {
      border-bottom: 1px solid var(--line);
      padding: 10px 12px;
      text-align: left;
      vertical-align: middle;
      word-break: keep-all;
    }

    thead th {
      background: #eef2f6;
      color: #263244;
      font-size: 12px;
      font-weight: 700;
      position: sticky;
      top: 0;
      z-index: 1;
    }

    .condition-table th {
      width: 140px;
    }

    .detail-table {
      min-width: 720px;
    }

    .vehicle-detail-table {
      min-width: 1180px;
    }

    .condition-table td {
      font-weight: 700;
    }

    .vehicle-number-empty {
      color: var(--muted);
      font-weight: 700;
    }

    .chart-layout {
      display: grid;
      grid-template-columns: minmax(180px, 260px) 1fr;
      gap: 24px;
      align-items: center;
    }

    .donut {
      --chart-percent: 0%;
      width: min(240px, 100%);
      aspect-ratio: 1;
      border-radius: 50%;
      background: conic-gradient(var(--target) var(--chart-percent), #edf1f5 0);
      display: grid;
      place-items: center;
      position: relative;
      margin: 0 auto;
    }

    .donut::after {
      content: "";
      position: absolute;
      width: 62%;
      aspect-ratio: 1;
      border-radius: 50%;
      background: var(--surface);
      border: 1px solid var(--line);
    }

    .donut-center {
      position: relative;
      z-index: 1;
      text-align: center;
    }

    .donut-value {
      display: block;
      font-size: 28px;
      font-weight: 700;
      letter-spacing: 0;
      color: var(--target);
    }

    .donut-label {
      display: block;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      margin-top: 2px;
    }

    .metric-list {
      display: grid;
      grid-template-columns: repeat(3, minmax(130px, 1fr));
      gap: 1px;
      background: var(--line);
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
    }

    .metric-item {
      background: var(--surface);
      padding: 14px 16px;
    }

    .metric-item span {
      display: block;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }

    .metric-item strong {
      display: block;
      margin-top: 6px;
      font-size: 22px;
      letter-spacing: 0;
    }

    .numeric {
      text-align: right;
      font-variant-numeric: tabular-nums;
    }

    .emphasis {
      color: var(--target);
      font-weight: 700;
    }

    .muted {
      color: var(--muted);
    }

    .status {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 54px;
      min-height: 24px;
      padding: 2px 8px;
      border-radius: 6px;
      font-weight: 700;
      white-space: nowrap;
    }

    .status-yes {
      background: var(--ok-soft);
      color: var(--ok);
    }

    .status-target {
      background: var(--target-soft);
      color: var(--target);
    }

    .status-unknown {
      background: var(--unknown-soft);
      color: var(--unknown);
    }

    .status-normal {
      background: #edf1f5;
      color: #394455;
    }

    .empty {
      color: var(--muted);
      text-align: center;
      padding: 24px;
    }

    .table-wrap {
      overflow: auto;
    }

    .message {
      color: var(--muted);
      margin-bottom: 14px;
    }

    .text-link {
      color: var(--accent);
      font-weight: 700;
      text-decoration: none;
    }

    .text-link:hover,
    .text-link:focus {
      text-decoration: underline;
    }

    @media (max-width: 980px) {
      header,
      main {
        padding-left: 16px;
        padding-right: 16px;
      }

      .kpi-grid {
        grid-template-columns: 1fr;
      }

      .chart-layout,
      .metric-list {
        grid-template-columns: 1fr;
      }

      h1 {
        font-size: 23px;
      }
    }
  </style>
</head>
<body>
  <header>
    <a class="back-link" href="${escapeHtml(MAIN_OUTPUT_BASENAME)}">메인 요약 페이지로 돌아가기</a>
    <h1>공영주차장 차량 5부제 상세 내역</h1>
    <p id="detail-subhead" class="subhead">선택된 상세 조건을 표시합니다.</p>
  </header>

  <main>
    <section id="not-found" class="section" aria-label="상세 조건 없음" hidden>
      <div class="section-body">
        <h2>상세 조건 없음</h2>
        <p class="message">선택된 상세 조건을 찾을 수 없습니다.</p>
        <a class="text-link" href="${escapeHtml(MAIN_OUTPUT_BASENAME)}">메인 요약 페이지로 돌아가기</a>
      </div>
    </section>

    <div id="parking-detail-content" hidden>
      <section class="section" aria-label="평균 타겟 차량 비율">
        <div class="section-body">
          <h2 id="selected-parking-name"></h2>
          <div class="chart-layout">
            <div id="target-donut" class="donut" aria-label="평균 타겟 차량 비율">
              <div class="donut-center">
                <span id="donut-ratio" class="donut-value"></span>
                <span class="donut-label">평균 타겟 차량 비율</span>
              </div>
            </div>
            <div class="metric-list" aria-label="평균 지표">
              <div class="metric-item">
                <span>평균 전체 차량 대수</span>
                <strong id="metric-average-total"></strong>
              </div>
              <div class="metric-item">
                <span>평균 타겟 차량 대수</span>
                <strong id="metric-average-target"></strong>
              </div>
              <div class="metric-item">
                <span>분석 평일 수</span>
                <strong id="metric-analysis-days"></strong>
              </div>
            </div>
          </div>
        </div>
      </section>

      <section class="section" aria-label="일자별 현황">
        <div class="section-body">
          <h2>일자별 현황</h2>
          <p class="muted">생성 시각: ${escapeHtml(generatedAt)}. 전체 요약 상세 기록 수: ${formatNumber(metadata.detailRows)}건.</p>
        </div>
        <div class="table-wrap">
          <table class="detail-table">
            <thead>
              <tr>
                <th>일자</th>
                <th class="numeric">전체 차량 대수</th>
                <th class="numeric">타겟 차량 대수</th>
                <th class="numeric">타겟 차량 비율</th>
              </tr>
            </thead>
            <tbody id="daily-rows"></tbody>
          </table>
        </div>
      </section>
    </div>

    <div id="date-parking-detail-content" hidden>
      <section class="section" aria-label="선택 조건">
        <div class="section-body">
          <h2>선택 조건</h2>
          <table class="condition-table">
            <tbody>
              <tr>
                <th scope="row">일자</th>
                <td id="condition-date"></td>
                <th scope="row">요일</th>
                <td id="condition-weekday"></td>
              </tr>
              <tr>
                <th scope="row">주차장명</th>
                <td id="condition-parking"></td>
                <th scope="row">운영시간</th>
                <td id="condition-operation"></td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      <section class="section" aria-label="요약 지표">
        <div class="kpi-grid">
          <div class="kpi">
            <div class="kpi-label">전체 차량 대수</div>
            <div id="date-kpi-total" class="kpi-value"></div>
          </div>
          <div class="kpi">
            <div class="kpi-label">타겟 차량 대수</div>
            <div id="date-kpi-target" class="kpi-value emphasis"></div>
          </div>
          <div class="kpi">
            <div class="kpi-label">판단불가 차량 대수</div>
            <div id="date-kpi-unknown" class="kpi-value"></div>
          </div>
          <div class="kpi">
            <div class="kpi-label">타겟 비율</div>
            <div id="date-kpi-ratio" class="kpi-value"></div>
          </div>
          <div class="kpi">
            <div class="kpi-label">상세 기록 수</div>
            <div id="date-kpi-detail-count" class="kpi-value"></div>
          </div>
        </div>
      </section>

      <section class="section" aria-label="상세 이용기록">
        <div class="section-body">
          <h2>상세 이용기록</h2>
          <p id="date-detail-description" class="muted"></p>
        </div>
        <div class="table-wrap">
          <table class="detail-table vehicle-detail-table">
            <thead>
              <tr>
                <th>일자</th>
                <th>주차장명</th>
                <th>운영시간</th>
                <th>입차시간</th>
                <th>출차시간</th>
                <th>차량번호</th>
                <th>차량번호 끝자리</th>
                <th>운영시간 내 이용 여부</th>
                <th>타겟 여부</th>
              </tr>
            </thead>
            <tbody id="vehicle-detail-rows"></tbody>
          </table>
        </div>
      </section>
    </div>
  </main>

  <script>
    const PARKING_DETAIL_DATA = ${serializeJsonForScript(parkingDetailData)};
    const DATE_PARKING_DETAIL_DATA = ${serializeJsonForScript(dateParkingDetailData)};
    const numberFormatter = new Intl.NumberFormat("ko-KR");

    function formatNumber(value) {
      return numberFormatter.format(value);
    }

    function formatAverage(value) {
      return Number(value || 0).toFixed(1);
    }

    function formatPercent(value) {
      return Number(value || 0).toFixed(2) + "%";
    }

    function setText(id, text) {
      document.getElementById(id).textContent = text;
    }

    function setHidden(id, hidden) {
      document.getElementById(id).hidden = hidden;
    }

    function appendCell(row, text, className) {
      const cell = document.createElement("td");
      if (className) {
        cell.className = className;
      }
      cell.textContent = text;
      row.appendChild(cell);
      return cell;
    }

    function appendStatusCell(row, text, statusClass) {
      const cell = document.createElement("td");
      const badge = document.createElement("span");
      badge.className = "status " + statusClass;
      badge.textContent = text;
      cell.appendChild(badge);
      row.appendChild(cell);
      return cell;
    }

    function getTargetStatusClass(status) {
      if (status === "대상") {
        return "status-target";
      }
      if (status === "판단불가") {
        return "status-unknown";
      }
      return "status-normal";
    }

    function getOperationStatusClass(status) {
      return status === "예" ? "status-yes" : "status-normal";
    }

    function renderDailyRows(selected) {
      const tbody = document.getElementById("daily-rows");
      tbody.textContent = "";

      if (selected.dailyRows.length === 0) {
        const row = document.createElement("tr");
        const cell = document.createElement("td");
        cell.colSpan = 4;
        cell.className = "empty";
        cell.textContent = "분석 기간에 표시할 현황이 없습니다.";
        row.appendChild(cell);
        tbody.appendChild(row);
        return;
      }

      selected.dailyRows.forEach((daily) => {
        const row = document.createElement("tr");
        appendCell(row, daily.dateKey + " (" + daily.weekdayLabel + ")");
        appendCell(row, formatNumber(daily.total), "numeric");
        appendCell(row, formatNumber(daily.target), "numeric emphasis");
        appendCell(row, formatPercent(daily.ratio), "numeric");
        tbody.appendChild(row);
      });
    }

    function renderVehicleDetailRows(selected) {
      const tbody = document.getElementById("vehicle-detail-rows");
      tbody.textContent = "";

      if (selected.detailRows.length === 0) {
        const row = document.createElement("tr");
        const cell = document.createElement("td");
        cell.colSpan = 9;
        cell.className = "empty";
        cell.textContent = "선택 조건에 표시할 상세 이용기록이 없습니다.";
        row.appendChild(cell);
        tbody.appendChild(row);
        return;
      }

      selected.detailRows.forEach((detail) => {
        const row = document.createElement("tr");
        appendCell(row, detail.dateKey + " (" + detail.weekdayLabel + ")");
        appendCell(row, detail.parkingName);
        appendCell(row, detail.operation);
        appendCell(row, detail.entryTime);
        appendCell(row, detail.exitTime);
        appendCell(row, detail.vehicleNumber || "빈 차량번호", detail.vehicleNumber ? "" : "vehicle-number-empty");
        appendCell(row, detail.lastDigit);
        appendStatusCell(row, detail.inOperation, getOperationStatusClass(detail.inOperation));
        appendStatusCell(row, detail.targetStatus, getTargetStatusClass(detail.targetStatus));
        tbody.appendChild(row);
      });
    }

    function showNotFound() {
      setHidden("not-found", false);
      setHidden("parking-detail-content", true);
      setHidden("date-parking-detail-content", true);
    }

    function renderParkingSelected(selected) {
      setHidden("not-found", true);
      setHidden("parking-detail-content", false);
      setHidden("date-parking-detail-content", true);
      setText("detail-subhead", selected.parkingName + " / 분석 평일 " + formatNumber(selected.analysisDateCount) + "일");
      setText("selected-parking-name", selected.parkingName);
      setText("donut-ratio", formatPercent(selected.averageRatio));
      setText("metric-average-total", formatAverage(selected.averageTotal));
      setText("metric-average-target", formatAverage(selected.averageTarget));
      setText("metric-analysis-days", formatNumber(selected.analysisDateCount) + "일");

      const donut = document.getElementById("target-donut");
      const chartPercent = Math.max(0, Math.min(100, Number(selected.averageRatio || 0)));
      donut.style.setProperty("--chart-percent", chartPercent.toFixed(4) + "%");
      donut.setAttribute("aria-label", selected.parkingName + " 평균 타겟 차량 비율 " + formatPercent(selected.averageRatio));

      renderDailyRows(selected);
    }

    function renderDateParkingSelected(selected) {
      setHidden("not-found", true);
      setHidden("parking-detail-content", true);
      setHidden("date-parking-detail-content", false);

      setText("detail-subhead", selected.dateKey + " (" + selected.weekdayLabel + ") / " + selected.parkingName + " / 상세 기록 " + formatNumber(selected.detailRows.length) + "건");
      setText("condition-date", selected.dateKey);
      setText("condition-weekday", selected.weekdayLabel);
      setText("condition-parking", selected.parkingName);
      setText("condition-operation", selected.operation);
      setText("date-kpi-total", formatNumber(selected.total));
      setText("date-kpi-target", formatNumber(selected.target));
      setText("date-kpi-unknown", formatNumber(selected.unknown));
      setText("date-kpi-ratio", formatPercent(selected.ratio));
      setText("date-kpi-detail-count", formatNumber(selected.detailRows.length));
      setText("date-detail-description", "생성 시각: ${escapeHtml(generatedAt)}. 선택 조건 상세 기록 수: " + formatNumber(selected.detailRows.length) + "건.");

      renderVehicleDetailRows(selected);
    }

    const params = new URLSearchParams(window.location.search);
    const selectedDate = params.get("date");
    const selectedParking = params.get("parking");

    if (selectedDate && selectedParking) {
      const selectedDateParking = DATE_PARKING_DETAIL_DATA.find((row) => row.dateKey === selectedDate && row.parkingName === selectedParking);
      if (selectedDateParking) {
        renderDateParkingSelected(selectedDateParking);
      } else {
        showNotFound();
      }
    } else if (selectedParking) {
      const selectedParkingData = PARKING_DETAIL_DATA.find((row) => row.parkingName === selectedParking);
      if (selectedParkingData) {
        renderParkingSelected(selectedParkingData);
      } else {
        showNotFound();
      }
    } else {
      showNotFound();
    }
  </script>
</body>
</html>`;
}

function main() {
  const analysisDates = getAnalysisDates();
  const sourceRows = readSourceRows();
  const { groups, stats } = analyzeRows(sourceRows, analysisDates);
  validateAnalysis(groups, analysisDates);

  const summaryRows = buildSummaryRows(groups);
  const parkingRows = buildParkingRows(summaryRows, analysisDates);
  const metadata = buildMetadata(summaryRows, stats, analysisDates);
  const html = renderHtml(parkingRows, summaryRows, metadata, analysisDates);
  const detailHtml = renderDetailHtml(parkingRows, summaryRows, metadata);

  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.writeFileSync(MAIN_OUTPUT_FILE, html, "utf8");
  fs.writeFileSync(DETAIL_OUTPUT_FILE, detailHtml, "utf8");
  writeExcelOutputs(parkingRows, summaryRows);

  console.log(`mainOutput=${MAIN_OUTPUT_FILE}`);
  console.log(`detailOutput=${DETAIL_OUTPUT_FILE}`);
  console.log(`parkingTargetXlsx=${PARKING_TARGET_XLSX_FILE}`);
  console.log(`summaryDetailXlsx=${SUMMARY_DETAIL_XLSX_FILE}`);
  console.log(`sourceRows=${metadata.sourceRows}`);
  console.log(`analysisWeekdays=${analysisDates.join(",")}`);
  console.log(`summaryRows=${summaryRows.length}`);
  console.log(`parkingRows=${parkingRows.length}`);
  console.log(`detailRows=${metadata.detailRows}`);
  console.log(`totalVehicles=${metadata.totalVehicles}`);
  console.log(`targetVehicles=${metadata.targetVehicles}`);
  console.log(`unknownVehicles=${metadata.unknownVehicles}`);
  console.log(`targetRatio=${formatPercent(metadata.targetRatio)}`);
  console.log(`missingExitRows=${metadata.missingExitRows}`);
  console.log(`blankVehicleRows=${metadata.blankVehicleRows}`);
  console.log(`invalidExitOrderRows=${metadata.invalidExitOrderRows}`);
}

main();
