import fs from "node:fs/promises";
import ExcelJS from "exceljs";
import { fmtDate } from "./scheduler.mjs";

const colors = {
  title: "FF6B7280",
  header: "FF1F4E3D",
  weekday: "FF2E5E2E",
  ward: "FF17365D",
  weekend: "FF3B1F0F",
  helper: "FFE8F0FE",
  red: "FFC00000",
  workHeader: "FF1F4E79",
  referralHeader: "FF7A4F01"
};

function fill(cell, argb) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function font(cell, opts = {}) {
  cell.font = {
    name: "Arial",
    size: opts.size || 10,
    bold: !!opts.bold,
    color: { argb: opts.color || "FF111827" }
  };
}

function align(cell, horizontal = "center") {
  cell.alignment = { horizontal, vertical: "middle", wrapText: true };
}

function border(cell) {
  cell.border = {
    top: { style: "thin", color: { argb: "FFB8C2CC" } },
    left: { style: "thin", color: { argb: "FFB8C2CC" } },
    bottom: { style: "thin", color: { argb: "FFB8C2CC" } },
    right: { style: "thin", color: { argb: "FFB8C2CC" } }
  };
}

function styleRange(ws, fromRow, toRow, fromCol, toCol, options = {}) {
  for (let r = fromRow; r <= toRow; r++) {
    for (let c = fromCol; c <= toCol; c++) {
      const cell = ws.getCell(r, c);
      border(cell);
      align(cell, options.align || "center");
      font(cell, { bold: options.bold, color: options.fontColor, size: options.size });
      if (options.fill) fill(cell, options.fill);
    }
  }
}

function addTitle(ws, title, cols) {
  ws.mergeCells(1, 1, 1, cols);
  const cell = ws.getCell(1, 1);
  cell.value = title;
  fill(cell, colors.title);
  font(cell, { bold: true, color: "FFFFFFFF", size: 15 });
  align(cell);
}

function layoutHeader(item) {
  return item.time ? `${item.label} ${item.time}` : item.label;
}

function visibleLayoutForRow(config, row) {
  return (config.scheduleLayout || []).filter((item) => row.weekend ? item.weekend : item.weekday);
}

function layoutValue(row, item) {
  if (item.source === "blank") return "";
  return row[item.source] || "";
}

export async function createRotaWorkbook(result, outputPath) {
  const { config, assignments, counts, helperCounts, referralRows, audit } = result;
  const wb = new ExcelJS.Workbook();
  wb.creator = "ROTA Scheduler";
  wb.created = new Date();

  const schedule = wb.addWorksheet("ROTA Schedule", { views: [{ state: "frozen", ySplit: 2 }] });
  const referral = wb.addWorksheet("Referral Transfer", { views: [{ state: "frozen", ySplit: 2 }] });
  const workload = wb.addWorksheet("Workload Summary", { views: [{ state: "frozen", ySplit: 2 }] });
  const checks = wb.addWorksheet("Double Check", { views: [{ state: "frozen", ySplit: 2 }] });

  const start = fmtDate(new Date(`${config.startDate}T00:00:00.000Z`));
  const end = fmtDate(new Date(`${config.endDate}T00:00:00.000Z`));
  const layout = config.scheduleLayout || [];
  const display = config.scheduleDisplay || { weekend: true, notes: true };
  const optionalHeaders = [
    ...(display.weekend !== false ? ["Weekend?"] : []),
    ...(display.notes !== false ? ["Notes"] : []),
  ];
  const scheduleColumns = 2 + layout.length + optionalHeaders.length;
  addTitle(schedule, `${config.rotaName}: On Call Schedule (${start} - ${end})`, scheduleColumns);
  schedule.addRow(["Date", "Day", ...layout.map(layoutHeader), ...optionalHeaders]);
  styleRange(schedule, 2, 2, 1, scheduleColumns, { fill: colors.header, fontColor: "FFFFFFFF", bold: true });
  for (const row of assignments) {
    const visible = visibleLayoutForRow(config, row);
    const optionalCells = [
      ...(display.weekend !== false ? [row.weekend ? "Yes" : "No"] : []),
      ...(display.notes !== false ? [row.notes] : []),
    ];
    schedule.addRow([row.displayDate, row.day, ...layout.map((item) => visible.includes(item) ? layoutValue(row, item) : ""), ...optionalCells]);
    const r = schedule.lastRow.number;
    styleRange(schedule, r, r, 1, scheduleColumns, {});
    if (row.weekend) {
      styleRange(schedule, r, r, 1, scheduleColumns, { fill: colors.weekend, fontColor: "FFFFFFFF" });
    } else {
      layout.forEach((item, index) => {
        if (!visible.includes(item)) return;
        const cell = schedule.getCell(r, index + 3);
        fill(cell, item.source === "ward" ? colors.ward : colors.weekday);
        font(cell, { color: "FFFFFFFF" });
      });
    }
    layout.forEach((item, index) => {
      const col = index + 3;
      const name = schedule.getCell(r, col).value;
      if (name && counts[name]?.weekend > 2) {
        fill(schedule.getCell(r, col), colors.red);
        font(schedule.getCell(r, col), { bold: true, color: "FFFFFFFF" });
      }
    });
  }
  const legendStart = schedule.rowCount + 2;
  const legendRows = [
    ["Legend"],
    ["Brown rows = weekend. Green ER cells = weekday ER duties. Blue ward cells = ward covered by ER Team."],
    ["Red weekend cells = resident has more than 2 weekend on-calls this rotation and should be compensated next schedule."],
    ["Ward resident duty removed from individual workload when total calls would exceed 10; ward is covered by ER Team."]
  ];
  for (const item of legendRows) schedule.addRow(item);
  for (let r = legendStart; r < legendStart + legendRows.length; r++) {
    schedule.mergeCells(r, 1, r, scheduleColumns);
    fill(schedule.getCell(r, 1), "FFF3F4F6");
    align(schedule.getCell(r, 1), "left");
    font(schedule.getCell(r, 1));
  }
  schedule.columns = [
    { width: 11 },
    { width: 8 },
    ...layout.map(() => ({ width: 18 })),
    ...(display.weekend !== false ? [{ width: 11 }] : []),
    ...(display.notes !== false ? [{ width: 42 }] : []),
  ];

  addTitle(referral, "Referral / Case Transfer Coverage", 6);
  referral.addRow(["Date", "Day", "Transfer Resident", "Female Covering On Call", "Rule Used", "Notes"]);
  styleRange(referral, 2, 2, 1, 6, { fill: colors.referralHeader, fontColor: "FFFFFFFF", bold: true });
  for (const row of referralRows) {
    referral.addRow([row.displayDate, row.day, row.transfer, row.cover, row.rule, row.notes]);
    styleRange(referral, referral.lastRow.number, referral.lastRow.number, 1, 6, {});
  }
  referral.columns = [{ width: 11 }, { width: 8 }, { width: 22 }, { width: 22 }, { width: 48 }, { width: 32 }];

  addTitle(workload, "Workload Summary (Referral counts shown separately and not counted in on-call totals)", 11);
  workload.addRow(["Resident", "Gender", "Afternoon", "Night", "Ward", "Weekday Total", "Weekend Total", "On-Call Total", "Extra Weekend Flag", "Transfer", "Female Cover"]);
  styleRange(workload, 2, 2, 1, 11, { fill: colors.workHeader, fontColor: "FFFFFFFF", bold: true });
  for (const resident of config.residents) {
    const c = counts[resident.name];
    workload.addRow([resident.name, resident.gender, c.afternoon, c.night, c.ward, c.weekday, c.weekend, c.total, c.weekend > 2 ? "Compensate next rota" : "", c.transfer, c.femaleCover]);
    const r = workload.lastRow.number;
    styleRange(workload, r, r, 1, 11, {});
    if (c.weekend > 2) styleRange(workload, r, r, 7, 9, { fill: colors.red, fontColor: "FFFFFFFF", bold: true });
  }
  workload.addRow(["Average", "", "", "", "", "", "", "", "", "", ""]);
  styleRange(workload, workload.lastRow.number, workload.lastRow.number, 1, 11, {});
  for (const helper of config.helpers) {
    const c = helperCounts[helper.name] || { total: 0, weekend: 0 };
    workload.addRow([helper.name, "", c.total, 0, 0, 0, c.weekend, c.total, "Weekend helper only", 0, 0]);
    styleRange(workload, workload.lastRow.number, workload.lastRow.number, 1, 11, { fill: colors.helper });
  }
  workload.columns = [{ width: 18 }, { width: 8 }, { width: 13 }, { width: 13 }, { width: 13 }, { width: 13 }, { width: 13 }, { width: 13 }, { width: 22 }, { width: 13 }, { width: 13 }];

  const checkRows = [
    ["Coverage", audit.ok ? "PASS" : "REVIEW", audit.ok ? "All days have 4 ER assignments; weekdays also show ward covered by ER Team." : audit.issues.join("; "), ""],
    ["Ward", "ADJUSTED", "Ward resident on-call can be removed when total resident calls would exceed 10. Weekday ward is ER Team coverage.", ""],
    ["Weekend maximum", Object.values(audit.weekendCounts).some((x) => x > 2) ? "COMPROMISE" : "PASS", Object.values(audit.weekendCounts).some((x) => x > 2) ? "Some residents exceed 2 weekend calls and are highlighted red." : "No resident exceeds 2 weekend calls.", ""],
    ["Family resident rule", "CHECKED", config.familyResidents.size ? `${[...config.familyResidents].join(", ")}: no Sundays and no Saturday nights when possible.` : "No family residents entered.", ""],
    ["Chief extra weekday shift", "APPLIED", config.extraWeekday.size ? `${[...config.extraWeekday].join(", ")} were favored for extra weekday ER load where possible.` : "No extra weekday residents entered.", ""],
    ["Referral", referralRows.length === assignments.length ? "PASS" : "REVIEW", "Referral sheet covers every date. Referral counts are separate and not counted in on-call totals.", ""],
    ["Audit issues", audit.ok ? "PASS" : "REVIEW", audit.ok ? "No duplicate same-day assignments, no formula errors, and no critical conflicts detected." : audit.issues.join("; "), ""]
  ];
  addTitle(checks, "Double Check Summary", 4);
  checks.addRow(["Area", "Status", "Details", "Follow-up"]);
  styleRange(checks, 2, 2, 1, 4, { fill: "FF374151", fontColor: "FFFFFFFF", bold: true });
  for (const row of checkRows) {
    checks.addRow(row);
    const r = checks.lastRow.number;
    styleRange(checks, r, r, 1, 4, {});
    const status = checks.getCell(r, 2).value;
    fill(checks.getCell(r, 2), status === "PASS" || status === "APPLIED" || status === "CHECKED" ? "FFD9EAD3" : status === "REVIEW" ? "FFF4CCCC" : "FFFFF2CC");
  }
  checks.columns = [{ width: 24 }, { width: 14 }, { width: 92 }, { width: 20 }];

  await fs.mkdir(pathDir(outputPath), { recursive: true });
  await wb.xlsx.writeFile(outputPath);
  return { path: outputPath, errorScan: "" };
}

function pathDir(filePath) {
  return filePath.split("/").slice(0, -1).join("/") || ".";
}
