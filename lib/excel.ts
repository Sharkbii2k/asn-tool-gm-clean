"use client";

import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { HeaderRow, LinesRow, ParsedDoc } from "./types";

type SummaryRow = {
  Location: string;
  "Thùng chẵn": number;
  "Tổng số thùng lẻ": number;
  "Tổng": number;
};

const COLORS = {
  title: "FFF2CC",
  header: "D9EAD3",
  sub: "EDEDED",
  calc: "EAF2F8",
  ok: "E2F0D9",
  warn: "FCE4D6",
  green: "C6E0B4",
  blue: "BDD7EE",
  orange: "FCE4D6",
  red: "F4CCCC",
  alt1: "F7FBFF",
  alt2: "FFF9F0",
  border: "B7B7B7",
};

function styleCell(
  cell: ExcelJS.Cell,
  opts: {
    fill?: string;
    bold?: boolean;
    center?: boolean;
    border?: boolean;
    size?: number;
    wrap?: boolean;
  } = {}
) {
  if (opts.fill) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: opts.fill },
    };
  }

  cell.alignment = {
    vertical: "middle",
    horizontal: opts.center ? "center" : "left",
    wrapText: opts.wrap ?? true,
  };

  if (opts.bold || opts.size) {
    cell.font = {
      bold: !!opts.bold,
      size: opts.size || 11,
    };
  }

  if (opts.border) {
    cell.border = {
      top: { style: "thin", color: { argb: COLORS.border } },
      left: { style: "thin", color: { argb: COLORS.border } },
      right: { style: "thin", color: { argb: COLORS.border } },
      bottom: { style: "thin", color: { argb: COLORS.border } },
    };
  }
}

function safeText(v: unknown): string | number {
  if (v === null || v === undefined) return "";
  return v as string | number;
}

function sortDocs(docs: ParsedDoc[]): ParsedDoc[] {
  return [...docs].sort((a, b) => String(a.asnNo || "").localeCompare(String(b.asnNo || "")));
}

function sortHeaders(rows: HeaderRow[]): HeaderRow[] {
  return [...rows].sort((a, b) =>
    String(a["ASN No"] || "").localeCompare(String(b["ASN No"] || ""))
  );
}

function sortLines(rows: LinesRow[]): LinesRow[] {
  return [...rows].sort((a, b) =>
    String(a["ASN"] || "").localeCompare(String(b["ASN"] || ""))
  );
}

function getAsnFillMap(lines: LinesRow[]): Map<string, string> {
  const map = new Map<string, string>();
  let toggle = 0;

  for (const row of sortLines(lines)) {
    const asn = String(row["ASN"] || "");
    if (!asn) continue;
    if (!map.has(asn)) {
      map.set(asn, toggle % 2 === 0 ? COLORS.alt1 : COLORS.alt2);
      toggle += 1;
    }
  }

  return map;
}

function buildInlineSummary(docs: ParsedDoc[]) {
  const totalAsn = docs.length;
  const cpt = docs.filter((d) => String(d.lineNo || "").startsWith("C2")).length;
  const op = docs.filter((d) => String(d.lineNo || "").startsWith("C1")).length;
  const gp = docs.filter((d) => String(d.lineNo || "").startsWith("GP")).length;

  return [
    ["Tổng ASN", totalAsn, COLORS.green],
    ["CPT", cpt, COLORS.blue],
    ["OP", op, COLORS.orange],
    ["GP", gp, COLORS.red],
  ] as const;
}

function normalizeSummary(summary?: SummaryRow[]): SummaryRow[] {
  const base: SummaryRow[] = [
    { Location: "CPT", "Thùng chẵn": 0, "Tổng số thùng lẻ": 0, "Tổng": 0 },
    { Location: "OP", "Thùng chẵn": 0, "Tổng số thùng lẻ": 0, "Tổng": 0 },
    { Location: "GP", "Thùng chẵn": 0, "Tổng số thùng lẻ": 0, "Tổng": 0 },
  ];

  if (!summary?.length) return base;

  for (const row of summary) {
    const target = base.find((x) => x.Location === row.Location);
    if (!target) continue;
    target["Thùng chẵn"] = Number(row["Thùng chẵn"] || 0);
    target["Tổng số thùng lẻ"] = Number(row["Tổng số thùng lẻ"] || 0);
    target["Tổng"] = Number(row["Tổng"] || 0);
  }

  return base;
}

function writeAsnSheet(ws: ExcelJS.Worksheet, docs: ParsedDoc[]) {
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.columns = [
    { width: 8 },
    { width: 16 },
    { width: 15 },
    { width: 8 },
    { width: 12 },
    { width: 8 },
    { width: 16 },
    { width: 18 },
    { width: 16 },
    { width: 28 },
    { width: 12 },
    { width: 4 },
    { width: 14 },
    { width: 10 },
  ];

  const rows = sortDocs(docs);
  const asnHeaders = [
    "Seq",
    "PO No.",
    "Item No.",
    "Rev.",
    "Quantity",
    "Uom",
    "Net Weight (KG)",
    "Gross Weight (KG)",
    "Packing Spec.",
    "Lot/MI No./SO No./Invoice No",
    "Line No.",
  ];

  let row = 1;

  rows.forEach((doc, idx) => {
    const fill = idx % 2 === 0 ? COLORS.alt1 : COLORS.alt2;

    ws.mergeCells(row, 1, row, 11);
    for (let c = 1; c <= 11; c++) {
      styleCell(ws.getCell(row, c), { fill: COLORS.title, border: true });
    }

    ws.getCell(row, 1).value = safeText(doc.asnNo);
    styleCell(ws.getCell(row, 1), {
      fill: COLORS.title,
      bold: true,
      center: true,
      border: true,
      size: 15,
    });

    ws.getCell(row + 1, 1).value = "Date:";
    ws.getCell(row + 1, 2).value = safeText(doc.date);
    ws.getCell(row + 1, 4).value = "Time:";
    ws.getCell(row + 1, 5).value = safeText(doc.time);

    ws.getCell(row + 2, 1).value = "Route:";
    ws.getCell(row + 2, 2).value = safeText(doc.routeCode);
    ws.getCell(row + 2, 4).value = "Line No:";
    ws.getCell(row + 2, 5).value = safeText(doc.lineNo);

    styleCell(ws.getCell(row + 1, 1), { bold: true });
    styleCell(ws.getCell(row + 1, 4), { bold: true });
    styleCell(ws.getCell(row + 2, 1), { bold: true });
    styleCell(ws.getCell(row + 2, 4), { bold: true });

    const headerRow = row + 4;
    asnHeaders.forEach((label, i) => {
      ws.getCell(headerRow, i + 1).value = label;
      styleCell(ws.getCell(headerRow, i + 1), {
        fill: COLORS.header,
        bold: true,
        center: true,
        border: true,
      });
    });

    const items = doc.items || [];
    items.forEach((item, i) => {
      const r = headerRow + 1 + i;
      const vals = [
        item.seq,
        item.poNo,
        item.itemNo,
        item.rev,
        item.quantity,
        item.uom,
        item.netWeight,
        item.grossWeight,
        item.packingSpec,
        item.lotRef,
        item.lineNo || doc.lineNo,
      ];

      vals.forEach((v, c) => {
        ws.getCell(r, c + 1).value = safeText(v);
        styleCell(ws.getCell(r, c + 1), {
          border: true,
          center: true,
        });
        ws.getCell(r, c + 1).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: fill },
        };
      });
    });

    const totalRow = headerRow + 1 + items.length;
    ws.mergeCells(totalRow, 1, totalRow, 4);
    ws.getCell(totalRow, 1).value = "Total Quantity";
    styleCell(ws.getCell(totalRow, 1), {
      fill: COLORS.sub,
      bold: true,
      border: true,
    });

    ws.getCell(totalRow, 5).value = Number(doc.totalQuantity || 0);
    styleCell(ws.getCell(totalRow, 5), {
      fill: COLORS.sub,
      bold: true,
      center: true,
      border: true,
    });

    for (let c = 1; c <= 11; c++) {
      styleCell(ws.getCell(totalRow, c), { border: true });
    }

    row = totalRow + 3;
  });

  const summaryRows = buildInlineSummary(rows);
  summaryRows.forEach((entry, idx) => {
    const r = idx + 1;
    ws.getCell(r, 13).value = entry[0];
    ws.getCell(r, 14).value = entry[1];
    styleCell(ws.getCell(r, 13), {
      fill: entry[2],
      bold: true,
      center: true,
      border: true,
    });
    styleCell(ws.getCell(r, 14), {
      fill: entry[2],
      bold: true,
      center: true,
      border: true,
    });
  });
}

function writeHeaderSheet(ws: ExcelJS.Worksheet, headers: HeaderRow[]) {
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.columns = [
    { width: 16 },
    { width: 18 },
    { width: 12 },
    { width: 30 },
    { width: 36 },
    { width: 30 },
    { width: 42 },
    { width: 14 },
  ];

  const cols = [
    "ASN No",
    "ETA",
    "ETD",
    "Sold To",
    "Bill To",
    "Ship To",
    "Location",
    "Line No",
  ] as const;

  cols.forEach((label, idx) => {
    ws.getCell(1, idx + 1).value = label;
    styleCell(ws.getCell(1, idx + 1), {
      fill: COLORS.header,
      bold: true,
      center: true,
      border: true,
    });
  });

  const rows = sortHeaders(headers);
  rows.forEach((row, idx) => {
    const fill = idx % 2 === 0 ? COLORS.alt1 : COLORS.alt2;
    cols.forEach((key, cidx) => {
      const cell = ws.getCell(idx + 2, cidx + 1);
      cell.value = safeText(row[key]);
      styleCell(cell, { border: true, wrap: true });
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: fill },
      };
    });
    ws.getRow(idx + 2).height = 42;
  });
}

function writeLinesSheet(
  ws: ExcelJS.Worksheet,
  lines: LinesRow[],
  summary: SummaryRow[]
) {
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.columns = [
    { width: 16 },
    { width: 14 },
    { width: 8 },
    { width: 12 },
    { width: 10 },
    { width: 12 },
    { width: 10 },
    { width: 12 },
    { width: 14 },
    { width: 10 },
    { width: 12 },
    { width: 12 },
    { width: 12 },
    { width: 12 },
    { width: 12 },
    { width: 10 },
  ];

  const cols = [
    "ASN",
    "Item",
    "Rev",
    "Quantity",
    "Packing",
    "Thùng chẵn",
    "SL lẻ PCS",
    "Tổng Cartons",
    "Line No",
    "Location",
    "Packing Found",
    "Calc Status",
  ] as const;

  cols.forEach((label, idx) => {
    ws.getCell(1, idx + 1).value = label;
    styleCell(ws.getCell(1, idx + 1), {
      fill: COLORS.header,
      bold: true,
      center: true,
      border: true,
    });
  });

  const rows = sortLines(lines);
  const fillMap = getAsnFillMap(rows);

  rows.forEach((row, idx) => {
    const fill = fillMap.get(String(row["ASN"] || "")) || COLORS.alt1;

    cols.forEach((key, cidx) => {
      const cell = ws.getCell(idx + 2, cidx + 1);
      cell.value = safeText(row[key]);
      styleCell(cell, { border: true, center: true });
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: fill },
      };
    });

    ws.getCell(idx + 2, 5).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: row["Packing Found"] === "YES" ? COLORS.ok : COLORS.warn },
    };
    ws.getCell(idx + 2, 6).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: COLORS.calc },
    };
    ws.getCell(idx + 2, 7).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: COLORS.calc },
    };
    ws.getCell(idx + 2, 8).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: COLORS.calc },
    };
    ws.getCell(idx + 2, 11).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: row["Packing Found"] === "YES" ? COLORS.ok : COLORS.warn },
    };
    ws.getCell(idx + 2, 12).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: row["Calc Status"] === "OK" ? COLORS.ok : COLORS.warn },
    };
  });

  ws.mergeCells(1, 13, 1, 16);
  ws.getCell(1, 13).value = "Tổng Cartons";
  styleCell(ws.getCell(1, 13), {
    fill: COLORS.title,
    bold: true,
    center: true,
    border: true,
  });

  ["Location", "Thùng chẵn", "Tổng số thùng lẻ", "Tổng"].forEach((label, idx) => {
    ws.getCell(2, 13 + idx).value = label;
    styleCell(ws.getCell(2, 13 + idx), {
      fill: COLORS.header,
      bold: true,
      center: true,
      border: true,
    });
  });

  const normalized = normalizeSummary(summary);
  normalized.forEach((row, idx) => {
    const r = idx + 3;
    ws.getCell(r, 13).value = row.Location;
    ws.getCell(r, 14).value = row["Thùng chẵn"];
    ws.getCell(r, 15).value = row["Tổng số thùng lẻ"];
    ws.getCell(r, 16).value = row["Tổng"];

    for (let c = 13; c <= 16; c++) {
      styleCell(ws.getCell(r, c), { border: true, center: true });
      ws.getCell(r, c).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: c === 13 ? COLORS.sub : COLORS.calc },
      };
    }
  });
}

export async function exportExcel(
  docs: ParsedDoc[],
  headers: HeaderRow[],
  lines: LinesRow[],
  summary: SummaryRow[]
) {
  const workbook = new ExcelJS.Workbook();

  const asnSheet = workbook.addWorksheet("ASN");
  const headerSheet = workbook.addWorksheet("Header");
  const linesSheet = workbook.addWorksheet("Lines");

  writeAsnSheet(asnSheet, docs || []);
  writeHeaderSheet(headerSheet, headers || []);
  writeLinesSheet(linesSheet, lines || [], summary || []);

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  saveAs(blob, "ASN_TOOL_GM_final.xlsx");
}
