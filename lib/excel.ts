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
    horizontal: opts.center ? "center" : undefined,
    wrapText: true,
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

function sortByAsn<T extends { asnNo?: string; ASN?: string }>(rows: T[]): T[] {
  return [...rows].sort((a, b) => {
    const av = String(a.asnNo ?? a.ASN ?? "");
    const bv = String(b.asnNo ?? b.ASN ?? "");
    return av.localeCompare(bv);
  });
}

function getAsnFillMap(lines: LinesRow[]): Map<string, string> {
  const map = new Map<string, string>();
  let toggle = 0;

  for (const row of sortByAsn(lines)) {
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
  const cpt = docs.filter((d) => String(d.lineNo || "").includes("C2")).length;
  const op = docs.filter((d) => String(d.lineNo || "").includes("C1")).length;
  const gp = docs.filter((d) => String(d.lineNo || "").includes("GP")).length;

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
  ws.columns = [
    { width: 8 },  // A Seq
    { width: 16 }, // B PO
    { width: 15 }, // C Item
    { width: 8 },  // D Rev
    { width: 12 }, // E Qty
    { width: 8 },  // F Uom
    { width: 16 }, // G Net
    { width: 18 }, // H Gross
    { width: 16 }, // I Packing Spec
    { width: 28 }, // J Lot Ref
    { width: 12 }, // K Line No
    { width: 4 },  // L spacer
    { width: 14 }, // M summary label
    { width: 10 }, // N summary value
  ];

  const blockDocs = sortByAsn(docs);
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

  for (let i = 0; i < blockDocs.length; i++) {
    const doc = blockDocs[i];
    const blockFill = i % 2 === 0 ? COLORS.alt1 : COLORS.alt2;

    ws.mergeCells(row, 1, row, 11);
    for (let c = 1; c <= 11; c++) {
      styleCell(ws.getCell(row, c), {
        fill: COLORS.title,
        border: true,
      });
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
    asnHeaders.forEach((label, idx) => {
      ws.getCell(headerRow, idx + 1).value = label;
      styleCell(ws.getCell(headerRow, idx + 1), {
        fill: COLORS.header,
        bold: true,
        center: true,
        border: true,
      });
    });

    const items = doc.items || [];
    items.forEach((item, itemIndex) => {
      const r = headerRow + 1 + itemIndex;
      const values = [
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

      values.forEach((v, c) => {
        ws.getCell(r, c + 1).value = safeText(v);
        styleCell(ws.getCell(r, c + 1), { border: true, center: true });
        ws.getCell(r, c + 1).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: blockFill },
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
  }

  const summaryRows = buildInlineSummary(blockDocs);
  summaryRows.forEach((entry, idx) => {
    const excelRow = idx + 1;
    ws.getCell(excelRow, 13).value = entry[0];
    ws.getCell(excelRow, 14).value = entry[1];

    styleCell(ws.getCell(excelRow, 13), {
      fill: entry[2],
      bold: true,
      center: true,
      border: true,
    });
    styleCell(ws.getCell(excelRow, 14), {
      fill: entry[2],
      bold: true,
      center: true,
      border: true,
    });
  });
}

function writeHeaderSheet(ws: ExcelJS.Worksheet, headers: HeaderRow[]) {
  ws.columns = [
    { width: 16 }, // ASN
    { width: 18 }, // ETA
    { width: 12 }, // ETD
    { width: 38 }, // Sold To
    { width: 52 }, // Bill To
    { width: 38 }, // Ship To
    { width: 62 }, // Location
    { width: 12 }, // Line No
  ];

  const headerCols = [
    "ASN No",
    "ETA",
    "ETD",
    "Sold To",
    "Bill To",
    "Ship To",
    "Location",
    "Line No",
  ] as const;

  headerCols.forEach((label, idx) => {
    ws.getCell(1, idx + 1).value = label;
    styleCell(ws.getCell(1, idx + 1), {
      fill: COLORS.header,
      bold: true,
      center: true,
      border: true,
    });
  });

  const rows = sortByAsn(headers as any) as HeaderRow[];
  rows.forEach((row, idx) => {
    const fill = idx % 2 === 0 ? COLORS.alt1 : COLORS.alt2;
    headerCols.forEach((key, cidx) => {
      ws.getCell(idx + 2, cidx + 1).value = safeText(row[key]);
      styleCell(ws.getCell(idx + 2, cidx + 1), { border: true });
      ws.getCell(idx + 2, cidx + 1).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: fill },
      };
    });
  });
}

function writeLinesSheet(
  ws: ExcelJS.Worksheet,
  lines: LinesRow[],
  summaryRows: SummaryRow[]
) {
  ws.columns = [
    { width: 16 }, // ASN
    { width: 14 }, // Item
    { width: 8 },  // Rev
    { width: 12 }, // Quantity
    { width: 10 }, // Packing
    { width: 12 }, // Thung chan
    { width: 10 }, // SL le PCS
    { width: 12 }, // Tong cartons
    { width: 14 }, // Line No
    { width: 10 }, // Location
    { width: 12 }, // Packing Found
    { width: 12 }, // Calc Status
    { width: 12 }, // M
    { width: 12 }, // N
    { width: 12 }, // O
    { width: 10 }, // P
  ];

  const lineCols = [
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

  lineCols.forEach((label, idx) => {
    ws.getCell(1, idx + 1).value = label;
    styleCell(ws.getCell(1, idx + 1), {
      fill: COLORS.header,
      bold: true,
      center: true,
      border: true,
    });
  });

  const sortedLines = sortByAsn(lines);
  const fillMap = getAsnFillMap(sortedLines);

  sortedLines.forEach((row, idx) => {
    const asnFill = fillMap.get(String(row["ASN"] || "")) || COLORS.alt1;

    lineCols.forEach((key, cidx) => {
      ws.getCell(idx + 2, cidx + 1).value = safeText(row[key]);
      styleCell(ws.getCell(idx + 2, cidx + 1), {
        border: true,
        center: true,
      });
      ws.getCell(idx + 2, cidx + 1).fill = {
        type:
