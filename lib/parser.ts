"use client";

import * as pdfjsLib from "pdfjs-dist";
import Tesseract from "tesseract.js";
import { HeaderRow, ParsedDoc, ParsedItem } from "./types";

pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;

function normalizeText(text: string): string {
  return String(text || "")
    .replace(/\u00a0/g, " ")
    .replace(/[|]/g, " ")
    .replace(/[：]/g, ":")
    .replace(/\r/g, "\n")
    // nối đúng line no bị OCR tách dòng: C2- \n 013D
    .replace(/((?:C\d|GP)-)\s*\n\s*([0-9OIl]{3}[A-Z])/gi, "$1$2")
    .replace(/((?:C\d|GP)-)\s+([0-9OIl]{3}[A-Z])/gi, "$1$2")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{2,}/g, "\n")
    .trim();
}

function flattenText(text: string): string {
  return normalizeText(text).replace(/\n/g, " ").replace(/\s+/g, " ").trim();
}

function normRev(v: string): string {
  const txt = String(v || "").trim().replace(/\.0$/, "");
  return /^\d+$/.test(txt) ? txt.padStart(2, "0") : txt;
}

function first(pattern: RegExp, text: string): string {
  return text.match(pattern)?.[1]?.trim() || "";
}

function normalizeLineSuffix(v: string): string {
  const s = String(v || "").toUpperCase().replace(/[^0-9OIL A-Z]/g, "").replace(/\s+/g, "");
  if (s.length < 4) return s;
  const head = s.slice(0, 3).replace(/O/g, "0").replace(/[IL]/g, "1");
  const tail = s.slice(3, 4).replace(/0/g, "D");
  return `${head}${tail}`;
}

function normalizeLineNo(v: string): string {
  const m = String(v || "").toUpperCase().match(/((?:C\d|GP)-)([0-9OIL]{3}[A-Z])/);
  if (!m) return "";
  return `${m[1]}${normalizeLineSuffix(m[2])}`;
}

async function pdfToOcrText(file: File): Promise<string> {
  const bytes = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;

  let finalText = "";

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 2.2 });

    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    if (!context) continue;

    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    await page.render({
      canvasContext: context,
      viewport,
    }).promise;

    const blob: Blob = await new Promise((resolve) => {
      canvas.toBlob((b) => resolve(b as Blob), "image/png");
    });

    const result = await Tesseract.recognize(blob, "eng");
    finalText += "\n" + result.data.text;
  }

  return finalText;
}

export async function fileToText(file: File): Promise<string> {
  const lower = file.name.toLowerCase();

  if (lower.endsWith(".pdf")) {
    return await pdfToOcrText(file);
  }

  const result = await Tesseract.recognize(file, "eng");
  return result.data.text || "";
}

/**
 * Chỉ lấy Line No thật của Goodmark:
 * - C2-013D
 * - C2-001D
 * - C2-007D
 * - C2-014D
 * Không lấy route XC2-TC2
 * Không lấy lot XC062951
 */
function extractDocLineNo(rawText: string): string {
  const raw = normalizeText(rawText);
  const one = flattenText(rawText);

  const candidates: string[] = [];

  // 1) Bắt mẫu gần cột cuối kiểu: XC062951 C2-013D
  for (const m of one.matchAll(/\bXC\d{5,6}\s+((?:C\d|GP)-[0-9OIL]{3}[A-Z0-9])\b/gi)) {
    candidates.push(m[1]);
  }

  // 2) Bắt mẫu sau chữ Line No nếu OCR giữ được header
  for (const m of one.matchAll(/Line\s*No\.?\s*((?:C\d|GP)-[0-9OIL]{3}[A-Z0-9])/gi)) {
    candidates.push(m[1]);
  }

  // 3) Fallback: mọi mẫu line thật trong text
  for (const m of one.matchAll(/\b((?:C\d|GP)-[0-9OIL]{3}[A-Z0-9])\b/gi)) {
    candidates.push(m[1]);
  }

  const normalized = candidates
    .map((v) => {
      const s = String(v).toUpperCase();
      const mm = s.match(/^((?:C\d|GP)-)([0-9OIL]{3})([A-Z0-9])$/);
      if (!mm) return "";

      const prefix = mm[1];
      const nums = mm[2]
        .replace(/O/g, "0")
        .replace(/I/g, "1")
        .replace(/L/g, "1");

      let suffix = mm[3];
      if (suffix === "0" || suffix === "O") suffix = "D";

      return `${prefix}${nums}${suffix}`;
    })
    .filter(Boolean)
    .filter((v) => !/XC/i.test(v))
    .filter((v) => /^(?:C\d|GP)-\d{3}[A-Z]$/.test(v));

  if (!normalized.length) return "";

  const counts = new Map<string, number>();
  for (const v of normalized) {
    counts.set(v, (counts.get(v) || 0) + 1);
  }

  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0][0];
}

function parseItems(rawText: string, docLineNo: string): ParsedItem[] {
  const text = flattenText(rawText)
    .replace(/GOOD MARK INDUSTRIAL VIETNAM COMPANY LIMITED\(\d+\)/gi, " ")
    .replace(/Delivery Note/gi, " ")
    .replace(/Issued By.*$/i, " ")
    .replace(/Security Confirmed.*$/i, " ")
    .replace(/Received By.*$/i, " ")
    .replace(/Total Quantity.*$/i, " ");

  const items: ParsedItem[] = [];
  const seen = new Set<string>();

  // PO No | Item No | Rev | Quantity | Uom | NetWeight(optional)
  const regex =
    /(\d{6,}-\d+)\s+(\d{7,})\s+(\d{2})\s+(\d+)\s+(PC|PCS|EA|SET|PR)\s*([\d.]+)?/gi;

  let match: RegExpExecArray | null;
  let seq = 1;

  while ((match = regex.exec(text)) !== null) {
    const poNo = match[1];
    const itemNo = match[2];
    const rev = normRev(match[3]);
    const quantity = Number(match[4]);
    const uom = match[5];
    const netWeight = match[6] || "";

    const tail = text.slice(match.index, Math.min(text.length, match.index + 260));

    const lotSo = tail.match(/So:\s*([0-9]{4,})/i)?.[1] || "";
    const lotXc = tail.match(/\bXC([0-9]{5,6})\b/i)?.[1] || "";

    // Khóa cứng: mọi item trong ASN dùng line no của document
    const lineNo = docLineNo || "";

    const key = `${poNo}|${itemNo}|${rev}|${quantity}|${lineNo}`;
    if (seen.has(key)) continue;
    seen.add(key);

    items.push({
      seq: seq++,
      poNo,
      itemNo,
      rev,
      quantity,
      uom,
      netWeight,
      grossWeight: "",
      packingSpec: "",
      lotRef: [
        lotSo ? `So: ${lotSo}` : "",
        lotXc ? `XC${lotXc}` : "",
      ]
        .filter(Boolean)
        .join("\n"),
      lineNo,
    });
  }

  return items;
}

export function parseTextToDoc(text: string, sourceFile: string): ParsedDoc {
  const raw = normalizeText(text);
  const one = flattenText(text);

  const asnNo =
    first(/ASN\s*No\s*:\s*([A-Z]{2}\d{6,})/i, one) ||
    first(/\b([A-Z]{2}\d{6,})\b/, one);

  const eta = first(/ETA\s*:\s*((?:20\d{2}-\d{2}-\d{2})\s+\d{2}:\d{2})/i, one);
  const etd = first(/ETD\s*:\s*((?:20\d{2}-\d{2}-\d{2})\s+\d{2}:\d{2})/i, one);
  const routeCode = first(/\b(XC\d+-TC\d+)\b/i, one);

  const soldTo = first(/Sold To\s*:\s*(.*?)\s*ASN\s*No\s*:/i, one);
  const billTo = first(/Bill To\s*:\s*(.*?)\s*Ship To\s*:/i, one);
  const shipTo = first(/Ship To\s*:\s*(.*?)\s*ETA\s*:/i, one);
  const location = first(/Location\s*:\s*(.*?)\s*ETD\s*:/i, one);

 const lineNo = extractDocLineNo(raw) || extractDocLineNo(one);
  const items = parseItems(raw, lineNo);
  const totalQuantity = items.reduce((sum, x) => sum + Number(x.quantity || 0), 0);

  const [date = "", time = ""] = eta.split(" ");

  return {
    sourceFile,
    asnNo,
    eta,
    etd,
    date,
    time,
    soldTo,
    billTo,
    shipTo,
    location,
    routeCode,
    lineNo,
    totalQuantity,
    items,
    rawText: raw,
  };
}

export function docsToHeaderRows(docs: ParsedDoc[]): HeaderRow[] {
  return docs.map((doc) => ({
    "ASN No": doc.asnNo,
    "ETA": doc.eta,
    "ETD": doc.etd,
    "Sold To": doc.soldTo,
    "Bill To": doc.billTo,
    "Ship To": doc.shipTo,
    "Location": doc.location,
    "Line No": doc.lineNo,
  }));
}
