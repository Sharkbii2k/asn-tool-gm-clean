"use client";

import * as pdfjsLib from "pdfjs-dist";
import Tesseract from "tesseract.js";
import { HeaderRow, ParsedDoc, ParsedItem } from "./types";

pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;

function normalizeText(text: string) {
  return String(text || "")
    .replace(/\u00a0/g, " ")
    .replace(/[|]/g, " ")
    .replace(/[：]/g, ":")
    .replace(/C2\s*-\s*([0-9A-Z]+)/gi, "C2-$1")
    .replace(/C1\s*-\s*([0-9A-Z]+)/gi, "C1-$1")
    .replace(/GP\s*-\s*([0-9A-Z]+)/gi, "GP-$1")
    .replace(/C2-\s+([0-9A-Z]+)/gi, "C2-$1")
    .replace(/C1-\s+([0-9A-Z]+)/gi, "C1-$1")
    .replace(/GP-\s+([0-9A-Z]+)/gi, "GP-$1")
    .replace(/\s+/g, " ")
    .trim();
}

function normRev(v: string) {
  const txt = String(v || "").trim().replace(/\.0$/, "");
  return /^\d+$/.test(txt) ? txt.padStart(2, "0") : txt;
}

function first(pattern: RegExp, text: string) {
  return text.match(pattern)?.[1]?.trim() || "";
}

async function pdfToOcrText(file: File): Promise<string> {
  const bytes = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;

  let finalText = "";

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 2.0 });

    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    if (!context) continue;

    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    await page.render({
      canvasContext: context,
      viewport
    }).promise;

    const blob: Blob = await new Promise((resolve) =>
      canvas.toBlob((b) => resolve(b as Blob), "image/png")
    );

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

function parseItems(rawText: string, fallbackLineNo: string): ParsedItem[] {
  const text = normalizeText(rawText)
    .replace(/GOOD MARK INDUSTRIAL VIETNAM COMPANY LIMITED\(\d+\)/gi, " ")
    .replace(/Delivery Note/gi, " ")
    .replace(/Issued By.*$/i, " ")
    .replace(/Security Confirmed.*$/i, " ")
    .replace(/Received By.*$/i, " ")
    .replace(/Total Quantity.*$/i, " ");

  const items: ParsedItem[] = [];
  const seen = new Set<string>();

  // OCR-friendly parser: PO + Item + Rev + Qty + UOM + optional weight + optional line
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

    // try to find nearest line no after the match
    const tail = text.slice(match.index, Math.min(text.length, match.index + 220));
    const lineNo = tail.match(/\b((?:C\d|GP)-[0-9A-Z]+)\b/i)?.[1] || fallbackLineNo || "";

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
      lotRef: "",
      lineNo
    });
  }

  return items;
}

export function parseTextToDoc(text: string, sourceFile: string): ParsedDoc {
  const raw = normalizeText(text);

  const asnNo =
    first(/ASN\s*No\s*:\s*([A-Z]{2}\d{6,})/i, raw) ||
    first(/\b([A-Z]{2}\d{6,})\b/, raw);

  const eta = first(/ETA\s*:\s*((?:20\d{2}-\d{2}-\d{2})\s+\d{2}:\d{2})/i, raw);
  const etd = first(/ETD\s*:\s*((?:20\d{2}-\d{2}-\d{2})\s+\d{2}:\d{2})/i, raw);
  const routeCode = first(/\b(XC\d+-TC\d+)\b/i, raw);

  const lineMatches = Array.from(raw.matchAll(/\b((?:C\d|GP)-[0-9A-Z]+)\b/gi)).map((m) => m[1]);
  const lineNo = lineMatches.length ? lineMatches[lineMatches.length - 1] : "";

  const soldTo = first(/Sold To\s*:\s*(.*?)\s*ASN No\s*:/i, raw);
  const billTo = first(/Bill To\s*:\s*(.*?)\s*Ship To\s*:/i, raw);
  const shipTo = first(/Ship To\s*:\s*(.*?)\s*ETA\s*:/i, raw);
  const location = first(/Location\s*:\s*(.*?)\s*ETD\s*:/i, raw);

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
    rawText: raw
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
    "Line No": doc.lineNo
  }));
}
