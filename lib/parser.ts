"use client";
import * as pdfjsLib from "pdfjs-dist";
import Tesseract from "tesseract.js";
import { HeaderRow, ParsedDoc, ParsedItem } from "./types";
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;

const clean = (t:string)=>String(t||"").replace(/\u00a0/g," ").replace(/[|]/g," ").replace(/C2-\s+([0-9A-Z]+)/gi,"C2-$1").replace(/C1-\s+([0-9A-Z]+)/gi,"C1-$1").replace(/GP-\s+([0-9A-Z]+)/gi,"GP-$1").replace(/\s+/g," ").trim();
const pick = (re:RegExp, t:string)=>t.match(re)?.[1]?.trim() || "";
const rev2 = (v:string)=>{ const x=String(v||"").trim().replace(/\.0$/,""); return /^\d+$/.test(x)?x.padStart(2,"0"):x; };

export async function fileToText(file: File): Promise<string> {
  const lower=file.name.toLowerCase();
  if(lower.endsWith(".pdf")){
    const bytes=await file.arrayBuffer(); const pdf=await pdfjsLib.getDocument({data:bytes}).promise; let text="";
    for(let i=1;i<=pdf.numPages;i++){ const p=await pdf.getPage(i); const c=await p.getTextContent(); text += c.items.map((it:any)=>("str" in it?it.str:"")).join(" ") + "\n"; }
    if(text.trim()) return text;
  }
  const result = await Tesseract.recognize(file, "eng");
  return result.data.text || "";
}

function parseRows(text:string, defaultLine:string): ParsedItem[] {
  const flat = clean(text).replace(/Issued By.*$/i," ").replace(/Security Confirmed.*$/i," ").replace(/Received By.*$/i," ").replace(/Total Quantity.*$/i," ");
  const re = /(\d+)\s+([0-9A-Z-]{6,})\s+([0-9]{6,})\s+(\d{2})\s+([0-9]+(?:\.\d+)?)\s+(PC|SET|EA|PR)\s*([0-9]+(?:\.\d+)?)?\s*(?:([0-9]+(?:\.\d+)?)\s*)?(?:(\d+\*[0-9]+\+[0-9]+)\s*)?(?:So:\s*([0-9]{4,})\s*)?(?:XC([0-9]{5,6})\s*)?((?:C\d|GP)-[0-9A-Z]+)/gi;
  const out:ParsedItem[]=[]; const seen=new Set<string>(); let m;
  while((m=re.exec(flat))!==null){
    const key=`${m[1]}|${m[2]}|${m[3]}|${m[4]}|${m[5]}|${m[12]||defaultLine}`;
    if(seen.has(key)) continue; seen.add(key);
    out.push({ seq:Number(m[1]), poNo:m[2], itemNo:m[3], rev:rev2(m[4]), quantity:Number(m[5]), uom:m[6], netWeight:m[7]?Number(m[7]):"", grossWeight:m[8]?Number(m[8]):"", packingSpec:m[9]||"", lotRef:[m[10]?`So: ${m[10]}`:"", m[11]?`XC${m[11]}`:""].filter(Boolean).join("\n"), lineNo:m[12]||defaultLine||"" });
  }
  return out;
}

export function parseTextToDoc(text:string, sourceFile:string): ParsedDoc {
  const flat=clean(text);
  const asnNo = pick(/ASN\s*No:\s*([A-Z]{2}\d{6,})/i, flat) || pick(/\b([A-Z]{2}\d{6,})\b/, flat);
  const eta = pick(/ETA:\s*((?:20\d{2}-\d{2}-\d{2})\s+\d{2}:\d{2})/i, flat);
  const etd = pick(/ETD:\s*((?:20\d{2}-\d{2}-\d{2})\s+\d{2}:\d{2})/i, flat);
  const routeCode = pick(/\b(XC\d+-TC\d+)\b/i, flat);
  const lineNo = pick(/((?:C\d|GP)-[0-9A-Z]+)/i, flat);
  const soldTo = pick(/Sold To:\s*(.*?)\s*ASN No:/i, flat);
  const billTo = pick(/Bill To:\s*(.*?)\s*Ship To:/i, flat);
  const shipTo = pick(/Ship To:\s*(.*?)\s*ETA:/i, flat);
  const location = pick(/Location:\s*(.*?)\s*ETD:/i, flat);
  const items = parseRows(text, lineNo);
  return { sourceFile, asnNo, eta, etd, date:eta.split(" ")[0]||"", time:eta.split(" ")[1]||"", soldTo, billTo, shipTo, location, routeCode, lineNo, totalQuantity:items.reduce((s,x)=>s+Number(x.quantity||0),0), items, rawText:flat };
}

export function docsToHeaderRows(docs:ParsedDoc[]): HeaderRow[] {
  return docs.map(doc=>({ "ASN No":doc.asnNo, "ETA":doc.eta, "ETD":doc.etd, "Sold To":doc.soldTo, "Bill To":doc.billTo, "Ship To":doc.shipTo, "Location":doc.location, "Line No":doc.lineNo }));
}
