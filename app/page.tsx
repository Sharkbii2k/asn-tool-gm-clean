"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import NavTabs from "@/components/NavTabs";
import { buildLinesByAsn, buildSummary } from "@/lib/calc";
import { exportExcel } from "@/lib/excel";
import { docsToHeaderRows, fileToText, parseTextToDoc } from "@/lib/parser";
import { loadPacking, savePacking } from "@/lib/storage";
import { HeaderRow, LinesRow, PackingRow, ParsedDoc } from "@/lib/types";

export default function HomePage() {
  const [tab, setTab] = useState<"scan"|"packing"|"result">("scan");
  const [files, setFiles] = useState<File[]>([]);
  const [isScanning, setIsScanning] = useState(false);
  const [progress, setProgress] = useState("");
  const [docs, setDocs] = useState<ParsedDoc[]>([]);
  const [headers, setHeaders] = useState<HeaderRow[]>([]);
  const [lines, setLines] = useState<LinesRow[]>([]);
  const [packing, setPacking] = useState<PackingRow[]>([]);
  const [search, setSearch] = useState("");

  useEffect(()=>{ setPacking(loadPacking()); },[]);
  useEffect(()=>{ savePacking(packing); },[packing]);
  useEffect(()=>{ if(!docs.length) return; setHeaders(docsToHeaderRows(docs)); setLines(buildLinesByAsn(docs, packing)); },[docs, packing]);

  const scanFiles = async () => {
    if(!files.length) return;
    setIsScanning(true);
    try{
      const parsed: ParsedDoc[] = [];
      const total = Math.min(files.length, 50);
      for(let i=0;i<total;i++){
        setProgress(`Processing ${i+1} / ${total}...`);
        const text = await fileToText(files[i]);
        parsed.push(parseTextToDoc(text, files[i].name));
      }
      setDocs(parsed);
      setHeaders(docsToHeaderRows(parsed));
      setLines(buildLinesByAsn(parsed, packing));
      setProgress("Scan completed.");
      setTab("result");
    } catch (e) {
      console.error(e);
      setProgress("Scan failed.");
    } finally {
      setIsScanning(false);
    }
  };

  const importPacking = async (file?: File) => {
    if(!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<any>(sheet);
    const mapped = rows.map((r)=>({
      item: String(r.item ?? r.Item ?? r["Item"] ?? r["Mã hàng"] ?? "").trim(),
      rev: String(r.rev ?? r.Rev ?? r["Rev"] ?? "").trim(),
      pack: r.pack ?? r.Packing ?? r["Packing"] ?? ""
    })).filter((r)=>r.item);
    setPacking(mapped);
  };

  const summary = useMemo(()=>buildSummary(lines), [lines]);

  return <main className="app-shell">
    <div className="mb-6 text-center"><h1 className="text-5xl font-black tracking-tight text-brand-navy">ASN TOOL GM</h1></div>
    <NavTabs tab={tab} setTab={setTab} />

    {tab==="scan" && <section className="mt-8 space-y-8">
      <div className="card p-8">
        <h2 className="mb-5 text-3xl font-bold text-brand-navy">Scan ASN</h2>
        <label className="mb-4 flex cursor-pointer items-center justify-between rounded-3xl border border-brand-line bg-white px-5 py-5">
          <span className="text-base">{files.length ? `${files.length} file(s) selected` : "Choose PDF / JPG / PNG files"}</span>
          <input type="file" multiple accept=".pdf,.png,.jpg,.jpeg" className="hidden" onChange={(e)=>setFiles(Array.from(e.target.files || []).slice(0,50))} />
          <span className="small-btn">Choose files</span>
        </label>
        <div className="grid gap-3 md:grid-cols-2">
          <button className="secondary-btn" onClick={scanFiles} disabled={!files.length || isScanning}>{isScanning ? "Scanning..." : "Scan ASN"}</button>
          <button className="primary-btn" onClick={()=>exportExcel(docs, headers, lines, summary)} disabled={!docs.length}>Export Excel</button>
        </div>
        {progress && <p className="mt-4 text-base text-slate-600">{progress}</p>}
      </div>

      <div className="card p-8">
        <h3 className="mb-4 text-2xl font-bold text-brand-navy">Header Preview</h3>
        <div className="table-wrap"><table className="ui-table">
          <thead><tr>{["ASN No","ETA","ETD","Sold To","Bill To","Ship To","Location","Line No"].map(h=><th key={h}>{h}</th>)}</tr></thead>
          <tbody>{headers.map((row,i)=><tr key={i}>{["ASN No","ETA","ETD","Sold To","Bill To","Ship To","Location","Line No"].map(k=><td key={k}>{(row as any)[k]}</td>)}</tr>)}</tbody>
        </table></div>
      </div>
    </section>}

    {tab==="packing" && <section className="mt-8">
      <div className="card p-8">
        <h2 className="mb-5 text-3xl font-bold text-brand-navy">Packing Master</h2>
        <input className="input mb-4" placeholder="Search item..." value={search} onChange={(e)=>setSearch(e.target.value)} />
        <div className="table-wrap"><table className="ui-table">
          <thead><tr><th>Item</th><th>Rev</th><th>Packing</th><th>Action</th></tr></thead>
          <tbody>{packing.filter((r)=>!search || `${r.item} ${r.rev} ${r.pack}`.toLowerCase().includes(search.toLowerCase())).map((row,idx)=>
            <tr key={idx}>
              <td><input className="w-full rounded-xl border border-brand-line px-3 py-2" value={String(row.item||"")} onChange={(e)=>{ const next=[...packing]; next[idx]={...next[idx], item:e.target.value}; setPacking(next); }} /></td>
              <td><input className="w-full rounded-xl border border-brand-line px-3 py-2" value={String(row.rev||"")} onChange={(e)=>{ const next=[...packing]; next[idx]={...next[idx], rev:e.target.value}; setPacking(next); }} /></td>
              <td><input className="w-full rounded-xl border border-brand-line px-3 py-2" value={String(row.pack||"")} onChange={(e)=>{ const next=[...packing]; next[idx]={...next[idx], pack:e.target.value}; setPacking(next); }} /></td>
              <td><button className="small-btn" onClick={()=>setPacking(packing.filter((_,i)=>i!==idx))}>Delete</button></td>
            </tr>)}
          </tbody></table></div>
        <div className="mt-4 flex flex-wrap gap-3">
          <button className="small-btn" onClick={()=>setPacking([...packing,{item:"",rev:"",pack:""}])}>Add Row</button>
          <label className="small-btn cursor-pointer">Import Excel<input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e)=>importPacking(e.target.files?.[0])} /></label>
        </div>
      </div>
    </section>}

    {tab==="result" && <section className="mt-8 space-y-8">
      <div className="card p-8">
        <h2 className="mb-5 text-3xl font-bold text-brand-navy">Lines Result</h2>
        <div className="table-wrap"><table className="ui-table">
          <thead><tr>{["ASN","Item","Rev","Quantity","Packing","Thùng chẵn","SL lẻ PCS","Tổng Cartons","Line No","Location","Packing Found","Calc Status"].map(h=><th key={h}>{h}</th>)}</tr></thead>
          <tbody>{lines.map((row,i)=><tr key={i}>{["ASN","Item","Rev","Quantity","Packing","Thùng chẵn","SL lẻ PCS","Tổng Cartons","Line No","Location","Packing Found","Calc Status"].map(k=><td key={k}>{(row as any)[k]}</td>)}</tr>)}</tbody>
        </table></div>
      </div>
      <div className="card p-8">
        <h2 className="mb-5 text-3xl font-bold text-brand-navy">Total Cartons</h2>
        <div className="table-wrap"><table className="ui-table">
          <thead><tr><th>Location</th><th>Thùng chẵn</th><th>Tổng số thùng lẻ</th><th>Tổng</th></tr></thead>
          <tbody>{summary.map((row,i)=><tr key={i}><td>{row.Location}</td><td>{row["Thùng chẵn"]}</td><td>{row["Tổng số thùng lẻ"]}</td><td>{row["Tổng"]}</td></tr>)}</tbody>
        </table></div>
      </div>
      <button className="primary-btn" onClick={()=>exportExcel(docs, headers, lines, summary)} disabled={!docs.length}>Export Excel</button>
    </section>}
  </main>;
}
