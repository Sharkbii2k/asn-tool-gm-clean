import { LinesRow, PackingRow, ParsedDoc } from "./types";
const rev2=(v:string)=>{ const x=String(v||"").trim().replace(/\.0$/,""); return /^\d+$/.test(x)?x.padStart(2,"0"):x; };
const loc=(line:string)=>String(line).includes("C2")?"CPT":String(line).includes("C1")?"OP":String(line).includes("GP")?"GP":"OTHER";

export function buildLinesByAsn(docs: ParsedDoc[], packing: PackingRow[]): LinesRow[] {
  const map=new Map<string,number>();
  packing.forEach(r=>{ const p=Number(r.pack); if(r.item && r.rev && p>0) map.set(`${String(r.item).trim()}__${rev2(String(r.rev))}`, p); });
  const out:LinesRow[]=[];
  for(const doc of docs){
    for(const item of doc.items){
      const pack = map.get(`${String(item.itemNo).trim()}__${rev2(item.rev)}`) || 0;
      if(pack>0){
        const full=Math.floor(Number(item.quantity)/pack), loosePcs=Number(item.quantity)%pack, looseCarton=loosePcs>0?1:0;
        out.push({ "ASN":doc.asnNo,"Item":item.itemNo,"Rev":rev2(item.rev),"Quantity":Number(item.quantity),"Packing":pack,"Thùng chẵn":full,"SL lẻ PCS":loosePcs,"Tổng Cartons":full+looseCarton,"Line No":item.lineNo||doc.lineNo,"Location":loc(item.lineNo||doc.lineNo),"Packing Found":"YES","Calc Status":"OK","__loose_carton":looseCarton });
      } else {
        out.push({ "ASN":doc.asnNo,"Item":item.itemNo,"Rev":rev2(item.rev),"Quantity":Number(item.quantity),"Packing":"","Thùng chẵn":"","SL lẻ PCS":"","Tổng Cartons":"","Line No":item.lineNo||doc.lineNo,"Location":loc(item.lineNo||doc.lineNo),"Packing Found":"NO","Calc Status":"CHECK","__loose_carton":0 });
      }
    }
  }
  return out;
}
export function buildSummary(lines: LinesRow[]) {
  const rows=["CPT","OP","GP"].map(Location=>({Location,"Thùng chẵn":0,"Tổng số thùng lẻ":0,"Tổng":0}));
  lines.forEach(r=>{ const row=rows.find(x=>x.Location===r["Location"]); if(!row||r["Calc Status"]!=="OK") return; row["Thùng chẵn"]+=Number(r["Thùng chẵn"]||0); row["Tổng số thùng lẻ"]+=Number(r["__loose_carton"]||0); row["Tổng"]+=Number(r["Tổng Cartons"]||0); });
  return rows;
}
