import { PackingRow } from "./types";
const KEY="asn-tool-packing-master";
export function loadPacking(): PackingRow[] { if(typeof window==="undefined") return []; try{return JSON.parse(localStorage.getItem(KEY)||"[]")}catch{return []} }
export function savePacking(rows: PackingRow[]) { localStorage.setItem(KEY, JSON.stringify(rows)); }
