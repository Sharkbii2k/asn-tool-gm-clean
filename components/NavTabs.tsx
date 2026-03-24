"use client";
export default function NavTabs({tab,setTab}:{tab:"scan"|"packing"|"result"; setTab:(v:"scan"|"packing"|"result")=>void;}){
  return <div className="grid grid-cols-3 gap-3 rounded-[32px] bg-[#e8ecf2] p-2">
    <button className={`tab-chip ${tab==="scan"?"active":"inactive"}`} onClick={()=>setTab("scan")}>Scan ASN</button>
    <button className={`tab-chip ${tab==="packing"?"active":"inactive"}`} onClick={()=>setTab("packing")}>Packing</button>
    <button className={`tab-chip ${tab==="result"?"active":"inactive"}`} onClick={()=>setTab("result")}>Result</button>
  </div>;
}
