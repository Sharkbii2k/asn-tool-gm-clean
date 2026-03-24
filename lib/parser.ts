function parseItems(rawText: string, fallbackLineNo: string): ParsedItem[] {
  const text = rawText.replace(/\n/g, " ").replace(/ +/g, " ");

  const items: ParsedItem[] = [];

  // pattern thực tế của bạn:
  // PO + Item + Rev + Qty + UOM
  const regex =
    /(\d{6,}-\d+)\s+(\d{7,})\s+(\d{2})\s+(\d+)\s+(PC|PCS)/g;

  let match;

  let seq = 1;

  while ((match = regex.exec(text)) !== null) {
    const poNo = match[1];
    const itemNo = match[2];
    const rev = match[3];
    const qty = Number(match[4]);
    const uom = match[5];

    items.push({
      seq: seq++,
      poNo,
      itemNo,
      rev,
      quantity: qty,
      uom,
      netWeight: "",
      grossWeight: "",
      packingSpec: "",
      lotRef: "",
      lineNo: fallbackLineNo
    });
  }

  return items;
}
