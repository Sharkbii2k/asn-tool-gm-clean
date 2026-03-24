while ((m = rowRegex.exec(t)) !== null) {
    const seq = Number(m[1]);
    const poNo = m[2];
    const itemNo = m[3];
    const rev = normRev(m[4]);
    const quantity = Number(m[5]);
    const uom = m[6];
    const netWeight = m[7] ? Number(m[7]) : "";
    const grossWeight = m[8] ? Number(m[8]) : "";
    const packingSpec = m[9] || "";
    const soNo = m[10] ? `So: ${m[10]}` : "";
    const xcNo = m[11] ? `XC${m[11]}` : "";
    const lineNo = m[12] || fallbackLineNo || "";

    const key = `${seq}|${poNo}|${itemNo}|${rev}|${quantity}|${lineNo}`;
    if (seen.has(key)) continue;
    seen.add(key);

    items.push({
      seq,
      poNo,
      itemNo,
      rev,
      quantity,
      uom,
      netWeight,
      grossWeight,
      packingSpec,
      lotRef: [soNo, xcNo].filter(Boolean).join("\n"),
      lineNo
    });
  }

  // fallback for rows missing weight / packing columns
  if (!items.length) {
    const simpleRegex =
      /(\d+)\s+([0-9A-Z-]{6,})\s+([0-9]{6,})\s+(\d{2})\s+([0-9]+(?:\.\d+)?)\s+(PC|SET|EA|PR)\s*([0-9]+(?:\.\d+)?)?\s*(?:So:\s*([0-9]{4,}))?\s*(?:XC([0-9]{5,6}))?\s*((?:C\d|GP)-[0-9A-Z]+)/gi;

    while ((m = simpleRegex.exec(t)) !== null) {
      const seq = Number(m[1]);
      const poNo = m[2];
      const itemNo = m[3];
      const rev = normRev(m[4]);
      const quantity = Number(m[5]);
      const uom = m[6];
      const netWeight = m[7] ? Number(m[7]) : "";
      const soNo = m[8] ? `So: ${m[8]}` : "";
      const xcNo = m[9] ? `XC${m[9]}` : "";
      const lineNo = m[10] || fallbackLineNo || "";

      const key = `${seq}|${poNo}|${itemNo}|${rev}|${quantity}|${lineNo}`;
      if (seen.has(key)) continue;
      seen.add(key);

      items.push({
        seq,
        poNo,
        itemNo,
        rev,
        quantity,
        uom,
        netWeight,
        grossWeight: "",
        packingSpec: "",
        lotRef: [soNo, xcNo].filter(Boolean).join("\n"),
        lineNo
      });
    }
  }

  return items;
}

export function parseTextToDoc(text: string, sourceFile: string): ParsedDoc {
  const raw = squish(text);
  const one = flat(text);

  const asnNo =
    first(/ASN\s*No:\s*([A-Z]{2}\d{6,})/i, one) ||
    first(/\b([A-Z]{2}\d{6,})\b/, one);

  const eta = first(/ETA:\s*((?:20\d{2}-\d{2}-\d{2})\s+\d{2}:\d{2})/i, one);
  const etd = first(/ETD:\s*((?:20\d{2}-\d{2}-\d{2})\s+\d{2}:\d{2})/i, one);
  const routeCode = first(/\b(XC\d+-TC\d+)\b/i, one);

  // choose last real line code from table/header, not route
  const lineMatches = Array.from(one.matchAll(/\b((?:C\d|GP)-[0-9A-Z]+)\b/gi)).map((x) => x[1]);
  const lineNo = lineMatches.length ? lineMatches[lineMatches.length - 1] : "";

  const soldTo = first(/Sold To:\s*(.*?)\s*ASN No:/i, one);
  const billTo = first(/Bill To:\s*(.*?)\s*Ship To:/i, one);
  const shipTo = first(/Ship To:\s*(.*?)\s*ETA:/i, one);
  const location = first(/Location:\s*(.*?)\s*ETD:/i, one);

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
