export function extractSignatureBlock(body: string, senderName: string): string {
  if (senderName) {
    const idx = body.indexOf(senderName);
    if (idx >= 0) return body.substring(idx).trim();
  }
  const parts = body.split(/\r?\n\s*\r?\n/);
  return (parts.length > 1 ? parts.pop()! : body).trim();
}

export function parseSignature(
  sig: string,
  senderName: string,
  senderEmail: string,
  organization: string
) {
  const lines = sig.split(/\r?\n/).map(l => l.trim()).filter(l => l);

  // Name
  let name = "";
  if (senderName && sig.toLowerCase().includes(senderName.toLowerCase())) {
    name = senderName;
  } else {
    for (const line of lines) {
      if (line.includes(senderEmail))               continue;
      if (/\+?\d[\d\-\u2013()\s]{5,}\d/.test(line)) continue;
      if (/https?:\/\//i.test(line))               continue;
      if (/www\./i.test(line))                     continue;
      if (/^[\+\d]/.test(line))                    continue;
      name = line;
      break;
    }
  }

  // Email
  const email = senderEmail;

  // Phone (longest match)
  const phoneRe = /(\+?\d[\d\-\u2013()\s]{5,}\d)/g;
  const matches: string[] = [];
  let   m: RegExpExecArray | null;
  while (m = phoneRe.exec(sig)) matches.push(m[1]);
  const phone = matches.length
    ? matches.reduce((a,b) => a.length >= b.length ? a : b)
    : "";

  return { name, email, phone, organization };
}

