const MAX_BYTES = 8 * 1024 * 1024;

const PATTERNS = [
  { id: "email", label: "[EMAIL]", re: /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi },
  {
    id: "phone",
    label: "[PHONE]",
    re: /(?<!\d)(?:\+?1[\s.-]?)?(?:\(?\d{3}\)?[\s.-]?)\d{3}[\s.-]?\d{4}(?!\d)/g,
  },
  { id: "ssn", label: "[SSN]", re: /\b\d{3}-\d{2}-\d{4}\b/g },
  { id: "ip", label: "[IP]", re: /\b(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)\b/g },
];

const CARD_RE = /\b(?:\d[ -]*?){13,19}\b/g;

const els = {
  drop: document.getElementById("drop"),
  file: document.getElementById("file"),
  pick: document.getElementById("pick"),
  sample: document.getElementById("sample"),
  paste: document.getElementById("paste"),
  status: document.getElementById("status"),
  error: document.getElementById("error"),
  counts: document.getElementById("counts"),
  preview: document.getElementById("preview"),
  download: document.getElementById("download"),
  copy: document.getElementById("copy"),
};

let state = {
  name: "paste.txt",
  kind: "text",
  original: "",
  wiped: "",
  counts: { email: 0, phone: 0, ssn: 0, card: 0, ip: 0 },
  docxZip: null,
};

function luhnOk(digits) {
  let sum = 0;
  let alt = false;
  for (let i = digits.length - 1; i >= 0; i--) {
    let n = Number(digits[i]);
    if (alt) {
      n *= 2;
      if (n > 9) n -= 9;
    }
    sum += n;
    alt = !alt;
  }
  return sum % 10 === 0;
}

function wipeText(input) {
  const counts = { email: 0, phone: 0, ssn: 0, card: 0, ip: 0 };
  let out = input;

  out = out.replace(CARD_RE, (match) => {
    const digits = match.replace(/\D/g, "");
    if (digits.length < 13 || digits.length > 19 || !luhnOk(digits)) return match;
    counts.card += 1;
    return "[CARD]";
  });

  for (const p of PATTERNS) {
    out = out.replace(p.re, () => {
      counts[p.id] += 1;
      return p.label;
    });
  }

  return { wiped: out, counts };
}

function setError(msg) {
  els.error.hidden = !msg;
  els.error.textContent = msg || "";
}

function render() {
  const total = Object.values(state.counts).reduce((a, b) => a + b, 0);
  els.status.textContent = `${state.name} · ${total} hit${total === 1 ? "" : "s"} · files stayed in this tab`;
  els.counts.hidden = false;
  for (const k of ["email", "phone", "ssn", "card", "ip"]) {
    document.getElementById(`c-${k}`).textContent = String(state.counts[k]);
  }
  els.preview.hidden = false;
  const escaped = state.wiped
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replace(/\[(EMAIL|PHONE|SSN|CARD|IP)\]/g, "<mark>[$1]</mark>");
  els.preview.innerHTML = escaped || "(empty)";
  els.download.disabled = !state.wiped;
  els.copy.disabled = !state.wiped;
}

async function readAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(r.error);
    r.readAsArrayBuffer(file);
  });
}

async function readAsText(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result || ""));
    r.onerror = () => reject(r.error);
    r.readAsText(file);
  });
}

async function extractPdfText(buf) {
  const pdfjs = window.pdfjsLib;
  if (!pdfjs) throw new Error("PDF parser failed to load.");
  const doc = await pdfjs.getDocument({ data: buf, isEvalSupported: false }).promise;
  try {
    const parts = [];
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i);
      const content = await page.getTextContent();
      parts.push(content.items.map((it) => it.str).join(" "));
    }
    return parts.join("\n\n");
  } finally {
    await doc.destroy();
  }
}

async function extractDocxText(buf) {
  if (!window.JSZip) throw new Error("JSZip failed to load");
  const zip = await window.JSZip.loadAsync(buf);
  const xml = await zip.file("word/document.xml")?.async("string");
  if (!xml) throw new Error("Not a valid .docx");
  if (xml.length > MAX_BYTES || /<!DOCTYPE/i.test(xml)) throw new Error("Unsupported DOCX XML.");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("Invalid DOCX XML.");
  const ns = doc.documentElement.namespaceURI;
  if (!["http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main"].includes(ns)) {
    throw new Error("Unsupported DOCX XML namespace.");
  }
  const body = doc.getElementsByTagNameNS(ns, "body")[0];
  if (!body) throw new Error("Missing DOCX body.");
  const text = [...body.getElementsByTagNameNS(ns, "p")].map((p) => {
    let line = "";
    const visit = (node) => {
      for (const child of node.children) {
        if (child.namespaceURI === ns) {
          if (child.localName === "p" || child.localName === "del") continue;
          if (child.localName === "t") { line += child.textContent; continue; }
          if (child.localName === "tab") { line += "\t"; continue; }
          if (["br", "cr"].includes(child.localName)) { line += "\n"; continue; }
        }
        visit(child);
      }
    };
    visit(p);
    return line;
  }).join("\n");
  return { text };
}

async function createCleanDocx(wiped) {
  const { Document, Packer, Paragraph, TextRun } = window.docx;
  const doc = new Document({
    creator: "PIIWipe",
    sections: [{ children: wiped.split(/\r?\n/).map((text) =>
      new Paragraph({ children: [new TextRun(text)] })) }],
  });
  return Packer.toArrayBuffer(doc);
}

let inputVersion = 0;

function resetOutput(name = "paste.txt") {
  inputVersion += 1;
  state = { name, kind: "text", original: "", wiped: "",
    counts: { email: 0, phone: 0, ssn: 0, card: 0, ip: 0 }, docxZip: null };
  setError("");
  render();
  return inputVersion;
}

async function handleFile(file) {
  const ticket = resetOutput(file.name || "file.txt");
  if (file.size > MAX_BYTES) {
    setError("File is over the 8 MB cap.");
    return;
  }
  const name = file.name || "file.txt";
  const lower = name.toLowerCase();
  let next;
  try {
    if (lower.endsWith(".pdf")) {
      const buf = await readAsArrayBuffer(file);
      const original = await extractPdfText(buf);
      const { wiped, counts } = wipeText(original);
      next = { name, kind: "pdf", original, wiped, counts, docxZip: null };
    } else if (lower.endsWith(".docx")) {
      const buf = await readAsArrayBuffer(file);
      const { text } = await extractDocxText(buf);
      const { wiped, counts } = wipeText(text);
      next = { name, kind: "docx", original: text, wiped, counts, docxZip: null };
    } else {
      const original = await readAsText(file);
      const { wiped, counts } = wipeText(original);
      next = { name, kind: "text", original, wiped, counts, docxZip: null };
    }
    if (ticket !== inputVersion) return;
    state = next;
    render();
  } catch (err) {
    if (ticket !== inputVersion) return;
    setError(err?.message || "Could not read that file.");
  }
}

function handlePaste() {
  const original = els.paste.value;
  resetOutput();
  const { wiped, counts } = wipeText(original);
  state = { name: "paste.txt", kind: "text", original, wiped, counts, docxZip: null };
  setError("");
  render();
}

function downloadBlob(blob, filename) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

function wipedName(name, ext) {
  const base = name.replace(/\.[^.]+$/, "");
  return `${base}.wiped.${ext}`;
}

async function downloadClean() {
  if (!state.wiped) return;
  if (state.kind === "docx") {
    const bytes = await createCleanDocx(state.wiped);
    const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    downloadBlob(blob, wipedName(state.name, "docx"));
    return;
  }
  if (state.kind === "pdf" && window.PDFLib) {
    const doc = await window.PDFLib.PDFDocument.create();
    const font = await doc.embedFont(window.PDFLib.StandardFonts.Helvetica);
    let page = doc.addPage();
    const { width, height } = page.getSize();
    const size = 11;
    const lines = state.wiped.split(/\n/);
    let y = height - 48;
    for (const line of lines) {
      const wrapped = wrapLine(line, 92);
      for (const w of wrapped) {
        if (y < 48) {
          page = doc.addPage();
          y = height - 48;
        }
        page.drawText(w, { x: 48, y, size, font, maxWidth: width - 96 });
        y -= 14;
      }
    }
    const bytes = await doc.save();
    downloadBlob(new Blob([bytes], { type: "application/pdf" }), wipedName(state.name, "pdf"));
    return;
  }
  const ext = (state.name.split(".").pop() || "txt").toLowerCase();
  const safeExt = ["txt", "csv", "md", "json", "html", "log"].includes(ext) ? ext : "txt";
  downloadBlob(new Blob([state.wiped], { type: "text/plain" }), wipedName(state.name, safeExt));
}

function wrapLine(line, width) {
  if (line.length <= width) return [line || " "];
  const out = [];
  let rest = line;
  while (rest.length > width) {
    let cut = rest.lastIndexOf(" ", width);
    if (cut < 20) cut = width;
    out.push(rest.slice(0, cut));
    rest = rest.slice(cut).trimStart();
  }
  if (rest) out.push(rest);
  return out;
}

els.pick.addEventListener("click", () => els.file.click());
els.file.addEventListener("change", () => {
  const f = els.file.files?.[0];
  if (f) handleFile(f);
});
els.paste.addEventListener("input", handlePaste);
els.sample.addEventListener("click", () => {
  els.paste.value = `Invoice for Jane Doe
Email: jane.doe@example.com
Phone: (415) 555-0199
SSN: 078-05-1120
Card: 4111 1111 1111 1111
Server: 192.168.1.24
Please send the signed PDF back.`;
  handlePaste();
});
els.download.addEventListener("click", () => downloadClean().catch((e) => setError(e.message)));
els.copy.addEventListener("click", async () => {
  await navigator.clipboard.writeText(state.wiped);
  els.status.textContent = "Wiped text copied. Still never left this tab.";
});

["dragenter", "dragover"].forEach((ev) => {
  els.drop.addEventListener(ev, (e) => {
    e.preventDefault();
    els.drop.classList.add("over");
  });
});
["dragleave", "drop"].forEach((ev) => {
  els.drop.addEventListener(ev, (e) => {
    e.preventDefault();
    els.drop.classList.remove("over");
  });
});
els.drop.addEventListener("drop", (e) => {
  const f = e.dataTransfer?.files?.[0];
  if (f) handleFile(f);
});
