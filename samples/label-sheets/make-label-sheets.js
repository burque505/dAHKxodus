// Faux label stock: one PDF per product, each label's outline (and a 1/16 in inner guide) where the real
// sheet's labels are, for testing Mailings > Labels without label stock. Run: node make-label-sheets.js
const fs = require("fs");
const path = require("path");

const IN = 72, MM = 72 / 25.4;
const LETTER = [8.5 * IN, 11 * IN], A4 = [210 * MM, 297 * MM];
// name, page, label w/h, across, down, top, side, h pitch, v pitch (points)
const SHEETS = [
  ["Avery-5160-address-1x2.625in-30", LETTER, 2.625 * IN, 1 * IN, 3, 10, 0.5 * IN, 0.1875 * IN, 2.75 * IN, 1 * IN],
  ["Avery-5161-address-1x4in-20", LETTER, 4 * IN, 1 * IN, 2, 10, 0.5 * IN, 0.15625 * IN, 4.1875 * IN, 1 * IN],
  ["Avery-5162-address-1.333x4in-14", LETTER, 4 * IN, 1.3333 * IN, 2, 7, 0.8333 * IN, 0.15625 * IN, 4.1875 * IN, 1.3333 * IN],
  ["Avery-5163-shipping-2x4in-10", LETTER, 4 * IN, 2 * IN, 2, 5, 0.5 * IN, 0.15625 * IN, 4.1875 * IN, 2 * IN],
  ["Avery-5167-return-0.5x1.75in-80", LETTER, 1.75 * IN, 0.5 * IN, 4, 20, 0.5 * IN, 0.3 * IN, 2.05 * IN, 0.5 * IN],
  ["Avery-L7160-63.5x38.1mm-21", A4, 63.5 * MM, 38.1 * MM, 3, 7, 15.15 * MM, 7.21 * MM, 66.04 * MM, 38.1 * MM],
  ["Avery-L7163-99.1x38.1mm-14", A4, 99.1 * MM, 38.1 * MM, 2, 7, 15.15 * MM, 4.65 * MM, 101.6 * MM, 38.1 * MM],
];

const f = (n) => n.toFixed(2);
for (const [name, [pw, ph], w, h, across, down, top, side, hp, vp] of SHEETS) {
  let c = `0.6 w 0 0 0 RG\n`;
  for (let r = 0; r < down; r++)
    for (let k = 0; k < across; k++) {
      const x = side + k * hp, y = ph - top - r * vp - h;
      c += `${f(x)} ${f(y)} ${f(w)} ${f(h)} re S\n`;
      // (an inner dashed guide 1/16 in in: text should stay inside it)
      const g = IN / 16;
      c += `[2 2] 0 d 0.3 w 0.6 0.6 0.6 RG ${f(x + g)} ${f(y + g)} ${f(w - 2 * g)} ${f(h - 2 * g)} re S [] 0 d 0.6 w 0 0 0 RG\n`;
      c += `BT /F1 5 Tf 0.6 0.6 0.6 rg ${f(x + w - 14)} ${f(y + 2)} Td (${r * across + k + 1}) Tj ET 0 0 0 rg\n`;
    }
  c += `BT /F1 7 Tf ${f(side)} ${f(Math.max(4, ph - top + 6))} Td (FAUX ${name.replace(/-/g, " ")} - print at 100% / Actual size) Tj ET\n`;
  const objs = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${f(pw)} ${f(ph)}] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>`,
    `<< /Length ${Buffer.byteLength(c)} >>\nstream\n${c}endstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ];
  let pdf = "%PDF-1.4\n";
  const offs = [];
  objs.forEach((o, i) => { offs.push(Buffer.byteLength(pdf)); pdf += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const xref = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n` + offs.map((o) => `${String(o).padStart(10, "0")} 00000 n \n`).join("");
  pdf += `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  fs.writeFileSync(path.join(__dirname, `${name}.pdf`), pdf);
  console.log(name);
}
