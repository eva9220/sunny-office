// 讀取財產清冊 Excel（.xlsx），含「放在儲存格內」與「浮動」的照片
// 依標題列的欄名對應欄位，回傳 [{c, f:{no,name,qty,price,date,pub,mark,note}, ph:[Blob]}]

async function unzip(buf) {
  const dv = new DataView(buf), u8 = new Uint8Array(buf);
  let e = buf.byteLength - 22;
  while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error("not zip");
  const n = dv.getUint16(e + 10, true);
  let p = dv.getUint32(e + 16, true);
  const files = {}, td = new TextDecoder();
  for (let i = 0; i < n; i++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true), size = dv.getUint32(p + 20, true);
    const nl = dv.getUint16(p + 28, true), xl = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
    const off = dv.getUint32(p + 42, true), name = td.decode(u8.subarray(p + 46, p + 46 + nl));
    files[name] = { method, size, off };
    p += 46 + nl + xl + cl;
  }
  async function bytes(name) {
    const f = files[name];
    if (!f) return null;
    const s = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
    const raw = u8.slice(s, s + f.size);
    if (f.method === 0) return raw;
    const st = new Blob([raw]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(st).arrayBuffer());
  }
  return { has: (n) => !!files[n], bytes, text: async (n) => { const b = await bytes(n); return b ? td.decode(b) : null; } };
}

const tags = (node, name) => node ? Array.from(node.getElementsByTagNameNS("*", name)) : [];
const RID = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
function resolve(base, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/");
  parts.pop();
  for (const t of target.split("/")) {
    if (t === "..") parts.pop(); else if (t !== ".") parts.push(t);
  }
  return parts.join("/");
}
function colNum(ref) {
  let c = 0;
  for (const ch of ref.match(/^[A-Z]+/)[0]) c = c * 26 + ch.charCodeAt(0) - 64;
  return c - 1;
}
const MIME = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp", bmp: "image/bmp" };

function toDate(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number" || /^\d{5}(\.\d+)?$/.test(v)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(+v) * 864e5);
    return d.toISOString().slice(0, 10);
  }
  const m = String(v).trim().match(/^(\d{2,4})[.\/\-年](\d{1,2})[.\/\-月](\d{1,2})/);
  if (!m) return String(v);
  let y = +m[1];
  if (y < 1000) y += 1911; // 民國年
  return y + "-" + String(m[2]).padStart(2, "0") + "-" + String(m[3]).padStart(2, "0");
}

const FIELDS = [
  ["no", /編號|序號/], ["name", /名稱|書名|品名/], ["qty", /數量/], ["price", /單價|價格|金額/],
  ["date", /日期/], ["pub", /出版|廠牌|品牌/], ["mark", /標章/], ["note", /備註/], ["c", /分類|類別/]
];

export async function readXlsx(file) {
  const z = await unzip(await file.arrayBuffer());
  const xml = async (n) => { const t = await z.text(n); return t ? new DOMParser().parseFromString(t, "application/xml") : null; };
  const rels = async (n) => {
    const d = await xml(n.replace(/([^\/]+)$/, "_rels/$1.rels")), m = {};
    tags(d, "Relationship").forEach((r) => { m[r.getAttribute("Id")] = resolve(n, r.getAttribute("Target")); });
    return m;
  };

  // 第一個工作表
  const wb = await xml("xl/workbook.xml");
  const wbr = await rels("xl/workbook.xml");
  const sheetPath = wbr[tags(wb, "sheet")[0].getAttributeNS(RID, "id")] || "xl/worksheets/sheet1.xml";
  const sheet = await xml(sheetPath);
  const ss = tags(await xml("xl/sharedStrings.xml"), "si").map((si) => tags(si, "t").map((t) => t.textContent).join(""));

  // 放在儲存格內的照片：vm → metadata → richvalue → richValueRel → media
  let cellImg = () => null;
  if (z.has("xl/metadata.xml") && z.has("xl/richData/rdrichvalue.xml")) {
    const md = await xml("xl/metadata.xml");
    const rvb = tags(md, "rvb").map((e) => +e.getAttribute("i"));
    const vmBk = tags(tags(md, "valueMetadata")[0], "bk").map((bk) => +tags(bk, "rc")[0].getAttribute("v"));
    const keys = tags(await xml("xl/richData/rdrichvaluestructure.xml"), "s").map((s) => tags(s, "k").map((k) => k.getAttribute("n")));
    const rv = tags(await xml("xl/richData/rdrichvalue.xml"), "rv").map((r) => {
      const k = keys[+r.getAttribute("s")] || [], i = Math.max(0, k.indexOf("_rvRel:LocalImageIdentifier"));
      return +tags(r, "v")[i].textContent;
    });
    const relIds = tags(await xml("xl/richData/richValueRel.xml"), "rel").map((r) => r.getAttributeNS(RID, "id"));
    const rr = await rels("xl/richData/richValueRel.xml");
    cellImg = (vm) => { try { return rr[relIds[rv[rvb[vmBk[vm - 1]]]]] || null; } catch (e) { return null; } };
  }

  // 讀儲存格
  const grid = {}, imgs = {};
  const addImg = (r, c, path) => { if (path) (imgs[r] = imgs[r] || []).push({ c, path }); };
  tags(sheet, "c").forEach((cell) => {
    const ref = cell.getAttribute("r"), r = +ref.match(/\d+/)[0], c = colNum(ref), t = cell.getAttribute("t");
    if (cell.getAttribute("vm")) addImg(r, c, cellImg(+cell.getAttribute("vm")));
    let v = tags(cell, "v")[0];
    v = v ? v.textContent : null;
    if (t === "s") v = ss[+v];
    else if (t === "inlineStr") v = tags(cell, "t").map((x) => x.textContent).join("");
    else if (t === "e") v = null;
    else if (v != null && t !== "str" && t !== "b" && v !== "" && !isNaN(v)) v = +v;
    if (v != null && v !== "") (grid[r] = grid[r] || {})[c] = v;
  });

  // 浮動在儲存格上的照片
  const sr = await rels(sheetPath);
  for (const d of tags(sheet, "drawing")) {
    const dp = sr[d.getAttributeNS(RID, "id")];
    if (!dp) continue;
    const dx = await xml(dp), dr = await rels(dp);
    ["twoCellAnchor", "oneCellAnchor"].forEach((a) => tags(dx, a).forEach((an) => {
      const from = tags(an, "from")[0], blip = tags(an, "blip")[0];
      if (!from || !blip) return;
      addImg(+tags(from, "row")[0].textContent + 1, +tags(from, "col")[0].textContent, dr[blip.getAttributeNS(RID, "embed")]);
    }));
  }

  // 找標題列（有「名稱」的那一列）
  const rowsNo = Object.keys(grid).map(Number).sort((a, b) => a - b);
  const head = rowsNo.find((r) => Object.values(grid[r]).some((v) => typeof v === "string" && FIELDS[1][1].test(v)));
  if (!head) throw new Error("找不到標題列（需要有「名稱」欄）");
  const col = {};
  Object.entries(grid[head]).forEach(([c, v]) => {
    const f = FIELDS.find((x) => x[1].test(String(v)));
    if (f && col[f[0]] == null) col[f[0]] = +c;
  });

  const out = [];
  const lastRow = Math.max(head, ...rowsNo, ...Object.keys(imgs).map(Number));
  for (let r = head + 1; r <= lastRow; r++) {
    const g = grid[r] || {}, name = g[col.name];
    if (name == null || String(name).trim() === "") continue;
    const f = {};
    ["no", "name", "qty", "price", "pub", "mark", "note"].forEach((k) => {
      if (col[k] != null && g[col[k]] != null) f[k] = String(g[col[k]]).trim();
    });
    if (col.date != null) f.date = toDate(g[col.date]);
    const ph = [];
    for (const im of (imgs[r] || []).sort((a, b) => a.c - b.c)) {
      const b = await z.bytes(im.path);
      if (b) ph.push(new Blob([b], { type: MIME[im.path.split(".").pop().toLowerCase()] || "image/jpeg" }));
    }
    out.push({ c: col.c != null && g[col.c] != null ? String(g[col.c]).trim() : "", f, ph });
  }
  return out;
}
