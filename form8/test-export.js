/* Dev-only test: index.html ના <script> ને Node માં ચલાવી, xlsx બનાવી, ચેક કરે છે. */
const fs = require("fs");
const vm = require("vm");
const path = require("path");

const html = fs.readFileSync(path.join(__dirname, "index.html"), "utf8");
const m = html.match(/<script>([\s\S]*?)<\/script>/);
if (!m) throw new Error("script block not found");

const elStub = () => ({
  value: "", textContent: "", innerHTML: "", style: {}, classList: { add(){}, remove(){} },
  appendChild(){}, addEventListener(){}, removeChild(){}, setAttribute(){}, click(){},
  files: null, parentNode: null
});

const sandbox = {
  console,
  window: { TextEncoder },
  TextEncoder,
  Blob,
  Uint8Array, Uint32Array, DataView, Array, Object, JSON, String, Number, Math, Date, isFinite, parseFloat, isNaN, Error,
  setTimeout, clearTimeout,
  document: {
    querySelector: () => elStub(),
    createElement: () => elStub(),
    addEventListener: () => {},
    body: elStub()
  },
  localStorage: { getItem: () => null, setItem: () => {} }
};
vm.createContext(sandbox);
vm.runInContext(m[1], sandbox);

// ડેમો ડેટા
vm.runInContext(`
  state.facility = "CHC ધાનેરા";
  state.block = "ધાનેરા";
  state.worker = "R. Patel";
  state.month = 10;
  state.year = 2026;
  state.rows = [
    { id: "r1", date: "1", opd_total: "120", opd_new: "80", opd_fu: "40", ipd_in: "5", ipd_total: "12",
      referral: "3", lab: "25", delivery: "2", immun: "18", remark: "OK & fine" },
    { id: "r2", date: "2", opd_total: "95", opd_new: "60", opd_fu: "35", ipd_in: "4", ipd_total: "10",
      referral: "1", lab: "19", delivery: "1", immun: "22", remark: "રજા" },
    { id: "r3", date: "3", opd_total: "૧૧૦", opd_new: "70", opd_fu: "40", ipd_in: "6", ipd_total: "14",
      referral: "2", lab: "30", delivery: "0", immun: "12", remark: "" }
  ];
`, sandbox);

(async () => {
  const blob = vm.runInContext("buildXlsxBlob()", sandbox);
  const buf = Buffer.from(await blob.arrayBuffer());
  const out = path.join(__dirname, "test_out.xlsx");
  fs.writeFileSync(out, buf);
  console.log("wrote", out, buf.length, "bytes");

  const csv = vm.runInContext("buildCsv()", sandbox);
  fs.writeFileSync(path.join(__dirname, "test_out.csv"), csv, "utf8");
  console.log("csv bytes:", Buffer.byteLength(csv, "utf8"));
  console.log("totals:", JSON.stringify(vm.runInContext("totals()", sandbox)));
})().catch(e => { console.error("FAIL", e); process.exit(1); });
