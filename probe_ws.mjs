import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const input = await FileBlob.load("outputs/DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V4B_NATURE_DASHBOARD.xlsx");
const wb = await SpreadsheetFile.importXlsx(input);
const first = wb.worksheets.getItemAt ? wb.worksheets.getItemAt(0) : null;
let maxId = 0, n=0;
for (const ws of wb.worksheets) {
  n++;
  const id = ws.sheetId ?? ws.id ?? "?";
  if (typeof id === "number" && id>maxId) maxId=id;
}
console.log("count", n, "maxId", maxId);
// add a test sheet and inspect
const t = wb.worksheets.add("ZZ_TEST_PROBE");
console.log("new sheet keys:", Object.keys(t));
console.log("new sheetId before:", t.sheetId, "id:", t.id);
const proto = Object.getOwnPropertyNames(Object.getPrototypeOf(t)).filter(k=>/sheet|id|index|position/i.test(k));
console.log("proto id-ish:", proto);
try { t.sheetId = maxId+1; console.log("set sheetId ->", t.sheetId); } catch(e){ console.log("set sheetId err", e.message); }
