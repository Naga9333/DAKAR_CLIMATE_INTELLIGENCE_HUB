import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const input = await FileBlob.load("outputs/DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V4B_NATURE_DASHBOARD.xlsx");
const wb = await SpreadsheetFile.importXlsx(input);
const wss = wb.worksheets;
console.log("Total worksheets:", wss.length);
for (const ws of wss) {
  const proto = Object.getOwnPropertyNames(Object.getPrototypeOf(ws));
  const dims = ws.dimensions || (ws.getDimensions && ws.getDimensions());
  console.log(`- ${ws.name}`);
}
