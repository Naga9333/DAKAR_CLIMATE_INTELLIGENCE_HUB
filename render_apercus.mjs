import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
import fs from "node:fs/promises";

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load("outputs/DAKAR_CLIMATE_BASE_PAR_ANNEE.xlsx"));

const shots = [
  { sheet: "00_LISEZ_MOI", range: "A1:H30", file: "APERCU_00_LISEZ_MOI.png" },
  { sheet: "2024", range: "A1:H44", file: "APERCU_ANNEE_2024.png" },
  { sheet: "SYNTHESE_DONNEES", range: "A1:N32", file: "APERCU_SYNTHESE_DONNEES.png" },
  { sheet: "VISUALISATION_PROGRESSIVE", range: "A1:T120", file: "APERCU_VISUALISATION.png" },
];

for (const s of shots) {
  try {
    const blob = await wb.render({ sheet: s.sheet, range: s.range });
    const buf = Buffer.from(await blob.arrayBuffer());
    await fs.writeFile(`outputs/${s.file}`, buf);
    console.log("OK", s.file, `(${s.sheet} ${s.range})`, buf.length, "bytes", blob.type);
  } catch (e) {
    console.log("ERR", s.file, e.message.slice(0, 160));
  }
}
