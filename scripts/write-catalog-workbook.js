const AdmZip = require("adm-zip");
const fs = require("fs");
const os = require("os");
const path = require("path");

const ADDIN_ID = "7c1f5f2e-8a4b-4d3c-9e6f-2b0a1c9d8e7f";
const store = `\\\\${os.hostname()}\\ExcelAIAgentCatalog`;
const projectDir = path.join(__dirname, "..");
const template = path.join(
  projectDir,
  "node_modules",
  "office-addin-dev-settings",
  "templates",
  "ExcelWorkbookWithTaskPane.xlsx"
);
const zip = new AdmZip(template);
const xml =
  `<?xml version="1.0" encoding="utf-8"?>` +
  `<we:webextension xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
  `xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{${ADDIN_ID}}">` +
  `<we:reference id="${ADDIN_ID}" version="1.0.1.0" store="${store}" storeType="FileSystem"/>` +
  `<we:alternateReferences/>` +
  `<we:properties/>` +
  `<we:bindings/>` +
  `<we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>` +
  `</we:webextension>`;
zip.updateFile("xl/webextensions/webextension.xml", Buffer.from(xml, "utf8"));

const openFile = path.join(projectDir, "ExcelAIAgent.xlsx");
zip.writeZip(openFile);
fs.mkdirSync(path.join(process.env.LOCALAPPDATA, "ExcelAIAgent", "catalog"), { recursive: true });
fs.copyFileSync(path.join(projectDir, "manifest.xml"), path.join(process.env.LOCALAPPDATA, "ExcelAIAgent", "catalog", "manifest.xml"));
console.log(zip.readAsText("xl/webextensions/webextension.xml"));
console.log("openFile", openFile);
