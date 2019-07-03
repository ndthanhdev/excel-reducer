const path = require("path");
const fs = require("fs");
const { promisify } = require("util");
const XLSX = require("xlsx");
const instructions = require("./instructions.json");

const readdirAsync = promisify(fs.readdir);

const INPUT_PATH = "./input";

function getPaths() {
  return readdirAsync(INPUT_PATH).then(names =>
    names.map(name => path.resolve(INPUT_PATH, name))
  );
}

function readWorkbook(sPath) {
  return XLSX.readFile(sPath);
}

function toSheet(workbook) {
  return workbook.Sheets["HUYỆN"];
}

function toRow(sheet) {
  const sources = instructions.sources;

  function transform(value, transformer) {
    const script = String.prototype.replace.call(
      transformer,
      "__VALUE__",
      value
    );
    return eval(script);
  }

  const row = sources.map(instruct => {
    const raw = sheet[instruct[0]] && sheet[instruct[0]].v;
    const result = instruct.length > 1 ? transform(raw, instruct[1]) : raw;
    return result;
  });

  return row;
}

function rowsWorkBook(rows) {
  const header = instructions.header;
  const sheetName = instructions.sheetName;

  const wb = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([header, ...rows]);

  XLSX.utils.book_append_sheet(wb, sheet, sheetName);

  return wb;
}

function writeResultToFile(wb) {
  return XLSX.writeFile(wb, "./output.xlsx");
}

function processWb(sPath) {
  return Promise.resolve(sPath)
    .then(readWorkbook)
    .then(toSheet)
    .then(toRow)
    .catch(e => {
      console.error(`Error ${sPath}: ${JSON.stringify(e)}`);
      return undefined;
    });
}

const main = () =>
  Promise.resolve()
    .then(getPaths)
    .then(paths => Promise.all(paths.map(p => processWb(p))))
    .then(rows => rows.filter(row => row !== undefined)) // filter failed
    .then(rowsWorkBook)
    .then(writeResultToFile);

async function boundary() {
  main()
    .then(() => console.log("done"))
    .catch(e => console.error(e));
}

(async () => await boundary())();
