/**
 * Generator script for fixture files.
 *
 * - Generates large-dataset.csv (1000 rows)
 * - Reads all CSV files → csvToSheet → createWorkbook → write → saves XLSX
 *
 * Run via: pnpm run generate-fixtures
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { csvToSheet, createWorkbook, writeFile } from "../../src/index.js";

const fixturesDir = path.dirname(new URL(import.meta.url).pathname);
const csvDir = path.join(fixturesDir, "csv");
const xlsxDir = path.join(fixturesDir, "xlsx");

// Ensure output directory exists
fs.mkdirSync(xlsxDir, { recursive: true });

// --- Generate large-dataset.csv ---
const LARGE_ROW_COUNT = 1000;
const names = ["Alice", "Bob", "Charlie", "Diana", "Eve", "Frank", "Grace", "Hank", "Iris", "Jack"];

function generateLargeCsv(): string {
	const lines: string[] = ["ID,Name,Value,Flag"];
	for (let i = 1; i <= LARGE_ROW_COUNT; i++) {
		const name = names[i % names.length];
		const value = Math.round((i * 3.14 + 0.001) * 100) / 100;
		const flag = i % 3 === 0 ? "TRUE" : "FALSE";
		lines.push(`${i},${name},${value},${flag}`);
	}
	return lines.join("\n");
}

const largeCsvPath = path.join(csvDir, "large-dataset.csv");
console.log("Generating large-dataset.csv ...");
fs.writeFileSync(largeCsvPath, generateLargeCsv(), "utf-8");

// --- Convert all CSVs to XLSX ---
const csvFiles = fs.readdirSync(csvDir).filter((f) => f.endsWith(".csv"));

for (const csvFile of csvFiles) {
	const csvPath = path.join(csvDir, csvFile);
	const xlsxFile = csvFile.replace(/\.csv$/, ".xlsx");
	const xlsxPath = path.join(xlsxDir, xlsxFile);

	console.log(`${csvFile} → ${xlsxFile}`);
	const text = fs.readFileSync(csvPath, "utf-8");
	const ws = csvToSheet(text);
	const wb = createWorkbook(ws, "Sheet1");
	await writeFile(wb, xlsxPath);
}

console.log("Done. Generated XLSX files in tests/fixtures/xlsx/");
