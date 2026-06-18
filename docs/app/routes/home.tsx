import { Link } from "react-router";
import {
	ArrowLeftRight,
	ArrowRight,
	ArrowUpRight,
	Braces,
	CheckCircle2,
	Code2,
	Download,
	ExternalLink,
	FileSpreadsheet,
	Globe,
	LockKeyhole,
	Package,
	Paintbrush,
	Rows3,
	ShieldCheck,
	Table2,
	Zap,
} from "lucide-react";
import config from "virtual:ardo/config";

const proofItems = [
	{ label: "Runtime deps", value: "0" },
	{ label: "API shape", value: "async named exports" },
	{ label: "Targets", value: "Node, browser, edge" },
	{ label: "Output", value: "XLSX, CSV, TSV, HTML" },
];

const comparisonRows = [
	{
		label: "Runtime dependencies",
		xlsx: "0",
		sheetjs: "7",
		exceljs: "9",
	},
	{
		label: "Read/write model",
		xlsx: "Fully async",
		sheetjs: "Sync",
		exceljs: "Partial async",
	},
	{
		label: "Module format",
		xlsx: "ESM + CJS",
		sheetjs: "CJS",
		exceljs: "CJS",
	},
	{
		label: "Browser exports",
		xlsx: "read() + write()",
		sheetjs: "Separate bundle",
		exceljs: "No",
	},
	{
		label: "Styled reports",
		xlsx: "Typed style layer",
		sheetjs: "Utility model",
		exceljs: "Workbook classes",
	},
];

const useCases = [
	{
		title: "Ship styled browser reports without ExcelJS weight",
		icon: <Paintbrush size={22} strokeWidth={1.8} />,
		description:
			"Create branded workbook downloads with fonts, fills, borders, number formats, merged titles, column widths, row heights, totals, and frozen panes.",
		link: "/guide/styled-workbooks",
		linkText: "See styled workbooks",
	},
	{
		title: "Migrate SheetJS code without redesigning your data model",
		icon: <ArrowLeftRight size={22} strokeWidth={1.8} />,
		description:
			"Cell objects keep the familiar shape. Add await, switch to named imports, and move from sheet_to_json to sheetToJson.",
		link: "/guide/migration",
		linkText: "Read the migration guide",
	},
	{
		title: "Keep exports safer when data comes from users",
		icon: <LockKeyhole size={22} strokeWidth={1.8} />,
		description:
			"CSV formula-like fields are escaped by default, and HTML export sanitizes unsafe link targets unless you explicitly opt out.",
		link: "/guide/security",
		linkText: "Review security guidance",
	},
	{
		title: "Convert spreadsheet data both ways",
		icon: <Braces size={22} strokeWidth={1.8} />,
		description:
			"Move between worksheets, JSON objects, arrays of arrays, CSV, TSV, and HTML tables with typed utilities that work in modern runtimes.",
		link: "/guide/getting-started",
		linkText: "Build your first workbook",
	},
];

const capabilityGroups = [
	{
		title: "Workbook core",
		icon: <FileSpreadsheet size={20} strokeWidth={1.8} />,
		items: ["Multiple sheets", "Defined names", "Document properties", "Date systems"],
	},
	{
		title: "Cell data",
		icon: <Table2 size={20} strokeWidth={1.8} />,
		items: ["Strings", "Numbers", "Dates", "Booleans", "Errors"],
	},
	{
		title: "Report styling",
		icon: <Paintbrush size={20} strokeWidth={1.8} />,
		items: ["Fonts", "Fills", "Borders", "Alignment", "Number formats"],
	},
	{
		title: "Sheet structure",
		icon: <Rows3 size={20} strokeWidth={1.8} />,
		items: ["Merged cells", "Column widths", "Row heights", "Frozen panes", "Auto filters"],
	},
	{
		title: "Formulas and notes",
		icon: <Code2 size={20} strokeWidth={1.8} />,
		items: ["Cell formulas", "Array formulas", "Shared formula read", "Legacy comments", "Threaded comments"],
	},
	{
		title: "Export guards",
		icon: <ShieldCheck size={20} strokeWidth={1.8} />,
		items: ["CSV formula escaping", "HTML link sanitizing", "XlsxError codes", "ZIP and XML limits"],
	},
];

const heroCode = `import {
  arrayToSheet,
  createWorkbook,
  styleRange,
  write,
} from "xlsx-format";

const sheet = arrayToSheet(reportRows);

styleRange(sheet, "A2:D2", headerStyle);

const bytes = await write(createWorkbook(sheet, "Q2"), {
  type: "array",
  cellStyles: true,
});`;

export default function HomePage() {
	const version = config.project?.version ?? "0.0.0";

	return (
		<div className="xlsx-home">
			<section className="xlsx-hero" aria-labelledby="xlsx-home-title">
				<div className="xlsx-hero__content">
					<p className="xlsx-eyebrow">xlsx-format v{version} - zero dependency XLSX for modern apps</p>
					<h1 id="xlsx-home-title">Replace spreadsheet weight with an XLSX core built for modern apps.</h1>
					<p className="xlsx-hero__lead">
						Read, write, convert, and style modern Excel workbooks with strict TypeScript, fully async APIs,
						zero runtime dependencies, and browser-ready output.
					</p>

					<div className="xlsx-hero__actions" aria-label="Primary actions">
						<Link className="xlsx-button xlsx-button--primary" to="/guide/getting-started">
							Build your first workbook
							<ArrowRight size={17} aria-hidden="true" />
						</Link>
						<Link className="xlsx-button xlsx-button--secondary" to="/guide/migration">
							Compare the migration
							<ArrowUpRight size={17} aria-hidden="true" />
						</Link>
						<a
							className="xlsx-button xlsx-button--ghost"
							href="https://github.com/sebastian-software/xlsx-format"
							aria-label="Open xlsx-format on GitHub"
						>
							GitHub
							<ExternalLink size={16} aria-hidden="true" />
						</a>
					</div>

					<dl className="xlsx-proof" aria-label="Product proof points">
						{proofItems.map((item) => (
							<div key={item.label} className="xlsx-proof__item">
								<dt>{item.label}</dt>
								<dd>{item.value}</dd>
							</div>
						))}
					</dl>
				</div>

				<div className="xlsx-product-shot" aria-label="Styled workbook export preview">
					<div className="xlsx-product-shot__header">
						<span>report-builder.ts</span>
						<span>async XLSX export</span>
					</div>
					<pre className="xlsx-code-panel">
						<code>{heroCode}</code>
					</pre>
					<div className="xlsx-workbook-preview" aria-hidden="true">
						<div className="xlsx-workbook-preview__bar">
							<span>Overview</span>
							<span>Q2 report.xlsx</span>
						</div>
						<div className="xlsx-sheet">
							<div className="xlsx-sheet__title">Northstar Solar PPA - Q2 Report</div>
							<div className="xlsx-sheet__row xlsx-sheet__row--head">
								<span>Month</span>
								<span>Expected MWh</span>
								<span>Actual MWh</span>
								<span>Settlement</span>
							</div>
							<div className="xlsx-sheet__row">
								<span>Apr 2026</span>
								<span>12,400</span>
								<span>12,050</span>
								<span>-18,350</span>
							</div>
							<div className="xlsx-sheet__row">
								<span>May 2026</span>
								<span>13,100</span>
								<span>13,980</span>
								<span>44,200</span>
							</div>
							<div className="xlsx-sheet__row xlsx-sheet__row--total">
								<span>Q2 Total</span>
								<span>39,700</span>
								<span>40,490</span>
								<span>39,370</span>
							</div>
						</div>
					</div>
				</div>
			</section>

			<section className="xlsx-section xlsx-switch" aria-labelledby="switch-title">
				<div className="xlsx-section__intro">
					<p className="xlsx-kicker">Why switch?</p>
					<h2 id="switch-title">Most apps do not need a spreadsheet framework.</h2>
					<p>
						If the job is XLSX, CSV, TSV, HTML, and polished report exports, xlsx-format keeps the useful
						parts close and leaves the legacy weight out of your runtime.
					</p>
				</div>

				<div className="xlsx-comparison" role="table" aria-label="Library comparison">
					<div className="xlsx-comparison__header" role="row">
						<span role="columnheader">Decision point</span>
						<strong role="columnheader">xlsx-format</strong>
						<span role="columnheader">SheetJS</span>
						<span role="columnheader">ExcelJS</span>
					</div>
					{comparisonRows.map((row) => (
						<div key={row.label} className="xlsx-comparison__row" role="row">
							<span role="cell">{row.label}</span>
							<strong role="cell">
								<CheckCircle2 size={17} aria-hidden="true" />
								{row.xlsx}
							</strong>
							<span role="cell">{row.sheetjs}</span>
							<span role="cell">{row.exceljs}</span>
						</div>
					))}
				</div>
			</section>

			<section className="xlsx-section xlsx-use-cases" aria-labelledby="use-cases-title">
				<div className="xlsx-section__intro xlsx-section__intro--wide">
					<p className="xlsx-kicker">Built for real export work</p>
					<h2 id="use-cases-title">One lean XLSX layer for the places spreadsheets actually show up.</h2>
				</div>

				<div className="xlsx-use-case-list">
					{useCases.map((useCase) => (
						<article key={useCase.title} className="xlsx-use-case">
							<div className="xlsx-use-case__icon" aria-hidden="true">
								{useCase.icon}
							</div>
							<div>
								<h3>{useCase.title}</h3>
								<p>{useCase.description}</p>
							</div>
							<Link className="xlsx-inline-link" to={useCase.link}>
								{useCase.linkText}
								<ArrowRight size={16} aria-hidden="true" />
							</Link>
						</article>
					))}
				</div>
			</section>

			<section className="xlsx-section xlsx-runtime" aria-labelledby="runtime-title">
				<div className="xlsx-runtime__panel">
					<div className="xlsx-section__intro">
						<p className="xlsx-kicker">Runs where your app runs</p>
						<h2 id="runtime-title">No Node-only core. No browser-only fork.</h2>
						<p>
							xlsx-format reads and writes in-memory data sources: Uint8Array, ArrayBuffer, Node Buffer,
							base64 strings, binary strings, CSV text, and HTML tables.
						</p>
					</div>
					<div className="xlsx-runtime__grid">
						<div>
							<Globe size={21} aria-hidden="true" />
							<strong>Browser UI</strong>
							<span>File inputs, fetch responses, Blob downloads</span>
						</div>
						<div>
							<Zap size={21} aria-hidden="true" />
							<strong>Server jobs</strong>
							<span>Async read/write around your own file I/O</span>
						</div>
						<div>
							<Package size={21} aria-hidden="true" />
							<strong>Edge paths</strong>
							<span>No node:fs import, no runtime dependency chain</span>
						</div>
					</div>
				</div>
			</section>

			<section className="xlsx-section xlsx-capabilities" aria-labelledby="capabilities-title">
				<div className="xlsx-section__intro xlsx-section__intro--wide">
					<p className="xlsx-kicker">Feature surface</p>
					<h2 id="capabilities-title">Enough spreadsheet power for production reports, kept deliberately focused.</h2>
				</div>

				<div className="xlsx-capability-grid">
					{capabilityGroups.map((group) => (
						<article key={group.title} className="xlsx-capability">
							<div className="xlsx-capability__heading">
								<span aria-hidden="true">{group.icon}</span>
								<h3>{group.title}</h3>
							</div>
							<ul>
								{group.items.map((item) => (
									<li key={item}>{item}</li>
								))}
							</ul>
						</article>
					))}
				</div>
			</section>

			<section className="xlsx-final-cta" aria-labelledby="final-cta-title">
				<div>
					<p className="xlsx-kicker">Install the smaller spreadsheet layer</p>
					<h2 id="final-cta-title">Start with async XLSX today. Keep the workbook framework out of the bundle.</h2>
				</div>
				<div className="xlsx-install">
					<code>npm install xlsx-format</code>
					<Link className="xlsx-button xlsx-button--primary" to="/guide/getting-started">
						Create your first workbook
						<Download size={17} aria-hidden="true" />
					</Link>
				</div>
			</section>
		</div>
	);
}
