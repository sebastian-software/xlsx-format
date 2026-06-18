import { useEffect } from "react";
import { Link } from "react-router";
import { ArrowRight, ArrowUpRight, Check, Download, Minus } from "lucide-react";
import config from "virtual:ardo/config";

function GithubMark({ size = 16 }: { size?: number }) {
	return (
		<svg width={size} height={size} viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
			<path d="M12 .5A11.5 11.5 0 0 0 .5 12a11.5 11.5 0 0 0 7.86 10.92c.58.1.79-.25.79-.56v-2c-3.2.7-3.88-1.37-3.88-1.37-.53-1.34-1.3-1.7-1.3-1.7-1.05-.72.08-.7.08-.7 1.16.08 1.77 1.2 1.77 1.2 1.03 1.77 2.7 1.26 3.36.96.1-.75.4-1.26.73-1.55-2.55-.29-5.24-1.28-5.24-5.7 0-1.26.45-2.29 1.19-3.1-.12-.29-.52-1.46.11-3.05 0 0 .97-.31 3.18 1.18a11 11 0 0 1 5.8 0c2.2-1.49 3.17-1.18 3.17-1.18.63 1.59.23 2.76.11 3.05.74.81 1.19 1.84 1.19 3.1 0 4.43-2.7 5.4-5.26 5.69.41.36.78 1.07.78 2.16v3.2c0 .31.21.67.8.56A11.5 11.5 0 0 0 23.5 12 11.5 11.5 0 0 0 12 .5Z" />
		</svg>
	);
}

/* ------------------------------------------------------------------ */
/*  Content model                                                     */
/* ------------------------------------------------------------------ */

const proofItems = [
	{ value: "0", label: "runtime dependencies", cell: "B2" },
	{ value: "async", label: "named ESM + CJS exports", cell: "B3" },
	{ value: "91%", label: "test coverage", cell: "B4" },
	{ value: "4", label: "output formats", cell: "B5" },
];

const comparisonColumns = ["xlsx-format", "SheetJS", "ExcelJS"];

const comparisonRows = [
	{ label: "Runtime dependencies", values: ["0", "7", "9"], win: 0 },
	{ label: "Read / write model", values: ["Fully async", "Sync", "Partial async"], win: 0 },
	{ label: "Module format", values: ["ESM + CJS", "CJS", "CJS"], win: 0 },
	{ label: "Browser exports", values: ["read() + write()", "Separate bundle", "Not supported"], win: 0 },
	{ label: "Styled reports", values: ["Typed style layer", "Utility model", "Workbook classes"], win: 0 },
];

const useCases = [
	{
		id: "R01",
		title: "Styled browser reports, without ExcelJS",
		body: "Branded workbook downloads with fonts, fills, borders, number formats, merged titles, column widths, row heights, totals, and frozen panes.",
		link: "/guide/styled-workbooks",
		linkText: "Styled workbooks",
	},
	{
		id: "R02",
		title: "Migrate SheetJS code, keep your data model",
		body: "Cell objects keep the familiar shape. Add await, switch to named imports, move from sheet_to_json to sheetToJson. Done.",
		link: "/guide/migration",
		linkText: "Migration guide",
	},
	{
		id: "R03",
		title: "Safer exports when data comes from users",
		body: "CSV formula-like fields are escaped by default. HTML export sanitizes unsafe link targets unless you explicitly opt out.",
		link: "/guide/security",
		linkText: "Security guidance",
	},
	{
		id: "R04",
		title: "Convert spreadsheet data both ways",
		body: "Move between worksheets, JSON, arrays of arrays, CSV, TSV, and HTML tables with typed utilities built for modern runtimes.",
		link: "/guide/getting-started",
		linkText: "Getting started",
	},
];

const runtimes = [
	{
		tag: "Browser",
		title: "Ships to the client",
		body: "File inputs, fetch responses, and Blob downloads. No separate browser bundle.",
	},
	{
		tag: "Server",
		title: "Wraps your own I/O",
		body: "Async read and write around Node's fs, S3, or any byte source you already use.",
	},
	{
		tag: "Edge",
		title: "Runs on the edge",
		body: "No node:fs import and no dependency chain, so Workers and Deno Deploy just work.",
	},
];

const capabilityGroups = [
	{ title: "Workbook core", items: ["Multiple sheets", "Defined names", "Properties", "Date systems"] },
	{ title: "Cell data", items: ["Strings", "Numbers", "Dates", "Booleans", "Errors"] },
	{ title: "Report styling", items: ["Fonts", "Fills", "Borders", "Alignment", "Number formats"] },
	{ title: "Sheet structure", items: ["Merged cells", "Column widths", "Row heights", "Frozen panes", "Auto filters"] },
	{ title: "Formulas & notes", items: ["Cell formulas", "Array formulas", "Shared formulas", "Threaded comments"] },
	{ title: "Export guards", items: ["CSV escaping", "Link sanitizing", "Error codes", "ZIP & XML limits"] },
];

const codeLines: Array<Array<{ t: string; c?: string }>> = [
	[{ t: "import" }, { t: " { arrayToSheet, createWorkbook," }],
	[{ t: "         styleRange, write }" }, { t: " from " }, { t: '"xlsx-format"', c: "str" }],
	[],
	[{ t: "const" }, { t: " sheet " }, { t: "=", c: "op" }, { t: " arrayToSheet(reportRows)", c: "fn" }],
	[{ t: "styleRange(sheet, " }, { t: '"A2:D2"', c: "str" }, { t: ", headerStyle)" }],
	[],
	[{ t: "const" }, { t: " bytes " }, { t: "=", c: "op" }, { t: " await", c: "kw" }, { t: " write(book, {" }],
	[{ t: "  type: " }, { t: '"array"', c: "str" }, { t: ", cellStyles: " }, { t: "true", c: "kw" }, { t: " })" }],
];

const sheetRows = [
	{ cells: ["Apr 2026", "12,400", "12,050", "-18,350"], negative: [3] },
	{ cells: ["May 2026", "13,100", "13,980", "44,200"], selected: true },
	{ cells: ["Jun 2026", "14,200", "14,460", "13,520"] },
];

/* ------------------------------------------------------------------ */
/*  Page                                                              */
/* ------------------------------------------------------------------ */

export default function HomePage() {
	const version = config.project?.version ?? "0.0.0";

	// Scroll reveal. Content is visible by default; we only hide reveal targets
	// once JS is active (the .xf-js class), so SSR and no-JS render everything
	// and nothing ships blank. Reduced motion skips the whole mechanism.
	useEffect(() => {
		if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
		const root = document.querySelector<HTMLElement>(".xf");
		const targets = document.querySelectorAll<HTMLElement>("[data-reveal]");
		if (!root || !targets.length) return;
		root.classList.add("xf-js");
		const io = new IntersectionObserver(
			(entries, obs) => {
				for (const entry of entries) {
					if (entry.isIntersecting) {
						entry.target.classList.add("is-in");
						obs.unobserve(entry.target);
					}
				}
			},
			{ rootMargin: "0px 0px -10% 0px", threshold: 0.1 },
		);
		for (const el of targets) io.observe(el);
		return () => {
			io.disconnect();
			root.classList.remove("xf-js");
		};
	}, []);

	return (
		<div className="xf">
			{/* ============================ HERO ============================ */}
			<header className="xf-hero">
				<div className="xf-hero__grid" aria-hidden="true" />

				<div className="xf-hero__copy">
					<p className="xf-formula">
						<span className="xf-formula__fx">fx</span>
						<span className="xf-formula__val">
							zero-dependency XLSX <span className="xf-formula__sep">·</span> v{version}
						</span>
					</p>

					<h1 className="xf-hero__title">
						<span className="xf-hero__line">Real Excel files.</span>
						<span className="xf-hero__line xf-hero__line--accent">None of the weight.</span>
					</h1>

					<p className="xf-hero__lead">
						A modern XLSX reader and writer for TypeScript. Fully async, tree-shakeable, and ready for the
						browser, Node, and the edge. Read, write, convert, and style workbooks without carrying a
						spreadsheet framework.
					</p>

					<div className="xf-hero__actions">
						<Link className="xf-btn xf-btn--primary" to="/guide/getting-started">
							Build your first workbook
							<ArrowRight size={17} aria-hidden="true" />
						</Link>
						<Link className="xf-btn xf-btn--ghost" to="/guide/migration">
							Compare the migration
							<ArrowUpRight size={16} aria-hidden="true" />
						</Link>
						<a
							className="xf-btn xf-btn--bare"
							href="https://github.com/sebastian-software/xlsx-format"
							aria-label="View xlsx-format on GitHub"
						>
							<GithubMark size={16} />
							GitHub
						</a>
					</div>

					<dl className="xf-proof">
						{proofItems.map((item) => (
							<div key={item.label} className="xf-proof__cell">
								<dt>{item.label}</dt>
								<dd>{item.value}</dd>
							</div>
						))}
					</dl>
				</div>

				{/* The money shot: code in, styled workbook out */}
				<div className="xf-shot">
					<figure className="xf-editor">
						<figcaption className="xf-editor__bar">
							<span className="xf-dot" aria-hidden="true" />
							<span className="xf-editor__name">report-builder.ts</span>
							<span className="xf-editor__tag">async</span>
						</figcaption>
						<pre className="xf-code">
							<code>
								{codeLines.map((line, i) => (
									<span className="xf-code__row" key={i}>
										<span className="xf-code__ln" aria-hidden="true">
											{i + 1}
										</span>
										<span className="xf-code__src">
											{line.length === 0 ? (
												" "
											) : (
												line.map((tok, j) => (
													<span key={j} className={tok.c ? `tok tok--${tok.c}` : undefined}>
														{tok.t}
													</span>
												))
											)}
										</span>
									</span>
								))}
							</code>
						</pre>
					</figure>

					<div className="xf-pipe" aria-hidden="true">
						<span className="xf-pipe__line" />
						<span className="xf-pipe__chip">write()</span>
						<span className="xf-pipe__line" />
					</div>

					<figure className="xf-book" aria-label="Styled workbook output, Q2 solar report">
						<figcaption className="xf-book__bar">
							<span className="xf-book__file">Q2-report.xlsx</span>
							<span className="xf-book__sheet">Overview</span>
						</figcaption>
						<div className="xf-table" role="presentation">
							<div className="xf-table__cols" aria-hidden="true">
								<span className="xf-table__corner" />
								<span>A</span>
								<span>B</span>
								<span>C</span>
								<span>D</span>
							</div>
							<div className="xf-table__title">
								<span className="xf-table__rownum" aria-hidden="true">
									1
								</span>
								<span className="xf-table__titletext">Northstar Solar PPA · Q2 Report</span>
							</div>
							<div className="xf-table__row xf-table__row--head">
								<span className="xf-table__rownum" aria-hidden="true">
									2
								</span>
								<span>Month</span>
								<span>Plan</span>
								<span>Actual</span>
								<span>Settle</span>
							</div>
							{sheetRows.map((row, ri) => (
								<div
									className={`xf-table__row${row.selected ? " is-selected" : ""}`}
									key={row.cells[0]}
								>
									<span className="xf-table__rownum" aria-hidden="true">
										{ri + 3}
									</span>
									{row.cells.map((cell, ci) => (
										<span
											key={ci}
											className={row.negative?.includes(ci) ? "is-neg" : undefined}
										>
											{cell}
										</span>
									))}
								</div>
							))}
							<div className="xf-table__row xf-table__row--total">
								<span className="xf-table__rownum" aria-hidden="true">
									6
								</span>
								<span>Q2 Total</span>
								<span>39,700</span>
								<span>40,490</span>
								<span>39,370</span>
							</div>
						</div>
					</figure>
				</div>
			</header>

			{/* ====================== WHY / COMPARISON ====================== */}
			<section className="xf-section xf-why" aria-labelledby="why-title">
				<div className="xf-section__head" data-reveal>
					<span className="xf-label">A · Why switch</span>
					<h2 id="why-title">
						Most apps don&rsquo;t need a <em>spreadsheet framework</em>.
					</h2>
					<p>
						The popular libraries ship support for dozens of legacy formats and pull in 7 to 9 runtime
						dependencies. If the job is XLSX, CSV, and polished report exports, xlsx-format keeps the useful
						parts and leaves the rest out of your bundle.
					</p>
				</div>

				<div className="xf-matrix" data-reveal role="table" aria-label="Library comparison">
					<div className="xf-matrix__head" role="row">
						<span role="columnheader">Decision point</span>
						{comparisonColumns.map((col, i) => (
							<span key={col} role="columnheader" className={i === 0 ? "is-primary" : undefined}>
								{col}
							</span>
						))}
					</div>
					{comparisonRows.map((row) => (
						<div className="xf-matrix__row" role="row" key={row.label}>
							<span role="rowheader">{row.label}</span>
							{row.values.map((value, i) => (
								<span
									key={i}
									role="cell"
									className={i === row.win ? "is-win" : "is-rival"}
								>
									{i === row.win ? (
										<Check size={14} aria-hidden="true" strokeWidth={3} />
									) : (
										<Minus size={13} aria-hidden="true" className="xf-matrix__minus" />
									)}
									{value}
								</span>
							))}
						</div>
					))}
				</div>
			</section>

			{/* ========================= USE CASES ========================= */}
			<section className="xf-section xf-cases" aria-labelledby="cases-title">
				<div className="xf-section__head xf-section__head--wide" data-reveal>
					<span className="xf-label">B · Built for export work</span>
					<h2 id="cases-title">One lean layer for the places spreadsheets actually show up.</h2>
				</div>

				<ol className="xf-cases__list">
					{useCases.map((useCase) => (
						<li className="xf-case" key={useCase.id} data-reveal>
							<span className="xf-case__id" aria-hidden="true">
								{useCase.id}
							</span>
							<div className="xf-case__body">
								<h3>{useCase.title}</h3>
								<p>{useCase.body}</p>
							</div>
							<Link className="xf-case__link" to={useCase.link}>
								{useCase.linkText}
								<ArrowRight size={15} aria-hidden="true" />
							</Link>
						</li>
					))}
				</ol>
			</section>

			{/* ========================== RUNTIME ========================== */}
			<section className="xf-section xf-runtime" aria-labelledby="runtime-title">
				<div className="xf-section__head" data-reveal>
					<span className="xf-label">C · Runs where your app runs</span>
					<h2 id="runtime-title">No Node-only core. No browser-only fork.</h2>
					<p>
						xlsx-format reads and writes in-memory sources: Uint8Array, ArrayBuffer, Node Buffer, base64,
						binary strings, CSV text, and HTML tables. Pair it with whatever I/O you already have.
					</p>
				</div>

				<div className="xf-runtime__cards">
					{runtimes.map((rt) => (
						<article className="xf-rt" key={rt.tag} data-reveal>
							<span className="xf-rt__tag">{rt.tag}</span>
							<h3>{rt.title}</h3>
							<p>{rt.body}</p>
						</article>
					))}
				</div>
			</section>

			{/* ======================== CAPABILITIES ======================= */}
			<section className="xf-section xf-caps" aria-labelledby="caps-title">
				<div className="xf-section__head xf-section__head--wide" data-reveal>
					<span className="xf-label">D · Feature surface</span>
					<h2 id="caps-title">Enough spreadsheet power for production reports, kept deliberately focused.</h2>
				</div>

				<div className="xf-caps__grid">
					{capabilityGroups.map((group) => (
						<article className="xf-cap" key={group.title} data-reveal>
							<h3>{group.title}</h3>
							<ul>
								{group.items.map((item) => (
									<li key={item}>{item}</li>
								))}
							</ul>
						</article>
					))}
				</div>
			</section>

			{/* ========================= FINAL CTA ========================= */}
			<section className="xf-cta-band" aria-labelledby="cta-title" data-reveal>
				<div className="xf-cta-band__grid" aria-hidden="true" />
				<div className="xf-cta-band__inner">
					<div>
						<span className="xf-label xf-label--invert">Install the smaller layer</span>
						<h2 id="cta-title">Start with async XLSX today.</h2>
						<p>Keep the workbook framework out of your bundle. Your users download less, your build stays fast.</p>
					</div>
					<div className="xf-cta-band__action">
						<div className="xf-install">
							<span className="xf-install__prompt" aria-hidden="true">
								$
							</span>
							<code>npm install xlsx-format</code>
						</div>
						<Link className="xf-btn xf-btn--accent" to="/guide/getting-started">
							Create your first workbook
							<Download size={17} aria-hidden="true" />
						</Link>
					</div>
				</div>
			</section>
		</div>
	);
}
