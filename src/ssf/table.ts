/**
 * Initialize the built-in Excel number format table.
 *
 * Populates the standard format IDs (0-49 plus 56) defined in ECMA-376.
 * Format IDs 5-8 and 23-36 are locale-dependent and intentionally omitted
 * from the base table (they are mapped via {@link DEFAULT_FORMAT_MAP}).
 *
 * @param t - Optional existing table to populate; a new object is created if omitted
 * @returns The populated format table
 */
export function initFormatTable(t?: Record<number, string>): Record<number, string> {
	if (!t) {
		t = {};
	}
	t[0] = "General";
	t[1] = "0";
	t[2] = "0.00";
	t[3] = "#,##0";
	t[4] = "#,##0.00";
	t[9] = "0%";
	t[10] = "0.00%";
	t[11] = "0.00E+00";
	t[12] = "# ?/?";
	t[13] = "# ??/??";
	t[14] = "m/d/yy";
	t[15] = "d-mmm-yy";
	t[16] = "d-mmm";
	t[17] = "mmm-yy";
	t[18] = "h:mm AM/PM";
	t[19] = "h:mm:ss AM/PM";
	t[20] = "h:mm";
	t[21] = "h:mm:ss";
	t[22] = "m/d/yy h:mm";
	t[37] = "#,##0 ;(#,##0)";
	t[38] = "#,##0 ;[Red](#,##0)";
	t[39] = "#,##0.00;(#,##0.00)";
	t[40] = "#,##0.00;[Red](#,##0.00)";
	t[45] = "mm:ss";
	t[46] = "[h]:mm:ss";
	t[47] = "mmss.0";
	t[48] = "##0.0E+0";
	t[49] = "@";
	// Format 56: Chinese time format (Simplified Chinese AM/PM with hour/minute/second)
	t[56] = '"上午/下午 "hh"時"mm"分"ss"秒 "';
	return t;
}

/** Default number format table, initialized with standard Excel formats */
export let formatTable: Record<number, string> = initFormatTable();

/**
 * Mapping from locale-dependent format IDs to their base-table equivalents.
 *
 * When a format ID is not found in the main table, this map provides a fallback
 * by redirecting to a standard format ID. For example, format 27 (a locale-specific
 * date format) falls back to format 14 ("m/d/yy").
 *
 * Defaults were determined by systematically testing in Excel 2019.
 */
export const DEFAULT_FORMAT_MAP: Record<number, number> = {
	// Accounting formats 5-8 map to the parenthesized number formats 37-40
	5: 37,
	6: 38,
	7: 39,
	8: 40,
	// CJK locale formats 23-26 map to General (0)
	23: 0,
	24: 0,
	25: 0,
	26: 0,
	// CJK/locale date formats 27-31, 50-58 map to short date (14)
	27: 14,
	28: 14,
	29: 14,
	30: 14,
	31: 14,
	50: 14,
	51: 14,
	52: 14,
	53: 14,
	54: 14,
	55: 14,
	56: 14,
	57: 14,
	58: 14,
	// CJK number formats
	59: 1,
	60: 2,
	61: 3,
	62: 4,
	// CJK percentage/fraction formats
	67: 9,
	68: 10,
	69: 12,
	70: 13,
	// CJK date/time formats
	71: 14,
	72: 14,
	73: 15,
	74: 16,
	75: 17,
	76: 20,
	77: 21,
	78: 22,
	79: 45,
	80: 46,
	81: 47,
	82: 0,
};

/**
 * Accounting format strings for IDs that have no direct equivalent in the base table.
 *
 * These use literal "$" currency symbols and special alignment characters
 * (_  for padding, \( and \) for literal parentheses).
 */
export const DEFAULT_FORMAT_STRINGS: Record<number, string> = {
	5: '"$"#,##0_);\\("$"#,##0\\)',
	63: '"$"#,##0_);\\("$"#,##0\\)',
	6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
	64: '"$"#,##0_);[Red]\\("$"#,##0\\)',
	7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
	65: '"$"#,##0.00_);\\("$"#,##0.00\\)',
	8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	66: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
	42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
	43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
	44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
};

/**
 * Register a custom number format string in the format table.
 *
 * If no index is provided, searches for an existing match or the first empty slot
 * in the valid range (0-0x0187). If no slot is found, uses 0x0187 as a last resort.
 *
 * @param fmt - The format string to register (e.g. "#,##0.00")
 * @param idx - Optional explicit format index to assign
 * @returns The format index where the string was registered
 */
export function loadFormat(fmt: string, idx?: number): number {
	if (typeof idx !== "number") {
		idx = Number(idx) || -1;
		// 0x0188 = 392, the maximum number of format entries
		for (let i = 0; i < 0x0188; ++i) {
			if (formatTable[i] === undefined) {
				if (idx < 0) {
					idx = i;
				}
				continue;
			}
			if (formatTable[i] === fmt) {
				idx = i;
				break;
			}
		}
		if (idx < 0) {
			idx = 0x187;
		}
	}
	formatTable[idx] = fmt;
	return idx;
}

/**
 * Bulk-load a table of number format strings, overwriting existing entries.
 * @param tbl - Map of format index to format string
 */
export function loadFormatTable(tbl: Record<number, string>): void {
	for (let i = 0; i < 0x0188; ++i) {
		if (tbl[i] !== undefined) {
			loadFormat(tbl[i], i);
		}
	}
}

/**
 * Reset the format table back to the built-in Excel defaults.
 *
 * Called before read/write operations to ensure a clean state.
 */
export function resetFormatTable(): void {
	formatTable = initFormatTable();
}
