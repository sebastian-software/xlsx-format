/** Built-in Excel number format table */
export function SSF_init_table(t?: Record<number, string>): Record<number, string> {
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
	t[56] = '"上午/下午 "hh"時"mm"分"ss"秒 "';
	return t;
}

/** Default number format table */
export let table_fmt: Record<number, string> = SSF_init_table();

/** Defaults determined by systematically testing in Excel 2019 */
export const SSF_default_map: Record<number, number> = {
	5: 37,
	6: 38,
	7: 39,
	8: 40,
	23: 0,
	24: 0,
	25: 0,
	26: 0,
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
	59: 1,
	60: 2,
	61: 3,
	62: 4,
	67: 9,
	68: 10,
	69: 12,
	70: 13,
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

/** Accounting formats with no equivalent in the base table */
export const SSF_default_str: Record<number, string> = {
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

export function SSF_load(fmt: string, idx?: number): number {
	if (typeof idx !== "number") {
		idx = Number(idx) || -1;
		for (let i = 0; i < 0x0188; ++i) {
			if (table_fmt[i] === undefined) {
				if (idx < 0) {
					idx = i;
				}
				continue;
			}
			if (table_fmt[i] === fmt) {
				idx = i;
				break;
			}
		}
		if (idx < 0) {
			idx = 0x187;
		}
	}
	table_fmt[idx] = fmt;
	return idx;
}

export function SSF_load_table(tbl: Record<number, string>): void {
	for (let i = 0; i < 0x0188; ++i) {
		if (tbl[i] !== undefined) {
			SSF_load(tbl[i], i);
		}
	}
}

export function make_ssf(): void {
	table_fmt = SSF_init_table();
}
