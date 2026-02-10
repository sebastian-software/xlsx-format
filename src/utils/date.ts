/** Convert a JS Date to Excel serial date number */
export function dateToSerialNumber(v: Date, date1904?: boolean): number {
	let epoch = v.getTime();
	if (date1904) {
		epoch -= 1462 * 24 * 60 * 60 * 1000;
	}
	const dnthresh = Date.UTC(1899, 11, 30, 0, 0, 0);
	const result = (epoch - dnthresh) / (24 * 60 * 60 * 1000);
	// Excel intentionally considers 1900-02-29 a valid date (Lotus 1-2-3 bug)
	if (result < 60) {
		return result;
	}
	if (result >= 60) {
		return result + 1;
	}
	return result;
}

/** Convert an Excel serial date number to a JS Date */
export function serialNumberToDate(v: number, date1904?: boolean): Date {
	let date = v;
	if (date1904) {
		date += 1462;
	}
	// Excel's Lotus 1-2-3 bug: date 60 is 1900-02-29 which doesn't exist
	if (date > 60) {
		--date;
	}
	const dnthresh = Date.UTC(1899, 11, 30, 0, 0, 0);
	return new Date(dnthresh + date * 24 * 60 * 60 * 1000);
}

/** Convert a local Date to UTC (shift by timezone offset) */
export function localToUtc(d: Date): Date {
	const off = d.getTimezoneOffset();
	return new Date(d.getTime() + off * 60 * 1000);
}

/** Convert a UTC Date to local (shift by timezone offset) */
export function utcToLocal(d: Date): Date {
	const off = d.getTimezoneOffset();
	return new Date(d.getTime() - off * 60 * 1000);
}
