/**
 * Convert a JavaScript Date to an Excel serial date number.
 *
 * Excel serial dates count days since 1899-12-30 (the epoch).
 * The 1900 date system intentionally includes a fictitious Feb 29, 1900
 * (serial number 60) to maintain compatibility with Lotus 1-2-3.
 * For serial numbers >= 60, we add 1 to skip over this phantom day.
 *
 * @param v - JavaScript Date to convert
 * @param date1904 - If true, use the 1904 date system (Mac Excel default), which shifts the epoch by 1462 days
 * @returns Excel serial date number
 */
export function dateToSerialNumber(v: Date, date1904?: boolean): number {
	let epoch = v.getTime();
	if (date1904) {
		// 1904 date system: subtract 1462 days (the offset between 1900-01-01 and 1904-01-01)
		epoch -= 1462 * 24 * 60 * 60 * 1000;
	}
	// Excel epoch: 1899-12-30T00:00:00Z (Dec 30, 1899)
	const dnthresh = Date.UTC(1899, 11, 30, 0, 0, 0);
	const result = (epoch - dnthresh) / (24 * 60 * 60 * 1000);
	// Excel intentionally considers 1900-02-29 a valid date (serial 60) — Lotus 1-2-3 bug.
	// For serial numbers < 60, no adjustment is needed.
	if (result < 60) {
		return result;
	}
	// For serial numbers >= 60, add 1 to account for the phantom Feb 29, 1900
	if (result >= 60) {
		return result + 1;
	}
	return result;
}

/**
 * Convert an Excel serial date number to a JavaScript Date.
 *
 * Reverses the conversion done by {@link dateToSerialNumber}, accounting for
 * the 1904 date system offset and the Lotus 1-2-3 leap year bug at serial 60.
 *
 * @param v - Excel serial date number
 * @param date1904 - If true, use the 1904 date system (adds 1462 days)
 * @returns JavaScript Date corresponding to the serial number
 */
export function serialNumberToDate(v: number, date1904?: boolean): Date {
	let date = v;
	if (date1904) {
		// 1904 system: add back the 1462-day offset
		date += 1462;
	}
	// Lotus 1-2-3 bug: serial 60 is the fictitious 1900-02-29.
	// For serial > 60, subtract 1 to compensate for the phantom day.
	if (date > 60) {
		--date;
	}
	// Excel epoch: 1899-12-30T00:00:00Z
	const dnthresh = Date.UTC(1899, 11, 30, 0, 0, 0);
	return new Date(dnthresh + date * 24 * 60 * 60 * 1000);
}

/**
 * Shift a local Date to UTC by adding the timezone offset.
 *
 * Useful when a Date was constructed from local-time components but
 * needs to be treated as a UTC timestamp.
 *
 * @param d - Date in local time
 * @returns New Date shifted to represent the same wall-clock time in UTC
 */
export function localToUtc(d: Date): Date {
	const off = d.getTimezoneOffset();
	return new Date(d.getTime() + off * 60 * 1000);
}

/**
 * Shift a UTC Date to local time by subtracting the timezone offset.
 *
 * The inverse of {@link localToUtc}.
 *
 * @param d - Date in UTC
 * @returns New Date shifted to represent the same wall-clock time in local time
 */
export function utcToLocal(d: Date): Date {
	const off = d.getTimezoneOffset();
	return new Date(d.getTime() - off * 60 * 1000);
}
