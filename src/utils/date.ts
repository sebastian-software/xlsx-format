/**
 * Convert a JavaScript Date to an Excel serial date number.
 *
 * Excel serial dates in the 1900 date system can be mapped to real JavaScript
 * dates with the 1899-12-30 epoch. This preserves modern Excel serials such as
 * 45292 -> 2024-01-01 while avoiding a representational value for Excel's
 * fictitious 1900-02-29.
 *
 * @param v - JavaScript Date to convert
 * @param date1904 - If true, use the 1904 date system (Mac Excel default), which shifts the epoch by 1462 days
 * @returns Excel serial date number
 */
export function dateToSerialNumber(v: Date, date1904?: boolean): number {
	const epoch = v.getTime();
	if (date1904) {
		const date1904Epoch = Date.UTC(1904, 0, 1, 0, 0, 0);
		return (epoch - date1904Epoch) / (24 * 60 * 60 * 1000);
	}
	// Excel epoch: 1899-12-30T00:00:00Z (Dec 30, 1899)
	const dnthresh = Date.UTC(1899, 11, 30, 0, 0, 0);
	return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
}

/**
 * Convert an Excel serial date number to a JavaScript Date.
 *
 * Reverses the conversion done by {@link dateToSerialNumber}. The SSF display
 * formatter represents Excel's fictitious serial 60 separately; this helper
 * returns real JavaScript Dates for machine-readable conversion paths.
 *
 * @param v - Excel serial date number
 * @param date1904 - If true, use the 1904 date system (adds 1462 days)
 * @returns JavaScript Date corresponding to the serial number
 */
export function serialNumberToDate(v: number, date1904?: boolean): Date {
	if (date1904) {
		const date1904Epoch = Date.UTC(1904, 0, 1, 0, 0, 0);
		return new Date(date1904Epoch + v * 24 * 60 * 60 * 1000);
	}
	// Excel epoch: 1899-12-30T00:00:00Z
	const dnthresh = Date.UTC(1899, 11, 30, 0, 0, 0);
	return new Date(dnthresh + v * 24 * 60 * 60 * 1000);
}

/**
 * Copy local calendar components into a UTC Date for Excel's timezone-free serials.
 *
 * @param d - Date whose local wall-clock time should be preserved
 * @returns New Date with the same components in UTC
 */
export function localToUtc(d: Date): Date {
	const result = new Date(0);
	result.setUTCFullYear(d.getFullYear(), d.getMonth(), d.getDate());
	result.setUTCHours(d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds());
	return result;
}

/**
 * Copy UTC calendar components into a local Date.
 *
 * JavaScript resolves nonexistent or repeated local times using its usual
 * compatible DST disambiguation, just as the local Date constructor does.
 *
 * @param d - Date carrying timezone-free worksheet components in UTC
 * @returns New Date with those components in local time
 */
export function utcToLocal(d: Date): Date {
	const result = new Date(0);
	result.setFullYear(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
	result.setHours(d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds(), d.getUTCMilliseconds());
	return result;
}
