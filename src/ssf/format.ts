/* ssf.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* Ported to TypeScript for xlsx-format. 1:1 faithful port. */

import { dateToSerialNumber } from "../utils/date.js";
import { formatTable, DEFAULT_FORMAT_MAP, DEFAULT_FORMAT_STRINGS } from "./table.js";

const reverseString = (x: string): string => x.split("").reverse().join("");
const padWithZeros = (value: any, width: number): string => ("" + value).padStart(width, "0");
const padWithSpaces = (value: any, width: number): string => ("" + value).padStart(width, " ");
const rightPadWithSpaces = (value: any, width: number): string => ("" + value).padEnd(width, " ");
const padRoundedZeros = (value: any, width: number): string => ("" + Math.round(value)).padStart(width, "0");

function isGeneralFormat(s: string, i?: number): boolean {
	i = i || 0;
	return (
		s.length >= 7 + i &&
		(s.charCodeAt(i) | 32) === 103 &&
		(s.charCodeAt(i + 1) | 32) === 101 &&
		(s.charCodeAt(i + 2) | 32) === 110 &&
		(s.charCodeAt(i + 3) | 32) === 101 &&
		(s.charCodeAt(i + 4) | 32) === 114 &&
		(s.charCodeAt(i + 5) | 32) === 97 &&
		(s.charCodeAt(i + 6) | 32) === 108
	);
}

const days: string[][] = Array.from({ length: 7 }, (_, i) => {
	const d = new Date(2017, 0, i + 1); // 2017-01-01 is a Sunday
	return [
		new Intl.DateTimeFormat("en-US", { weekday: "short" }).format(d),
		new Intl.DateTimeFormat("en-US", { weekday: "long" }).format(d),
	];
});
const months: string[][] = Array.from({ length: 12 }, (_, i) => {
	const d = new Date(2000, i, 1);
	const short = new Intl.DateTimeFormat("en-US", { month: "short" }).format(d);
	const long = new Intl.DateTimeFormat("en-US", { month: "long" }).format(d);
	return [short[0], short, long];
});

interface SSFDateVal {
	daySerial: number;
	timeSeconds: number;
	subSeconds: number;
	year: number;
	month: number;
	day: number;
	hours: number;
	minutes: number;
	seconds: number;
	dayOfWeek: number;
}

function normalizeExcelNumber(value: number): number {
	const precStr = value.toPrecision(16);
	if (precStr.indexOf("e") > -1) {
		const mantissa = precStr.slice(0, precStr.indexOf("e"));
		const ml =
			mantissa.indexOf(".") > -1
				? mantissa.slice(0, mantissa.slice(0, 2) === "0." ? 17 : 16)
				: mantissa.slice(0, 15) + "0".repeat(mantissa.length - 15);
		return +ml + +("1" + precStr.slice(precStr.indexOf("e"))) - 1 || +precStr;
	}
	const normalizedStr =
		precStr.indexOf(".") > -1 ? precStr.slice(0, precStr.slice(0, 2) === "0." ? 17 : 16) : precStr.slice(0, 15) + "0".repeat(precStr.length - 15);
	return Number(normalizedStr);
}

function SSF_fix_hijri(_date: Date, o: number[]): number {
	o[0] -= 581;
	const dow = _date.getDay();
	if (_date.getTime() < -2203891200000) {
		return (dow + 6) % 7;
	}
	return dow;
}

export function parseExcelDateCode(value: number, opts?: any, hijriMode?: boolean): SSFDateVal | null {
	if (value > 2958465 || value < 0) {
		return null;
	}
	value = normalizeExcelNumber(value);
	let date = value | 0;
	let time = Math.floor(86400 * (value - date));
	const out: SSFDateVal = {
		daySerial: date,
		timeSeconds: time,
		subSeconds: 86400 * (value - date) - time,
		year: 0,
		month: 0,
		day: 0,
		hours: 0,
		minutes: 0,
		seconds: 0,
		dayOfWeek: 0,
	};
	if (Math.abs(out.subSeconds) < 1e-6) {
		out.subSeconds = 0;
	}
	if (opts && opts.date1904) {
		date += 1462;
	}
	if (out.subSeconds > 0.9999) {
		out.subSeconds = 0;
		if (++time === 86400) {
			out.timeSeconds = time = 0;
			++date;
			++out.daySerial;
		}
	}
	let dout: number[];
	let dow = 0;
	if (date === 60) {
		dout = hijriMode ? [1317, 10, 29] : [1900, 2, 29];
		dow = 3;
	} else if (date === 0) {
		dout = hijriMode ? [1317, 8, 29] : [1900, 1, 0];
		dow = 6;
	} else {
		if (date > 60) {
			--date;
		}
		const baseDate = new Date(1900, 0, 1);
		baseDate.setDate(baseDate.getDate() + date - 1);
		dout = [baseDate.getFullYear(), baseDate.getMonth() + 1, baseDate.getDate()];
		dow = baseDate.getDay();
		if (date < 60) {
			dow = (dow + 6) % 7;
		}
		if (hijriMode) {
			dow = SSF_fix_hijri(baseDate, dout);
		}
	}
	out.year = dout[0];
	out.month = dout[1];
	out.day = dout[2];
	out.seconds = time % 60;
	time = Math.floor(time / 60);
	out.minutes = time % 60;
	time = Math.floor(time / 60);
	out.hours = time;
	out.dayOfWeek = dow;
	return out;
}

function SSF_strip_decimal(o: string): string {
	return o.indexOf(".") === -1 ? o : o.replace(/(?:\.0*|(\.\d*[1-9])0+)$/, "$1");
}

function SSF_normalize_exp(o: string): string {
	if (o.indexOf("E") === -1) {
		return o;
	}
	return o.replace(/(?:\.0*|(\.\d*[1-9])0+)[Ee]/, "$1E").replace(/(E[+-])(\d)$/, "$10$2");
}

function SSF_small_exp(v: number): string {
	const w = v < 0 ? 12 : 11;
	let o = SSF_strip_decimal(v.toFixed(12));
	if (o.length <= w) {
		return o;
	}
	o = v.toPrecision(10);
	if (o.length <= w) {
		return o;
	}
	return v.toExponential(5);
}

function SSF_large_exp(v: number): string {
	const o = SSF_strip_decimal(v.toFixed(11));
	return o.length > (v < 0 ? 12 : 11) || o === "0" || o === "-0" ? v.toPrecision(6) : o;
}

function SSF_general_num(v: number): string {
	if (!isFinite(v)) {
		return isNaN(v) ? "#NUM!" : "#DIV/0!";
	}
	const V = Math.floor(Math.log(Math.abs(v)) * Math.LOG10E);
	let o: string;

	if (V >= -4 && V <= -1) {
		o = v.toPrecision(10 + V);
	} else if (Math.abs(V) <= 9) {
		o = SSF_small_exp(v);
	} else if (V === 10) {
		o = v.toFixed(10).substring(0, 12);
	} else {
		o = SSF_large_exp(v);
	}

	return SSF_strip_decimal(SSF_normalize_exp(o.toUpperCase()));
}

function SSF_general(v: any, opts: any): string {
	switch (typeof v) {
		case "string":
			return v;
		case "boolean":
			return v ? "TRUE" : "FALSE";
		case "number":
			return (v | 0) === v ? v.toString(10) : SSF_general_num(v);
		case "undefined":
			return "";
		case "object":
			if (v == null) {
				return "";
			}
			if (v instanceof Date) {
				return formatNumber(14, dateToSerialNumber(v, opts && opts.date1904), opts);
			}
	}
	throw new Error("unsupported value in General format: " + v);
}

function SSF_write_date(type: number, fmt: string, val: SSFDateVal, ss0?: number): string {
	let result = "";
	let scaledSeconds = 0;
	let scaleFactor = 0;
	let year = val.year;
	let numericOut: number = 0;
	let outputLength = 0;
	switch (type) {
		case 98 /* 'b' buddhist year */:
			year = val.year + 543;
		/* falls through */
		case 121 /* 'y' year */:
			switch (fmt.length) {
				case 1:
				case 2:
					numericOut = year % 100;
					outputLength = 2;
					break;
				default:
					numericOut = year % 10000;
					outputLength = 4;
					break;
			}
			break;
		case 109 /* 'm' month */:
			switch (fmt.length) {
				case 1:
				case 2:
					numericOut = val.month;
					outputLength = fmt.length;
					break;
				case 3:
					return months[val.month - 1][1];
				case 5:
					return months[val.month - 1][0];
				default:
					return months[val.month - 1][2];
			}
			break;
		case 100 /* 'd' day */:
			switch (fmt.length) {
				case 1:
				case 2:
					numericOut = val.day;
					outputLength = fmt.length;
					break;
				case 3:
					return days[val.dayOfWeek][0];
				default:
					return days[val.dayOfWeek][1];
			}
			break;
		case 104 /* 'h' 12-hour */:
			switch (fmt.length) {
				case 1:
				case 2:
					numericOut = 1 + ((val.hours + 11) % 12);
					outputLength = fmt.length;
					break;
				default:
					throw new Error("bad hour format: " + fmt);
			}
			break;
		case 72 /* 'H' 24-hour */:
			switch (fmt.length) {
				case 1:
				case 2:
					numericOut = val.hours;
					outputLength = fmt.length;
					break;
				default:
					throw new Error("bad hour format: " + fmt);
			}
			break;
		case 77 /* 'M' minutes */:
			switch (fmt.length) {
				case 1:
				case 2:
					numericOut = val.minutes;
					outputLength = fmt.length;
					break;
				default:
					throw new Error("bad minute format: " + fmt);
			}
			break;
		case 115 /* 's' seconds */:
			if (fmt !== "s" && fmt !== "ss" && fmt !== ".0" && fmt !== ".00" && fmt !== ".000") {
				throw new Error("bad second format: " + fmt);
			}
			if (val.subSeconds === 0 && (fmt === "s" || fmt === "ss")) {
				return padWithZeros(val.seconds, fmt.length);
			}
			if (ss0! >= 2) {
				scaleFactor = ss0 === 3 ? 1000 : 100;
			} else {
				scaleFactor = ss0 === 1 ? 10 : 1;
			}
			scaledSeconds = Math.round(scaleFactor * (val.seconds + val.subSeconds));
			if (scaledSeconds >= 60 * scaleFactor) {
				scaledSeconds = 0;
			}
			if (fmt === "s") {
				return scaledSeconds === 0 ? "0" : "" + scaledSeconds / scaleFactor;
			}
			result = padWithZeros(scaledSeconds, 2 + ss0!);
			if (fmt === "ss") {
				return result.substring(0, 2);
			}
			return "." + result.substring(2, fmt.length - 1);
		case 90 /* 'Z' absolute time */:
			switch (fmt) {
				case "[h]":
				case "[hh]":
					numericOut = val.daySerial * 24 + val.hours;
					break;
				case "[m]":
				case "[mm]":
					numericOut = (val.daySerial * 24 + val.hours) * 60 + val.minutes;
					break;
				case "[s]":
				case "[ss]":
					numericOut = ((val.daySerial * 24 + val.hours) * 60 + val.minutes) * 60 + (ss0 === 0 ? Math.round(val.seconds + val.subSeconds) : val.seconds);
					break;
				default:
					throw new Error("bad abstime format: " + fmt);
			}
			outputLength = fmt.length === 3 ? 1 : 2;
			break;
		case 101 /* 'e' era */:
			numericOut = year;
			outputLength = 1;
			break;
	}
	return outputLength > 0 ? padWithZeros(numericOut, outputLength) : "";
}

function commaify(str: string): string {
	return str.replace(/\B(?=(\d{3})+$)/g, ",");
}

const pct1 = /%/g;
function write_num_pct(type: string, fmt: string, val: number): string {
	const sfmt = fmt.replace(pct1, "");
	const mul = fmt.length - sfmt.length;
	return write_num(type, sfmt, val * Math.pow(10, 2 * mul)) + "%".repeat(mul);
}
function write_num_cm(type: string, fmt: string, val: number): string {
	let idx = fmt.length - 1;
	while (fmt.charCodeAt(idx - 1) === 44) {
		--idx;
	}
	return write_num(type, fmt.substring(0, idx), val / Math.pow(10, 3 * (fmt.length - idx)));
}
function write_num_exp(fmt: string, val: number): string {
	let o: string;
	const idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
	if (fmt.match(/^#+0.0E\+0$/)) {
		if (val === 0) {
			return "0.0E+0";
		}
		if (val < 0) {
			return "-" + write_num_exp(fmt, -val);
		}
		const period = fmt.indexOf(".");
		const ee =
			Math.floor(Math.log(val) * Math.LOG10E) % period < 0
				? (Math.floor(Math.log(val) * Math.LOG10E) % period) + period
				: Math.floor(Math.log(val) * Math.LOG10E) % period;
		o = (val / Math.pow(10, ee)).toPrecision(idx + 1 + ((period + ee) % period));
		if (o.indexOf("e") === -1) {
			const fakee = Math.floor(Math.log(val) * Math.LOG10E);
			if (o.indexOf(".") === -1) {
				o = o.charAt(0) + "." + o.substring(1) + "E+" + (fakee - o.length + ee);
			} else {
				o += "E+" + (fakee - ee);
			}
			while (o.substring(0, 2) === "0.") {
				o = o.charAt(0) + o.substring(2, period) + "." + o.substring(2 + period);
				o = o.replace(/^0+([1-9])/, "$1").replace(/^0+\./, "0.");
			}
			o = o.replace(/\+-/, "-");
		}
		o = o.replace(
			/^([+-]?)(\d*)\.(\d*)[Ee]/,
			($$, $1, $2, $3) => $1 + $2 + $3.substring(0, (period + ee) % period) + "." + $3.substring(ee) + "E",
		);
	} else {
		o = val.toExponential(idx);
	}
	if (fmt.match(/E\+00$/) && o.match(/e[+-]\d$/)) {
		o = o.substring(0, o.length - 1) + "0" + o.charAt(o.length - 1);
	}
	if (fmt.match(/E-/) && o.match(/e\+/)) {
		o = o.replace(/e\+/, "e");
	}
	return o.replace("e", "E");
}

function SSF_frac(value: number, maxDenominator: number, mixed?: boolean): number[] {
	const sgn = value < 0 ? -1 : 1;
	let absValue = value * sgn;
	let prevPrevNumer = 0,
		prevNumer = 1,
		numerator = 0;
	let prevPrevDenom = 1,
		prevDenom = 0,
		denominator = 0;
	let intPart = Math.floor(absValue);
	while (prevDenom < maxDenominator) {
		intPart = Math.floor(absValue);
		numerator = intPart * prevNumer + prevPrevNumer;
		denominator = intPart * prevDenom + prevPrevDenom;
		if (absValue - intPart < 0.00000005) {
			break;
		}
		absValue = 1 / (absValue - intPart);
		prevPrevNumer = prevNumer;
		prevNumer = numerator;
		prevPrevDenom = prevDenom;
		prevDenom = denominator;
	}
	if (denominator > maxDenominator) {
		if (prevDenom > maxDenominator) {
			denominator = prevPrevDenom;
			numerator = prevPrevNumer;
		} else {
			denominator = prevDenom;
			numerator = prevNumer;
		}
	}
	if (!mixed) {
		return [0, sgn * numerator, denominator];
	}
	const wholePart = Math.floor((sgn * numerator) / denominator);
	return [wholePart, sgn * numerator - wholePart * denominator, denominator];
}

const frac1 = /# (\?+)( ?)\/( ?)(\d+)/;
function write_num_f1(r: string[], aval: number, sign: string): string {
	const den = parseInt(r[4], 10);
	const rr = Math.round(aval * den);
	const base = Math.floor(rr / den);
	const myn = rr - base * den;
	const myd = den;
	return (
		sign +
		(base === 0 ? "" : "" + base) +
		" " +
		(myn === 0
			? " ".repeat(r[1].length + 1 + r[4].length)
			: padWithSpaces(myn, r[1].length) + r[2] + "/" + r[3] + padWithZeros(myd, r[4].length))
	);
}
function write_num_f2(r: string[], aval: number, sign: string): string {
	return sign + (aval === 0 ? "" : "" + aval) + " ".repeat(r[1].length + 2 + r[4].length);
}

const dec1 = /^#*0*\.([0#]+)/;
const closeparen = /\)[^)]*[0#]/;
const phone = /\(###\) ###\\?-####/;

function hashq(str: string): string {
	let o = "";
	for (let i = 0; i !== str.length; ++i) {
		const cc = str.charCodeAt(i);
		switch (cc) {
			case 35:
				break;
			case 63:
				o += " ";
				break;
			case 48:
				o += "0";
				break;
			default:
				o += String.fromCharCode(cc);
		}
	}
	return o;
}

function rnd(val: number, d: number): string {
	const sgn = val < 0 ? -1 : 1;
	const dd = Math.pow(10, d);
	return "" + sgn * (Math.round(sgn * val * dd) / dd);
}
function dec(val: number, d: number): number {
	const _frac = val - Math.floor(val);
	const dd = Math.pow(10, d);
	if (d < ("" + Math.round(_frac * dd)).length) {
		return 0;
	}
	return Math.round(_frac * dd);
}
function carry(val: number, d: number): number {
	if (d < ("" + Math.round((val - Math.floor(val)) * Math.pow(10, d))).length) {
		return 1;
	}
	return 0;
}
function flr(val: number): string {
	if (val < 2147483647 && val > -2147483648) {
		return "" + (val >= 0 ? val | 0 : (val - 1) | 0);
	}
	return "" + Math.floor(val);
}

function write_num_flt(type: string, fmt: string, val: number): string {
	if (type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
		const ffmt = fmt.replace(/\( */, "").replace(/ \)/, "").replace(/\)/, "");
		if (val >= 0) {
			return write_num_flt("n", ffmt, val);
		}
		return "(" + write_num_flt("n", ffmt, -val) + ")";
	}
	if (fmt.charCodeAt(fmt.length - 1) === 44) {
		return write_num_cm(type, fmt, val);
	}
	if (fmt.indexOf("%") !== -1) {
		return write_num_pct(type, fmt, val);
	}
	if (fmt.indexOf("E") !== -1) {
		return write_num_exp(fmt, val);
	}
	if (fmt.charCodeAt(0) === 36) {
		return "$" + write_num_flt(type, fmt.substring(fmt.charAt(1) === " " ? 2 : 1), val);
	}
	let o: string;
	let r: RegExpMatchArray | null;
	let ri: number;
	let ff: number[];
	const aval = Math.abs(val);
	const sign = val < 0 ? "-" : "";
	if (fmt.match(/^00+$/)) {
		return sign + padRoundedZeros(aval, fmt.length);
	}
	if (fmt.match(/^[#?]+$/)) {
		o = padRoundedZeros(val, 0);
		if (o === "0") {
			o = "";
		}
		return o.length > fmt.length ? o : hashq(fmt.substring(0, fmt.length - o.length)) + o;
	}
	if ((r = fmt.match(frac1))) {
		return write_num_f1(r, aval, sign);
	}
	if (fmt.match(/^#+0+$/)) {
		return sign + padRoundedZeros(aval, fmt.length - fmt.indexOf("0"));
	}
	if ((r = fmt.match(dec1))) {
		o = rnd(val, r[1].length)
			.replace(/^([^.]+)$/, "$1." + hashq(r[1]))
			.replace(/\.$/, "." + hashq(r[1]))
			.replace(/\.(\d*)$/, ($$, $1) => "." + $1 + "0".repeat(hashq(r![1]).length - $1.length));
		return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./, ".");
	}
	fmt = fmt.replace(/^#+([0.])/, "$1");
	if ((r = fmt.match(/^(0*)\.(#*)$/))) {
		return (
			sign +
			rnd(aval, r[2].length)
				.replace(/\.(\d*[1-9])0*$/, ".$1")
				.replace(/^(-?\d*)$/, "$1.")
				.replace(/^0\./, r[1].length ? "0." : ".")
		);
	}
	if (fmt.match(/^#{1,3},##0(\.?)$/)) {
		return sign + commaify(padRoundedZeros(aval, 0));
	}
	if ((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
		return val < 0
			? "-" + write_num_flt(type, fmt, -val)
			: commaify("" + (Math.floor(val) + carry(val, r[1].length))) +
					"." +
					padWithZeros(dec(val, r[1].length), r[1].length);
	}
	if ((r = fmt.match(/^#,#*,#0/))) {
		return write_num_flt(type, fmt.replace(/^#,#*,/, ""), val);
	}
	if ((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
		o = reverseString(write_num_flt(type, fmt.replace(/[\\-]/g, ""), val));
		ri = 0;
		return reverseString(
			reverseString(fmt.replace(/\\/g, "")).replace(/[0#]/g, (x) => {
				return ri < o.length ? o.charAt(ri++) : x === "0" ? "0" : "";
			}),
		);
	}
	if (fmt.match(phone)) {
		o = write_num_flt(type, "##########", val);
		return "(" + o.substring(0, 3) + ") " + o.substring(3, 3) + "-" + o.substring(6);
	}
	let oa = "";
	if ((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(r[4].length, 7);
		ff = SSF_frac(aval, Math.pow(10, ri) - 1, false);
		o = sign;
		oa = write_num("n", r[1], ff[1]);
		if (oa.charAt(oa.length - 1) === " ") {
			oa = oa.substring(0, oa.length - 1) + "0";
		}
		o += oa + r[2] + "/" + r[3];
		oa = rightPadWithSpaces(ff[2], ri);
		if (oa.length < r[4].length) {
			oa = hashq(r[4].substring(r[4].length - oa.length)) + oa;
		}
		o += oa;
		return o;
	}
	if ((r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(Math.max(r[1].length, r[4].length), 7);
		ff = SSF_frac(aval, Math.pow(10, ri) - 1, true);
		return (
			sign +
			(ff[0] || (ff[1] ? "" : "0")) +
			" " +
			(ff[1]
				? padWithSpaces(ff[1], ri) + r[2] + "/" + r[3] + rightPadWithSpaces(ff[2], ri)
				: " ".repeat(2 * ri + 1 + r[2].length + r[3].length))
		);
	}
	if ((r = fmt.match(/^[#0?]+$/))) {
		o = padRoundedZeros(val, 0);
		if (fmt.length <= o.length) {
			return o;
		}
		return hashq(fmt.substring(0, fmt.length - o.length)) + o;
	}
	if ((r = fmt.match(/^([#0?]+)\.([#0]+)$/))) {
		o = val.toFixed(Math.min(r[2].length, 10)).replace(/([^0])0+$/, "$1");
		ri = o.indexOf(".");
		const lres = fmt.indexOf(".") - ri;
		const rres = fmt.length - o.length - lres;
		return hashq(fmt.substring(0, lres) + o + fmt.substring(fmt.length - rres));
	}
	if ((r = fmt.match(/^00,000\.([#0]*0)$/))) {
		ri = dec(val, r[1].length);
		return val < 0
			? "-" + write_num_flt(type, fmt, -val)
			: commaify(flr(val))
					.replace(/^\d,\d{3}$/, "0$&")
					.replace(/^\d*$/, ($$) => "00," + ($$.length < 3 ? padWithZeros(0, 3 - $$.length) : "") + $$) +
					"." +
					padWithZeros(ri, r[1].length);
	}
	switch (fmt) {
		case "###,##0.00":
			return write_num_flt(type, "#,##0.00", val);
		case "###,###":
		case "##,###":
		case "#,###": {
			const x = commaify(padRoundedZeros(aval, 0));
			return x !== "0" ? sign + x : "";
		}
		case "###,###.00":
			return write_num_flt(type, "###,##0.00", val).replace(/^0\./, ".");
		case "#,###.00":
			return write_num_flt(type, "#,##0.00", val).replace(/^0\./, ".");
	}
	throw new Error("unsupported format |" + fmt + "|");
}

function write_num_int(type: string, fmt: string, val: number): string {
	if (type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
		const ffmt = fmt.replace(/\( */, "").replace(/ \)/, "").replace(/\)/, "");
		if (val >= 0) {
			return write_num_int("n", ffmt, val);
		}
		return "(" + write_num_int("n", ffmt, -val) + ")";
	}
	if (fmt.charCodeAt(fmt.length - 1) === 44) {
		return write_num_cm(type, fmt, val);
	}
	if (fmt.indexOf("%") !== -1) {
		return write_num_pct(type, fmt, val);
	}
	if (fmt.indexOf("E") !== -1) {
		return write_num_exp(fmt, val);
	}
	if (fmt.charCodeAt(0) === 36) {
		return "$" + write_num_int(type, fmt.substring(fmt.charAt(1) === " " ? 2 : 1), val);
	}
	let o: string;
	let r: RegExpMatchArray | null;
	let ri: number;
	let ff: number[];
	const aval = Math.abs(val);
	const sign = val < 0 ? "-" : "";
	if (fmt.match(/^00+$/)) {
		return sign + padWithZeros(aval, fmt.length);
	}
	if (fmt.match(/^[#?]+$/)) {
		o = "" + val;
		if (val === 0) {
			o = "";
		}
		return o.length > fmt.length ? o : hashq(fmt.substring(0, fmt.length - o.length)) + o;
	}
	if ((r = fmt.match(frac1))) {
		return write_num_f2(r, aval, sign);
	}
	if (fmt.match(/^#+0+$/)) {
		return sign + padWithZeros(aval, fmt.length - fmt.indexOf("0"));
	}
	if ((r = fmt.match(dec1))) {
		o = ("" + val).replace(/^([^.]+)$/, "$1." + hashq(r[1])).replace(/\.$/, "." + hashq(r[1]));
		o = o.replace(/\.(\d*)$/, ($$, $1) => "." + $1 + "0".repeat(hashq(r![1]).length - $1.length));
		return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./, ".");
	}
	fmt = fmt.replace(/^#+([0.])/, "$1");
	if ((r = fmt.match(/^(0*)\.(#*)$/))) {
		return (
			sign +
			("" + aval)
				.replace(/\.(\d*[1-9])0*$/, ".$1")
				.replace(/^(-?\d*)$/, "$1.")
				.replace(/^0\./, r[1].length ? "0." : ".")
		);
	}
	if (fmt.match(/^#{1,3},##0(\.?)$/)) {
		return sign + commaify("" + aval);
	}
	if ((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
		return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify("" + val) + "." + "0".repeat(r[1].length);
	}
	if ((r = fmt.match(/^#,#*,#0/))) {
		return write_num_int(type, fmt.replace(/^#,#*,/, ""), val);
	}
	if ((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
		o = reverseString(write_num_int(type, fmt.replace(/[\\-]/g, ""), val));
		ri = 0;
		return reverseString(
			reverseString(fmt.replace(/\\/g, "")).replace(/[0#]/g, (x) => {
				return ri < o.length ? o.charAt(ri++) : x === "0" ? "0" : "";
			}),
		);
	}
	if (fmt.match(phone)) {
		o = write_num_int(type, "##########", val);
		return "(" + o.substring(0, 3) + ") " + o.substring(3, 3) + "-" + o.substring(6);
	}
	let oa = "";
	if ((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(r[4].length, 7);
		ff = SSF_frac(aval, Math.pow(10, ri) - 1, false);
		o = sign;
		oa = write_num("n", r[1], ff[1]);
		if (oa.charAt(oa.length - 1) === " ") {
			oa = oa.substring(0, oa.length - 1) + "0";
		}
		o += oa + r[2] + "/" + r[3];
		oa = rightPadWithSpaces(ff[2], ri);
		if (oa.length < r[4].length) {
			oa = hashq(r[4].substring(r[4].length - oa.length)) + oa;
		}
		o += oa;
		return o;
	}
	if ((r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(Math.max(r[1].length, r[4].length), 7);
		ff = SSF_frac(aval, Math.pow(10, ri) - 1, true);
		return (
			sign +
			(ff[0] || (ff[1] ? "" : "0")) +
			" " +
			(ff[1]
				? padWithSpaces(ff[1], ri) + r[2] + "/" + r[3] + rightPadWithSpaces(ff[2], ri)
				: " ".repeat(2 * ri + 1 + r[2].length + r[3].length))
		);
	}
	if ((r = fmt.match(/^[#0?]+$/))) {
		o = "" + val;
		if (fmt.length <= o.length) {
			return o;
		}
		return hashq(fmt.substring(0, fmt.length - o.length)) + o;
	}
	if ((r = fmt.match(/^([#0]+)\.([#0]+)$/))) {
		o = val.toFixed(Math.min(r[2].length, 10)).replace(/([^0])0+$/, "$1");
		ri = o.indexOf(".");
		const lres = fmt.indexOf(".") - ri;
		const rres = fmt.length - o.length - lres;
		return hashq(fmt.substring(0, lres) + o + fmt.substring(fmt.length - rres));
	}
	if ((r = fmt.match(/^00,000\.([#0]*0)$/))) {
		return val < 0
			? "-" + write_num_int(type, fmt, -val)
			: commaify("" + val)
					.replace(/^\d,\d{3}$/, "0$&")
					.replace(/^\d*$/, ($$) => "00," + ($$.length < 3 ? padWithZeros(0, 3 - $$.length) : "") + $$) +
					"." +
					padWithZeros(0, r[1].length);
	}
	switch (fmt) {
		case "###,###":
		case "##,###":
		case "#,###": {
			const x = commaify("" + aval);
			return x !== "0" ? sign + x : "";
		}
		default:
			if (fmt.match(/\.[0#?]*$/)) {
				return (
					write_num_int(type, fmt.slice(0, fmt.lastIndexOf(".")), val) +
					hashq(fmt.slice(fmt.lastIndexOf(".")))
				);
			}
	}
	throw new Error("unsupported format |" + fmt + "|");
}

function write_num(type: string, fmt: string, val: number): string {
	return (val | 0) === val ? write_num_int(type, fmt, val) : write_num_flt(type, fmt, val);
}

function SSF_split_fmt(fmt: string): string[] {
	const out: string[] = [];
	let in_str = false;
	let j = 0;
	for (let i = 0; i < fmt.length; ++i) {
		switch (fmt.charCodeAt(i)) {
			case 34 /* '"' */:
				in_str = !in_str;
				break;
			case 95:
			case 42:
			case 92 /* '_' '*' '\\' */:
				++i;
				break;
			case 59 /* ';' */:
				out[out.length] = fmt.substring(j, i - j);
				j = i + 1;
		}
	}
	out[out.length] = fmt.substring(j);
	if (in_str) {
		throw new Error("Format |" + fmt + "| unterminated string ");
	}
	return out;
}

const SSF_abstime = /\[[HhMmSs\u0E0A\u0E19\u0E17]*\]/;

export function isDateFormat(fmt: string): boolean {
	let i = 0;
	let c = "";
	let o = "";
	while (i < fmt.length) {
		switch ((c = fmt.charAt(i))) {
			case "G":
				if (isGeneralFormat(fmt, i)) {
					i += 6;
				}
				i++;
				break;
			case '"':
				for (; fmt.charCodeAt(++i) !== 34 && i < fmt.length; ) {
					/* empty */
				}
				++i;
				break;
			case "\\":
				i += 2;
				break;
			case "_":
				i += 2;
				break;
			case "@":
				++i;
				break;
			case "B":
			case "b":
				if (fmt.charAt(i + 1) === "1" || fmt.charAt(i + 1) === "2") {
					return true;
				}
			/* falls through */
			case "M":
			case "D":
			case "Y":
			case "H":
			case "S":
			case "E":
			case "m":
			case "d":
			case "y":
			case "h":
			case "s":
			case "e":
			case "g":
				return true;
			case "A":
			case "a":
			case "\u4E0A":
				if (fmt.substring(i, 3).toUpperCase() === "A/P") {
					return true;
				}
				if (fmt.substring(i, 5).toUpperCase() === "AM/PM") {
					return true;
				}
				if (fmt.substring(i, 5).toUpperCase() === "\u4E0A\u5348/\u4E0B\u5348") {
					return true;
				}
				++i;
				break;
			case "[":
				o = c;
				while (fmt.charAt(i++) !== "]" && i < fmt.length) {
					o += fmt.charAt(i);
				}
				if (o.match(SSF_abstime)) {
					return true;
				}
				break;
			case ".":
			case "0":
			case "#":
				while (
					i < fmt.length &&
					("0#?.,E+-%".indexOf((c = fmt.charAt(++i))) > -1 ||
						(c === "\\" && fmt.charAt(i + 1) === "-" && "0#".indexOf(fmt.charAt(i + 2)) > -1))
				) {
					/* empty */
				}
				break;
			case "?":
				while (fmt.charAt(++i) === c) {
					/* empty */
				}
				break;
			case "*":
				++i;
				if (fmt.charAt(i) === " " || fmt.charAt(i) === "*") {
					++i;
				}
				break;
			case "(":
			case ")":
				++i;
				break;
			case "1":
			case "2":
			case "3":
			case "4":
			case "5":
			case "6":
			case "7":
			case "8":
			case "9":
				while (i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1) {
					/* empty */
				}
				break;
			case " ":
				++i;
				break;
			default:
				++i;
				break;
		}
	}
	return false;
}

interface FmtToken {
	type: string;
	value: string;
}

function eval_fmt(fmt: string, value: any, opts: any, flen: number): string {
	const out: (FmtToken | null)[] = [];
	let tokenStr = "";
	let i = 0;
	let char = "";
	let lastTokenType = "t";
	let dateVal: SSFDateVal | null = null;
	let scanIdx: number;
	let charCode: number;
	let hourFormat = "H";

	/* Tokenize */
	while (i < fmt.length) {
		switch ((char = fmt.charAt(i))) {
			case "G":
				if (!isGeneralFormat(fmt, i)) {
					throw new Error("unrecognized character " + char + " in " + fmt);
				}
				out[out.length] = { type: "G", value: "General" };
				i += 7;
				break;
			case '"':
				for (tokenStr = ""; (charCode = fmt.charCodeAt(++i)) !== 34 && i < fmt.length; ) {
					tokenStr += String.fromCharCode(charCode);
				}
				out[out.length] = { type: "t", value: tokenStr };
				++i;
				break;
			case "\\": {
				const nextChar = fmt.charAt(++i);
				const t2 = nextChar === "(" || nextChar === ")" ? nextChar : "t";
				out[out.length] = { type: t2, value: nextChar };
				++i;
				break;
			}
			case "_":
				out[out.length] = { type: "t", value: " " };
				i += 2;
				break;
			case "@":
				out[out.length] = { type: "T", value: value };
				++i;
				break;
			case "B":
			case "b":
				if (fmt.charAt(i + 1) === "1" || fmt.charAt(i + 1) === "2") {
					if (dateVal == null) {
						dateVal = parseExcelDateCode(value, opts, fmt.charAt(i + 1) === "2");
						if (dateVal == null) {
							return "";
						}
					}
					out[out.length] = { type: "X", value: fmt.substring(i, 2) };
					lastTokenType = char;
					i += 2;
					break;
				}
			/* falls through */
			case "M":
			case "D":
			case "Y":
			case "H":
			case "S":
			case "E":
				char = char.toLowerCase();
			/* falls through */
			case "m":
			case "d":
			case "y":
			case "h":
			case "s":
			case "e":
			case "g":
				if (value < 0) {
					return "";
				}
				if (dateVal == null) {
					dateVal = parseExcelDateCode(value, opts);
					if (dateVal == null) {
						return "";
					}
				}
				tokenStr = char;
				while (++i < fmt.length && fmt.charAt(i).toLowerCase() === char) {
					tokenStr += char;
				}
				if (char === "m" && lastTokenType.toLowerCase() === "h") {
					char = "M";
				}
				if (char === "h") {
					char = hourFormat;
				}
				out[out.length] = { type: char, value: tokenStr };
				lastTokenType = char;
				break;
			case "A":
			case "a":
			case "\u4E0A": {
				const ampmToken: FmtToken = { type: char, value: char };
				if (dateVal == null) {
					dateVal = parseExcelDateCode(value, opts);
				}
				if (fmt.substring(i, 3).toUpperCase() === "A/P") {
					if (dateVal != null) {
						ampmToken.value = dateVal.hours >= 12 ? fmt.charAt(i + 2) : char;
					}
					ampmToken.type = "T";
					hourFormat = "h";
					i += 3;
				} else if (fmt.substring(i, 5).toUpperCase() === "AM/PM") {
					if (dateVal != null) {
						ampmToken.value = dateVal.hours >= 12 ? "PM" : "AM";
					}
					ampmToken.type = "T";
					i += 5;
					hourFormat = "h";
				} else if (fmt.substring(i, 5).toUpperCase() === "\u4E0A\u5348/\u4E0B\u5348") {
					if (dateVal != null) {
						ampmToken.value = dateVal.hours >= 12 ? "\u4E0B\u5348" : "\u4E0A\u5348";
					}
					ampmToken.type = "T";
					i += 5;
					hourFormat = "h";
				} else {
					ampmToken.type = "t";
					++i;
				}
				if (dateVal == null && ampmToken.type === "T") {
					return "";
				}
				out[out.length] = ampmToken;
				lastTokenType = char;
				break;
			}
			case "[":
				tokenStr = char;
				while (fmt.charAt(i++) !== "]" && i < fmt.length) {
					tokenStr += fmt.charAt(i);
				}
				if (tokenStr.slice(-1) !== "]") {
					throw new Error('unterminated "[" block: |' + tokenStr + "|");
				}
				if (tokenStr.match(SSF_abstime)) {
					if (dateVal == null) {
						dateVal = parseExcelDateCode(value, opts);
						if (dateVal == null) {
							return "";
						}
					}
					out[out.length] = { type: "Z", value: tokenStr.toLowerCase() };
					lastTokenType = tokenStr.charAt(1);
				} else if (tokenStr.indexOf("$") > -1) {
					tokenStr = (tokenStr.match(/\$([^-[\]]*)/) || [])[1] || "$";
					if (!isDateFormat(fmt)) {
						out[out.length] = { type: "t", value: tokenStr };
					}
				}
				break;
			case ".":
				if (dateVal != null) {
					tokenStr = char;
					while (++i < fmt.length && (char = fmt.charAt(i)) === "0") {
						tokenStr += char;
					}
					out[out.length] = { type: "s", value: tokenStr };
					break;
				}
			/* falls through */
			case "0":
			case "#":
				tokenStr = char;
				while (++i < fmt.length && "0#?.,E+-%".indexOf((char = fmt.charAt(i))) > -1) {
					tokenStr += char;
				}
				out[out.length] = { type: "n", value: tokenStr };
				break;
			case "?":
				tokenStr = char;
				while (fmt.charAt(++i) === char) {
					tokenStr += char;
				}
				out[out.length] = { type: char, value: tokenStr };
				lastTokenType = char;
				break;
			case "*":
				++i;
				if (fmt.charAt(i) === " " || fmt.charAt(i) === "*") {
					++i;
				}
				break;
			case "(":
			case ")":
				out[out.length] = { type: flen === 1 ? "t" : char, value: char };
				++i;
				break;
			case "1":
			case "2":
			case "3":
			case "4":
			case "5":
			case "6":
			case "7":
			case "8":
			case "9":
				tokenStr = char;
				while (i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1) {
					tokenStr += fmt.charAt(i);
				}
				out[out.length] = { type: "D", value: tokenStr };
				break;
			case " ":
				out[out.length] = { type: char, value: char };
				++i;
				break;
			case "$":
				out[out.length] = { type: "t", value: "$" };
				++i;
				break;
			default:
				if (",$-+/():!^&'~{}<>=\u20ACacfijklopqrtuvwxzP".indexOf(char) === -1) {
					throw new Error("unrecognized character " + char + " in " + fmt);
				}
				out[out.length] = { type: "t", value: char };
				++i;
				break;
		}
	}

	/* Scan for date/time parts */
	let dateTimePrecision = 0;
	let subSecondDigits = 0;
	let ssm: RegExpMatchArray | null;
	for (i = out.length - 1, lastTokenType = "t"; i >= 0; --i) {
		if (!out[i]) {
			continue;
		}
		switch (out[i]!.type) {
			case "h":
			case "H":
				out[i]!.type = hourFormat;
				lastTokenType = "h";
				if (dateTimePrecision < 1) {
					dateTimePrecision = 1;
				}
				break;
			case "s":
				if ((ssm = out[i]!.value.match(/\.0+$/))) {
					subSecondDigits = Math.max(subSecondDigits, ssm[0].length - 1);
					dateTimePrecision = 4;
				}
				if (dateTimePrecision < 3) {
					dateTimePrecision = 3;
				}
			/* falls through */
			case "d":
			case "y":
			case "e":
				lastTokenType = out[i]!.type;
				break;
			case "M":
				lastTokenType = out[i]!.type;
				if (dateTimePrecision < 2) {
					dateTimePrecision = 2;
				}
				break;
			case "m":
				if (lastTokenType === "s") {
					out[i]!.type = "M";
					if (dateTimePrecision < 2) {
						dateTimePrecision = 2;
					}
				}
				break;
			case "X":
				break;
			case "Z":
				if (dateTimePrecision < 1 && out[i]!.value.match(/[Hh]/)) {
					dateTimePrecision = 1;
				}
				if (dateTimePrecision < 2 && out[i]!.value.match(/[Mm]/)) {
					dateTimePrecision = 2;
				}
				if (dateTimePrecision < 3 && out[i]!.value.match(/[Ss]/)) {
					dateTimePrecision = 3;
				}
		}
	}

	/* time rounding */
	if (dateVal) {
		let _dt: SSFDateVal | null;
		switch (dateTimePrecision) {
			case 0:
				break;
			case 1:
			case 2:
			case 3:
				if (dateVal.subSeconds >= 0.5) {
					dateVal.subSeconds = 0;
					++dateVal.seconds;
				}
				if (dateVal.seconds >= 60) {
					dateVal.seconds = 0;
					++dateVal.minutes;
				}
				if (dateVal.minutes >= 60) {
					dateVal.minutes = 0;
					++dateVal.hours;
				}
				if (dateVal.hours >= 24) {
					dateVal.hours = 0;
					++dateVal.daySerial;
					_dt = parseExcelDateCode(dateVal.daySerial);
					if (_dt) {
						_dt.subSeconds = dateVal.subSeconds;
						_dt.seconds = dateVal.seconds;
						_dt.minutes = dateVal.minutes;
						_dt.hours = dateVal.hours;
						dateVal = _dt;
					}
				}
				break;
			case 4:
				switch (subSecondDigits) {
					case 1:
						dateVal.subSeconds = Math.round(dateVal.subSeconds * 10) / 10;
						break;
					case 2:
						dateVal.subSeconds = Math.round(dateVal.subSeconds * 100) / 100;
						break;
					case 3:
						dateVal.subSeconds = Math.round(dateVal.subSeconds * 1000) / 1000;
						break;
				}
				if (dateVal.subSeconds >= 1) {
					dateVal.subSeconds = 0;
					++dateVal.seconds;
				}
				if (dateVal.seconds >= 60) {
					dateVal.seconds = 0;
					++dateVal.minutes;
				}
				if (dateVal.minutes >= 60) {
					dateVal.minutes = 0;
					++dateVal.hours;
				}
				if (dateVal.hours >= 24) {
					dateVal.hours = 0;
					++dateVal.daySerial;
					_dt = parseExcelDateCode(dateVal.daySerial);
					if (_dt) {
						_dt.subSeconds = dateVal.subSeconds;
						_dt.seconds = dateVal.seconds;
						_dt.minutes = dateVal.minutes;
						_dt.hours = dateVal.hours;
						dateVal = _dt;
					}
				}
				break;
		}
	}

	/* replace fields */
	let numberFmtStr = "";
	let numFmtIdx: number;
	for (i = 0; i < out.length; ++i) {
		if (!out[i]) {
			continue;
		}
		switch (out[i]!.type) {
			case "t":
			case "T":
			case " ":
			case "D":
				break;
			case "X":
				out[i]!.value = "";
				out[i]!.type = ";";
				break;
			case "d":
			case "m":
			case "y":
			case "h":
			case "H":
			case "M":
			case "s":
			case "e":
			case "b":
			case "Z":
				out[i]!.value = SSF_write_date(out[i]!.type.charCodeAt(0), out[i]!.value, dateVal!, subSecondDigits);
				out[i]!.type = "t";
				break;
			case "n":
			case "?":
				numFmtIdx = i + 1;
				while (
					out[numFmtIdx] != null &&
					((char = out[numFmtIdx]!.type) === "?" ||
						char === "D" ||
						((char === " " || char === "t") &&
							out[numFmtIdx + 1] != null &&
							(out[numFmtIdx + 1]!.type === "?" || (out[numFmtIdx + 1]!.type === "t" && out[numFmtIdx + 1]!.value === "/"))) ||
						(out[i]!.type === "(" && (char === " " || char === "n" || char === ")")) ||
						(char === "t" &&
							(out[numFmtIdx]!.value === "/" ||
								(out[numFmtIdx]!.value === " " && out[numFmtIdx + 1] != null && out[numFmtIdx + 1]!.type === "?"))))
				) {
					out[i]!.value += out[numFmtIdx]!.value;
					out[numFmtIdx] = { value: "", type: ";" };
					++numFmtIdx;
				}
				numberFmtStr += out[i]!.value;
				i = numFmtIdx - 1;
				break;
			case "G":
				out[i]!.type = "t";
				out[i]!.value = SSF_general(value, opts);
				break;
		}
	}

	let partialValue = "";
	let adjustedValue: number;
	let formattedNumber: string;
	if (numberFmtStr.length > 0) {
		if (numberFmtStr.charCodeAt(0) === 40) {
			adjustedValue = value < 0 && numberFmtStr.charCodeAt(0) === 45 ? -value : value;
			formattedNumber = write_num("n", numberFmtStr, adjustedValue);
		} else {
			adjustedValue = value < 0 && flen > 1 ? -value : value;
			formattedNumber = write_num("n", numberFmtStr, adjustedValue);
			if (adjustedValue < 0 && out[0] && out[0].type === "t") {
				formattedNumber = formattedNumber.substring(1);
				out[0].value = "-" + out[0].value;
			}
		}
		numFmtIdx = formattedNumber.length - 1;
		let decpt = out.length;
		for (i = 0; i < out.length; ++i) {
			if (out[i] != null && out[i]!.type !== "t" && out[i]!.value.indexOf(".") > -1) {
				decpt = i;
				break;
			}
		}
		let lasti = out.length;
		if (decpt === out.length && formattedNumber.indexOf("E") === -1) {
			for (i = out.length - 1; i >= 0; --i) {
				if (out[i] == null || "n?".indexOf(out[i]!.type) === -1) {
					continue;
				}
				if (numFmtIdx >= out[i]!.value.length - 1) {
					numFmtIdx -= out[i]!.value.length;
					out[i]!.value = formattedNumber.substring(numFmtIdx + 1, out[i]!.value.length);
				} else if (numFmtIdx < 0) {
					out[i]!.value = "";
				} else {
					out[i]!.value = formattedNumber.substring(0, numFmtIdx + 1);
					numFmtIdx = -1;
				}
				out[i]!.type = "t";
				lasti = i;
			}
			if (numFmtIdx >= 0 && lasti < out.length) {
				out[lasti]!.value = formattedNumber.substring(0, numFmtIdx + 1) + out[lasti]!.value;
			}
		} else if (decpt !== out.length && formattedNumber.indexOf("E") === -1) {
			numFmtIdx = formattedNumber.indexOf(".") - 1;
			for (i = decpt; i >= 0; --i) {
				if (out[i] == null || "n?".indexOf(out[i]!.type) === -1) {
					continue;
				}
				scanIdx = out[i]!.value.indexOf(".") > -1 && i === decpt ? out[i]!.value.indexOf(".") - 1 : out[i]!.value.length - 1;
				partialValue = out[i]!.value.substring(scanIdx + 1);
				for (; scanIdx >= 0; --scanIdx) {
					if (numFmtIdx >= 0 && (out[i]!.value.charAt(scanIdx) === "0" || out[i]!.value.charAt(scanIdx) === "#")) {
						partialValue = formattedNumber.charAt(numFmtIdx--) + partialValue;
					}
				}
				out[i]!.value = partialValue;
				out[i]!.type = "t";
				lasti = i;
			}
			if (numFmtIdx >= 0 && lasti < out.length) {
				out[lasti]!.value = formattedNumber.substring(0, numFmtIdx + 1) + out[lasti]!.value;
			}
			numFmtIdx = formattedNumber.indexOf(".") + 1;
			for (i = decpt; i < out.length; ++i) {
				if (out[i] == null || ("n?(".indexOf(out[i]!.type) === -1 && i !== decpt)) {
					continue;
				}
				scanIdx = out[i]!.value.indexOf(".") > -1 && i === decpt ? out[i]!.value.indexOf(".") + 1 : 0;
				partialValue = out[i]!.value.substring(0, scanIdx);
				for (; scanIdx < out[i]!.value.length; ++scanIdx) {
					if (numFmtIdx < formattedNumber.length) {
						partialValue += formattedNumber.charAt(numFmtIdx++);
					}
				}
				out[i]!.value = partialValue;
				out[i]!.type = "t";
				lasti = i;
			}
		}
	}
	for (i = 0; i < out.length; ++i) {
		if (out[i] != null && "n?".indexOf(out[i]!.type) > -1) {
			adjustedValue = flen > 1 && value < 0 && i > 0 && out[i - 1]!.value === "-" ? -value : value;
			out[i]!.value = write_num(out[i]!.type, out[i]!.value, adjustedValue);
			out[i]!.type = "t";
		}
	}
	let retval = "";
	for (i = 0; i !== out.length; ++i) {
		if (out[i] != null) {
			retval += out[i]!.value;
		}
	}
	return retval;
}

const cfregex2 = /\[(=|>[=]?|<[>=]?)(-?\d+(?:\.\d*)?)\]/;
function chkcond(v: number, rr: RegExpMatchArray | null): boolean {
	if (rr == null) {
		return false;
	}
	const thresh = parseFloat(rr[2]);
	switch (rr[1]) {
		case "=":
			if (v == thresh) {
				return true;
			}
			break;
		case ">":
			if (v > thresh) {
				return true;
			}
			break;
		case "<":
			if (v < thresh) {
				return true;
			}
			break;
		case "<>":
			if (v != thresh) {
				return true;
			}
			break;
		case ">=":
			if (v >= thresh) {
				return true;
			}
			break;
		case "<=":
			if (v <= thresh) {
				return true;
			}
			break;
	}
	return false;
}

function choose_fmt(fmtStr: string, value: any): [number, string] {
	let fmt = SSF_split_fmt(fmtStr);
	const sectionCount = fmt.length;
	const lat = fmt[sectionCount - 1].indexOf("@");
	let ll = sectionCount;
	if (sectionCount < 4 && lat > -1) {
		--ll;
	}
	if (fmt.length > 4) {
		throw new Error("cannot find right format for |" + fmt.join("|") + "|");
	}
	if (typeof value !== "number") {
		return [4, fmt.length === 4 || lat > -1 ? fmt[fmt.length - 1] : "@"];
	}
	if (typeof value === "number" && !isFinite(value)) {
		value = 0;
	}
	switch (fmt.length) {
		case 1:
			fmt = lat > -1 ? ["General", "General", "General", fmt[0]] : [fmt[0], fmt[0], fmt[0], "@"];
			break;
		case 2:
			fmt = lat > -1 ? [fmt[0], fmt[0], fmt[0], fmt[1]] : [fmt[0], fmt[1], fmt[0], "@"];
			break;
		case 3:
			fmt = lat > -1 ? [fmt[0], fmt[1], fmt[0], fmt[2]] : [fmt[0], fmt[1], fmt[2], "@"];
			break;
		case 4:
			break;
	}
	const selectedFmt = value > 0 ? fmt[0] : value < 0 ? fmt[1] : fmt[2];
	if (fmt[0].indexOf("[") === -1 && fmt[1].indexOf("[") === -1) {
		return [ll, selectedFmt];
	}
	if (fmt[0].match(/\[[=<>]/) != null || fmt[1].match(/\[[=<>]/) != null) {
		const m1 = fmt[0].match(cfregex2);
		const m2 = fmt[1].match(cfregex2);
		return chkcond(value, m1)
			? [ll, fmt[0]]
			: chkcond(value, m2)
				? [ll, fmt[1]]
				: [ll, fmt[m1 != null && m2 != null ? 2 : 1]];
	}
	return [ll, selectedFmt];
}

/** Format a value using an Excel number format string */
export function formatNumber(fmt: string | number, value: any, options?: any): string {
	if (options == null) {
		options = {};
	}
	let sfmt = "";
	switch (typeof fmt) {
		case "string":
			if (fmt === "m/d/yy" && options.dateNF) {
				sfmt = options.dateNF;
			} else {
				sfmt = fmt;
			}
			break;
		case "number":
			if (fmt === 14 && options.dateNF) {
				sfmt = options.dateNF;
			} else {
				sfmt = (options.table != null ? options.table : formatTable)[fmt];
			}
			if (sfmt == null) {
				sfmt = (options.table && options.table[DEFAULT_FORMAT_MAP[fmt]]) || formatTable[DEFAULT_FORMAT_MAP[fmt]];
			}
			if (sfmt == null) {
				sfmt = DEFAULT_FORMAT_STRINGS[fmt] || "General";
			}
			break;
	}
	if (isGeneralFormat(sfmt, 0)) {
		return SSF_general(value, options);
	}
	if (value instanceof Date) {
		value = dateToSerialNumber(value, options.date1904);
	}
	const chosenFmt = choose_fmt(sfmt, value);
	if (isGeneralFormat(chosenFmt[1])) {
		return SSF_general(value, options);
	}
	if (value === true) {
		value = "TRUE";
	} else if (value === false) {
		value = "FALSE";
	} else if (value === "" || value == null) {
		return "";
	} else if (isNaN(value) && chosenFmt[1].indexOf("0") > -1) {
		return "#NUM!";
	} else if (!isFinite(value) && chosenFmt[1].indexOf("0") > -1) {
		return "#DIV/0!";
	}
	return eval_fmt(chosenFmt[1], value, options, chosenFmt[0]);
}
