/* ssf.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* Ported to TypeScript for xlsx-format. 1:1 faithful port. */

import { datenum } from "../utils/date.js";
import { fill } from "../utils/helpers.js";
import { table_fmt, SSF_default_map, SSF_default_str } from "./table.js";

function _strrev(x: string): string {
	let o = "";
	let i = x.length - 1;
	while (i >= 0) {
		o += x.charAt(i--);
	}
	return o;
}
function pad0(v: any, d: number): string {
	const t = "" + v;
	return t.length >= d ? t : fill("0", d - t.length) + t;
}
function pad_(v: any, d: number): string {
	const t = "" + v;
	return t.length >= d ? t : fill(" ", d - t.length) + t;
}
function rpad_(v: any, d: number): string {
	const t = "" + v;
	return t.length >= d ? t : t + fill(" ", d - t.length);
}
function pad0r1(v: any, d: number): string {
	const t = "" + Math.round(v);
	return t.length >= d ? t : fill("0", d - t.length) + t;
}
function pad0r2(v: any, d: number): string {
	const t = "" + v;
	return t.length >= d ? t : fill("0", d - t.length) + t;
}
const p2_32 = Math.pow(2, 32);
function pad0r(v: any, d: number): string {
	if (v > p2_32 || v < -p2_32) {
		return pad0r1(v, d);
	}
	const i = Math.round(v);
	return pad0r2(i, d);
}

function SSF_isgeneral(s: string, i?: number): boolean {
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

const days: string[][] = [
	["Sun", "Sunday"],
	["Mon", "Monday"],
	["Tue", "Tuesday"],
	["Wed", "Wednesday"],
	["Thu", "Thursday"],
	["Fri", "Friday"],
	["Sat", "Saturday"],
];
const months: string[][] = [
	["J", "Jan", "January"],
	["F", "Feb", "February"],
	["M", "Mar", "March"],
	["A", "Apr", "April"],
	["M", "May", "May"],
	["J", "Jun", "June"],
	["J", "Jul", "July"],
	["A", "Aug", "August"],
	["S", "Sep", "September"],
	["O", "Oct", "October"],
	["N", "Nov", "November"],
	["D", "Dec", "December"],
];

interface SSFDateVal {
	D: number;
	T: number;
	u: number;
	y: number;
	m: number;
	d: number;
	H: number;
	M: number;
	S: number;
	q: number;
}

function SSF_normalize_xl_unsafe(v: number): number {
	const s = v.toPrecision(16);
	if (s.indexOf("e") > -1) {
		const m = s.slice(0, s.indexOf("e"));
		const ml =
			m.indexOf(".") > -1
				? m.slice(0, m.slice(0, 2) === "0." ? 17 : 16)
				: m.slice(0, 15) + fill("0", m.length - 15);
		return +ml + +("1" + s.slice(s.indexOf("e"))) - 1 || +s;
	}
	const n =
		s.indexOf(".") > -1 ? s.slice(0, s.slice(0, 2) === "0." ? 17 : 16) : s.slice(0, 15) + fill("0", s.length - 15);
	return Number(n);
}

function SSF_fix_hijri(_date: Date, o: number[]): number {
	o[0] -= 581;
	const dow = _date.getDay();
	if (_date.getTime() < -2203891200000) {
		return (dow + 6) % 7;
	}
	return dow;
}

export function SSF_parse_date_code(v: number, opts?: any, b2?: boolean): SSFDateVal | null {
	if (v > 2958465 || v < 0) {
		return null;
	}
	v = SSF_normalize_xl_unsafe(v);
	let date = v | 0;
	let time = Math.floor(86400 * (v - date));
	const out: SSFDateVal = {
		D: date,
		T: time,
		u: 86400 * (v - date) - time,
		y: 0,
		m: 0,
		d: 0,
		H: 0,
		M: 0,
		S: 0,
		q: 0,
	};
	if (Math.abs(out.u) < 1e-6) {
		out.u = 0;
	}
	if (opts && opts.date1904) {
		date += 1462;
	}
	if (out.u > 0.9999) {
		out.u = 0;
		if (++time === 86400) {
			out.T = time = 0;
			++date;
			++out.D;
		}
	}
	let dout: number[];
	let dow = 0;
	if (date === 60) {
		dout = b2 ? [1317, 10, 29] : [1900, 2, 29];
		dow = 3;
	} else if (date === 0) {
		dout = b2 ? [1317, 8, 29] : [1900, 1, 0];
		dow = 6;
	} else {
		if (date > 60) {
			--date;
		}
		const d = new Date(1900, 0, 1);
		d.setDate(d.getDate() + date - 1);
		dout = [d.getFullYear(), d.getMonth() + 1, d.getDate()];
		dow = d.getDay();
		if (date < 60) {
			dow = (dow + 6) % 7;
		}
		if (b2) {
			dow = SSF_fix_hijri(d, dout);
		}
	}
	out.y = dout[0];
	out.m = dout[1];
	out.d = dout[2];
	out.S = time % 60;
	time = Math.floor(time / 60);
	out.M = time % 60;
	time = Math.floor(time / 60);
	out.H = time;
	out.q = dow;
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
				return SSF_format(14, datenum(v, opts && opts.date1904), opts);
			}
	}
	throw new Error("unsupported value in General format: " + v);
}

function SSF_write_date(type: number, fmt: string, val: SSFDateVal, ss0?: number): string {
	let o = "";
	let ss = 0;
	let tt = 0;
	let y = val.y;
	let out: number = 0;
	let outl = 0;
	switch (type) {
		case 98 /* 'b' buddhist year */:
			y = val.y + 543;
		/* falls through */
		case 121 /* 'y' year */:
			switch (fmt.length) {
				case 1:
				case 2:
					out = y % 100;
					outl = 2;
					break;
				default:
					out = y % 10000;
					outl = 4;
					break;
			}
			break;
		case 109 /* 'm' month */:
			switch (fmt.length) {
				case 1:
				case 2:
					out = val.m;
					outl = fmt.length;
					break;
				case 3:
					return months[val.m - 1][1];
				case 5:
					return months[val.m - 1][0];
				default:
					return months[val.m - 1][2];
			}
			break;
		case 100 /* 'd' day */:
			switch (fmt.length) {
				case 1:
				case 2:
					out = val.d;
					outl = fmt.length;
					break;
				case 3:
					return days[val.q][0];
				default:
					return days[val.q][1];
			}
			break;
		case 104 /* 'h' 12-hour */:
			switch (fmt.length) {
				case 1:
				case 2:
					out = 1 + ((val.H + 11) % 12);
					outl = fmt.length;
					break;
				default:
					throw new Error("bad hour format: " + fmt);
			}
			break;
		case 72 /* 'H' 24-hour */:
			switch (fmt.length) {
				case 1:
				case 2:
					out = val.H;
					outl = fmt.length;
					break;
				default:
					throw new Error("bad hour format: " + fmt);
			}
			break;
		case 77 /* 'M' minutes */:
			switch (fmt.length) {
				case 1:
				case 2:
					out = val.M;
					outl = fmt.length;
					break;
				default:
					throw new Error("bad minute format: " + fmt);
			}
			break;
		case 115 /* 's' seconds */:
			if (fmt !== "s" && fmt !== "ss" && fmt !== ".0" && fmt !== ".00" && fmt !== ".000") {
				throw new Error("bad second format: " + fmt);
			}
			if (val.u === 0 && (fmt === "s" || fmt === "ss")) {
				return pad0(val.S, fmt.length);
			}
			if (ss0! >= 2) {
				tt = ss0 === 3 ? 1000 : 100;
			} else {
				tt = ss0 === 1 ? 10 : 1;
			}
			ss = Math.round(tt * (val.S + val.u));
			if (ss >= 60 * tt) {
				ss = 0;
			}
			if (fmt === "s") {
				return ss === 0 ? "0" : "" + ss / tt;
			}
			o = pad0(ss, 2 + ss0!);
			if (fmt === "ss") {
				return o.substring(0, 2);
			}
			return "." + o.substring(2, fmt.length - 1);
		case 90 /* 'Z' absolute time */:
			switch (fmt) {
				case "[h]":
				case "[hh]":
					out = val.D * 24 + val.H;
					break;
				case "[m]":
				case "[mm]":
					out = (val.D * 24 + val.H) * 60 + val.M;
					break;
				case "[s]":
				case "[ss]":
					out = ((val.D * 24 + val.H) * 60 + val.M) * 60 + (ss0 === 0 ? Math.round(val.S + val.u) : val.S);
					break;
				default:
					throw new Error("bad abstime format: " + fmt);
			}
			outl = fmt.length === 3 ? 1 : 2;
			break;
		case 101 /* 'e' era */:
			out = y;
			outl = 1;
			break;
	}
	return outl > 0 ? pad0(out, outl) : "";
}

function commaify(s: string): string {
	const w = 3;
	if (s.length <= w) {
		return s;
	}
	const j = s.length % w;
	let o = s.substring(0, j);
	for (let i = j; i !== s.length; i += w) {
		o += (o.length > 0 ? "," : "") + s.substring(i, w);
	}
	return o;
}

const pct1 = /%/g;
function write_num_pct(type: string, fmt: string, val: number): string {
	const sfmt = fmt.replace(pct1, "");
	const mul = fmt.length - sfmt.length;
	return write_num(type, sfmt, val * Math.pow(10, 2 * mul)) + fill("%", mul);
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

function SSF_frac(x: number, D: number, mixed?: boolean): number[] {
	const sgn = x < 0 ? -1 : 1;
	let B = x * sgn;
	let P_2 = 0,
		P_1 = 1,
		P = 0;
	let Q_2 = 1,
		Q_1 = 0,
		Q = 0;
	let A = Math.floor(B);
	while (Q_1 < D) {
		A = Math.floor(B);
		P = A * P_1 + P_2;
		Q = A * Q_1 + Q_2;
		if (B - A < 0.00000005) {
			break;
		}
		B = 1 / (B - A);
		P_2 = P_1;
		P_1 = P;
		Q_2 = Q_1;
		Q_1 = Q;
	}
	if (Q > D) {
		if (Q_1 > D) {
			Q = Q_2;
			P = P_2;
		} else {
			Q = Q_1;
			P = P_1;
		}
	}
	if (!mixed) {
		return [0, sgn * P, Q];
	}
	const q = Math.floor((sgn * P) / Q);
	return [q, sgn * P - q * Q, Q];
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
			? fill(" ", r[1].length + 1 + r[4].length)
			: pad_(myn, r[1].length) + r[2] + "/" + r[3] + pad0(myd, r[4].length))
	);
}
function write_num_f2(r: string[], aval: number, sign: string): string {
	return sign + (aval === 0 ? "" : "" + aval) + fill(" ", r[1].length + 2 + r[4].length);
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
		return sign + pad0r(aval, fmt.length);
	}
	if (fmt.match(/^[#?]+$/)) {
		o = pad0r(val, 0);
		if (o === "0") {
			o = "";
		}
		return o.length > fmt.length ? o : hashq(fmt.substring(0, fmt.length - o.length)) + o;
	}
	if ((r = fmt.match(frac1))) {
		return write_num_f1(r, aval, sign);
	}
	if (fmt.match(/^#+0+$/)) {
		return sign + pad0r(aval, fmt.length - fmt.indexOf("0"));
	}
	if ((r = fmt.match(dec1))) {
		o = rnd(val, r[1].length)
			.replace(/^([^.]+)$/, "$1." + hashq(r[1]))
			.replace(/\.$/, "." + hashq(r[1]))
			.replace(/\.(\d*)$/, ($$, $1) => "." + $1 + fill("0", hashq(r![1]).length - $1.length));
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
		return sign + commaify(pad0r(aval, 0));
	}
	if ((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
		return val < 0
			? "-" + write_num_flt(type, fmt, -val)
			: commaify("" + (Math.floor(val) + carry(val, r[1].length))) +
					"." +
					pad0(dec(val, r[1].length), r[1].length);
	}
	if ((r = fmt.match(/^#,#*,#0/))) {
		return write_num_flt(type, fmt.replace(/^#,#*,/, ""), val);
	}
	if ((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
		o = _strrev(write_num_flt(type, fmt.replace(/[\\-]/g, ""), val));
		ri = 0;
		return _strrev(
			_strrev(fmt.replace(/\\/g, "")).replace(/[0#]/g, (x) => {
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
		oa = rpad_(ff[2], ri);
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
				? pad_(ff[1], ri) + r[2] + "/" + r[3] + rpad_(ff[2], ri)
				: fill(" ", 2 * ri + 1 + r[2].length + r[3].length))
		);
	}
	if ((r = fmt.match(/^[#0?]+$/))) {
		o = pad0r(val, 0);
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
					.replace(/^\d*$/, ($$) => "00," + ($$.length < 3 ? pad0(0, 3 - $$.length) : "") + $$) +
					"." +
					pad0(ri, r[1].length);
	}
	switch (fmt) {
		case "###,##0.00":
			return write_num_flt(type, "#,##0.00", val);
		case "###,###":
		case "##,###":
		case "#,###": {
			const x = commaify(pad0r(aval, 0));
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
		return sign + pad0(aval, fmt.length);
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
		return sign + pad0(aval, fmt.length - fmt.indexOf("0"));
	}
	if ((r = fmt.match(dec1))) {
		o = ("" + val).replace(/^([^.]+)$/, "$1." + hashq(r[1])).replace(/\.$/, "." + hashq(r[1]));
		o = o.replace(/\.(\d*)$/, ($$, $1) => "." + $1 + fill("0", hashq(r![1]).length - $1.length));
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
		return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify("" + val) + "." + fill("0", r[1].length);
	}
	if ((r = fmt.match(/^#,#*,#0/))) {
		return write_num_int(type, fmt.replace(/^#,#*,/, ""), val);
	}
	if ((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
		o = _strrev(write_num_int(type, fmt.replace(/[\\-]/g, ""), val));
		ri = 0;
		return _strrev(
			_strrev(fmt.replace(/\\/g, "")).replace(/[0#]/g, (x) => {
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
		oa = rpad_(ff[2], ri);
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
				? pad_(ff[1], ri) + r[2] + "/" + r[3] + rpad_(ff[2], ri)
				: fill(" ", 2 * ri + 1 + r[2].length + r[3].length))
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
					.replace(/^\d*$/, ($$) => "00," + ($$.length < 3 ? pad0(0, 3 - $$.length) : "") + $$) +
					"." +
					pad0(0, r[1].length);
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

export function fmt_is_date(fmt: string): boolean {
	let i = 0;
	let c = "";
	let o = "";
	while (i < fmt.length) {
		switch ((c = fmt.charAt(i))) {
			case "G":
				if (SSF_isgeneral(fmt, i)) {
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
	t: string;
	v: string;
}

function eval_fmt(fmt: string, v: any, opts: any, flen: number): string {
	const out: (FmtToken | null)[] = [];
	let o = "";
	let i = 0;
	let c = "";
	let lst = "t";
	let dt: SSFDateVal | null = null;
	let j: number;
	let cc: number;
	let hr = "H";

	/* Tokenize */
	while (i < fmt.length) {
		switch ((c = fmt.charAt(i))) {
			case "G":
				if (!SSF_isgeneral(fmt, i)) {
					throw new Error("unrecognized character " + c + " in " + fmt);
				}
				out[out.length] = { t: "G", v: "General" };
				i += 7;
				break;
			case '"':
				for (o = ""; (cc = fmt.charCodeAt(++i)) !== 34 && i < fmt.length; ) {
					o += String.fromCharCode(cc);
				}
				out[out.length] = { t: "t", v: o };
				++i;
				break;
			case "\\": {
				const w = fmt.charAt(++i);
				const t2 = w === "(" || w === ")" ? w : "t";
				out[out.length] = { t: t2, v: w };
				++i;
				break;
			}
			case "_":
				out[out.length] = { t: "t", v: " " };
				i += 2;
				break;
			case "@":
				out[out.length] = { t: "T", v: v };
				++i;
				break;
			case "B":
			case "b":
				if (fmt.charAt(i + 1) === "1" || fmt.charAt(i + 1) === "2") {
					if (dt == null) {
						dt = SSF_parse_date_code(v, opts, fmt.charAt(i + 1) === "2");
						if (dt == null) {
							return "";
						}
					}
					out[out.length] = { t: "X", v: fmt.substring(i, 2) };
					lst = c;
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
				c = c.toLowerCase();
			/* falls through */
			case "m":
			case "d":
			case "y":
			case "h":
			case "s":
			case "e":
			case "g":
				if (v < 0) {
					return "";
				}
				if (dt == null) {
					dt = SSF_parse_date_code(v, opts);
					if (dt == null) {
						return "";
					}
				}
				o = c;
				while (++i < fmt.length && fmt.charAt(i).toLowerCase() === c) {
					o += c;
				}
				if (c === "m" && lst.toLowerCase() === "h") {
					c = "M";
				}
				if (c === "h") {
					c = hr;
				}
				out[out.length] = { t: c, v: o };
				lst = c;
				break;
			case "A":
			case "a":
			case "\u4E0A": {
				const q: FmtToken = { t: c, v: c };
				if (dt == null) {
					dt = SSF_parse_date_code(v, opts);
				}
				if (fmt.substring(i, 3).toUpperCase() === "A/P") {
					if (dt != null) {
						q.v = dt.H >= 12 ? fmt.charAt(i + 2) : c;
					}
					q.t = "T";
					hr = "h";
					i += 3;
				} else if (fmt.substring(i, 5).toUpperCase() === "AM/PM") {
					if (dt != null) {
						q.v = dt.H >= 12 ? "PM" : "AM";
					}
					q.t = "T";
					i += 5;
					hr = "h";
				} else if (fmt.substring(i, 5).toUpperCase() === "\u4E0A\u5348/\u4E0B\u5348") {
					if (dt != null) {
						q.v = dt.H >= 12 ? "\u4E0B\u5348" : "\u4E0A\u5348";
					}
					q.t = "T";
					i += 5;
					hr = "h";
				} else {
					q.t = "t";
					++i;
				}
				if (dt == null && q.t === "T") {
					return "";
				}
				out[out.length] = q;
				lst = c;
				break;
			}
			case "[":
				o = c;
				while (fmt.charAt(i++) !== "]" && i < fmt.length) {
					o += fmt.charAt(i);
				}
				if (o.slice(-1) !== "]") {
					throw new Error('unterminated "[" block: |' + o + "|");
				}
				if (o.match(SSF_abstime)) {
					if (dt == null) {
						dt = SSF_parse_date_code(v, opts);
						if (dt == null) {
							return "";
						}
					}
					out[out.length] = { t: "Z", v: o.toLowerCase() };
					lst = o.charAt(1);
				} else if (o.indexOf("$") > -1) {
					o = (o.match(/\$([^-[\]]*)/) || [])[1] || "$";
					if (!fmt_is_date(fmt)) {
						out[out.length] = { t: "t", v: o };
					}
				}
				break;
			case ".":
				if (dt != null) {
					o = c;
					while (++i < fmt.length && (c = fmt.charAt(i)) === "0") {
						o += c;
					}
					out[out.length] = { t: "s", v: o };
					break;
				}
			/* falls through */
			case "0":
			case "#":
				o = c;
				while (++i < fmt.length && "0#?.,E+-%".indexOf((c = fmt.charAt(i))) > -1) {
					o += c;
				}
				out[out.length] = { t: "n", v: o };
				break;
			case "?":
				o = c;
				while (fmt.charAt(++i) === c) {
					o += c;
				}
				out[out.length] = { t: c, v: o };
				lst = c;
				break;
			case "*":
				++i;
				if (fmt.charAt(i) === " " || fmt.charAt(i) === "*") {
					++i;
				}
				break;
			case "(":
			case ")":
				out[out.length] = { t: flen === 1 ? "t" : c, v: c };
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
				o = c;
				while (i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1) {
					o += fmt.charAt(i);
				}
				out[out.length] = { t: "D", v: o };
				break;
			case " ":
				out[out.length] = { t: c, v: c };
				++i;
				break;
			case "$":
				out[out.length] = { t: "t", v: "$" };
				++i;
				break;
			default:
				if (",$-+/():!^&'~{}<>=\u20ACacfijklopqrtuvwxzP".indexOf(c) === -1) {
					throw new Error("unrecognized character " + c + " in " + fmt);
				}
				out[out.length] = { t: "t", v: c };
				++i;
				break;
		}
	}

	/* Scan for date/time parts */
	let bt = 0;
	let ss0 = 0;
	let ssm: RegExpMatchArray | null;
	for (i = out.length - 1, lst = "t"; i >= 0; --i) {
		if (!out[i]) {
			continue;
		}
		switch (out[i]!.t) {
			case "h":
			case "H":
				out[i]!.t = hr;
				lst = "h";
				if (bt < 1) {
					bt = 1;
				}
				break;
			case "s":
				if ((ssm = out[i]!.v.match(/\.0+$/))) {
					ss0 = Math.max(ss0, ssm[0].length - 1);
					bt = 4;
				}
				if (bt < 3) {
					bt = 3;
				}
			/* falls through */
			case "d":
			case "y":
			case "e":
				lst = out[i]!.t;
				break;
			case "M":
				lst = out[i]!.t;
				if (bt < 2) {
					bt = 2;
				}
				break;
			case "m":
				if (lst === "s") {
					out[i]!.t = "M";
					if (bt < 2) {
						bt = 2;
					}
				}
				break;
			case "X":
				break;
			case "Z":
				if (bt < 1 && out[i]!.v.match(/[Hh]/)) {
					bt = 1;
				}
				if (bt < 2 && out[i]!.v.match(/[Mm]/)) {
					bt = 2;
				}
				if (bt < 3 && out[i]!.v.match(/[Ss]/)) {
					bt = 3;
				}
		}
	}

	/* time rounding */
	if (dt) {
		let _dt: SSFDateVal | null;
		switch (bt) {
			case 0:
				break;
			case 1:
			case 2:
			case 3:
				if (dt.u >= 0.5) {
					dt.u = 0;
					++dt.S;
				}
				if (dt.S >= 60) {
					dt.S = 0;
					++dt.M;
				}
				if (dt.M >= 60) {
					dt.M = 0;
					++dt.H;
				}
				if (dt.H >= 24) {
					dt.H = 0;
					++dt.D;
					_dt = SSF_parse_date_code(dt.D);
					if (_dt) {
						_dt.u = dt.u;
						_dt.S = dt.S;
						_dt.M = dt.M;
						_dt.H = dt.H;
						dt = _dt;
					}
				}
				break;
			case 4:
				switch (ss0) {
					case 1:
						dt.u = Math.round(dt.u * 10) / 10;
						break;
					case 2:
						dt.u = Math.round(dt.u * 100) / 100;
						break;
					case 3:
						dt.u = Math.round(dt.u * 1000) / 1000;
						break;
				}
				if (dt.u >= 1) {
					dt.u = 0;
					++dt.S;
				}
				if (dt.S >= 60) {
					dt.S = 0;
					++dt.M;
				}
				if (dt.M >= 60) {
					dt.M = 0;
					++dt.H;
				}
				if (dt.H >= 24) {
					dt.H = 0;
					++dt.D;
					_dt = SSF_parse_date_code(dt.D);
					if (_dt) {
						_dt.u = dt.u;
						_dt.S = dt.S;
						_dt.M = dt.M;
						_dt.H = dt.H;
						dt = _dt;
					}
				}
				break;
		}
	}

	/* replace fields */
	let nstr = "";
	let jj: number;
	for (i = 0; i < out.length; ++i) {
		if (!out[i]) {
			continue;
		}
		switch (out[i]!.t) {
			case "t":
			case "T":
			case " ":
			case "D":
				break;
			case "X":
				out[i]!.v = "";
				out[i]!.t = ";";
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
				out[i]!.v = SSF_write_date(out[i]!.t.charCodeAt(0), out[i]!.v, dt!, ss0);
				out[i]!.t = "t";
				break;
			case "n":
			case "?":
				jj = i + 1;
				while (
					out[jj] != null &&
					((c = out[jj]!.t) === "?" ||
						c === "D" ||
						((c === " " || c === "t") &&
							out[jj + 1] != null &&
							(out[jj + 1]!.t === "?" || (out[jj + 1]!.t === "t" && out[jj + 1]!.v === "/"))) ||
						(out[i]!.t === "(" && (c === " " || c === "n" || c === ")")) ||
						(c === "t" &&
							(out[jj]!.v === "/" ||
								(out[jj]!.v === " " && out[jj + 1] != null && out[jj + 1]!.t === "?"))))
				) {
					out[i]!.v += out[jj]!.v;
					out[jj] = { v: "", t: ";" };
					++jj;
				}
				nstr += out[i]!.v;
				i = jj - 1;
				break;
			case "G":
				out[i]!.t = "t";
				out[i]!.v = SSF_general(v, opts);
				break;
		}
	}

	let vv = "";
	let myv: number;
	let ostr: string;
	if (nstr.length > 0) {
		if (nstr.charCodeAt(0) === 40) {
			myv = v < 0 && nstr.charCodeAt(0) === 45 ? -v : v;
			ostr = write_num("n", nstr, myv);
		} else {
			myv = v < 0 && flen > 1 ? -v : v;
			ostr = write_num("n", nstr, myv);
			if (myv < 0 && out[0] && out[0].t === "t") {
				ostr = ostr.substring(1);
				out[0].v = "-" + out[0].v;
			}
		}
		jj = ostr.length - 1;
		let decpt = out.length;
		for (i = 0; i < out.length; ++i) {
			if (out[i] != null && out[i]!.t !== "t" && out[i]!.v.indexOf(".") > -1) {
				decpt = i;
				break;
			}
		}
		let lasti = out.length;
		if (decpt === out.length && ostr.indexOf("E") === -1) {
			for (i = out.length - 1; i >= 0; --i) {
				if (out[i] == null || "n?".indexOf(out[i]!.t) === -1) {
					continue;
				}
				if (jj >= out[i]!.v.length - 1) {
					jj -= out[i]!.v.length;
					out[i]!.v = ostr.substring(jj + 1, out[i]!.v.length);
				} else if (jj < 0) {
					out[i]!.v = "";
				} else {
					out[i]!.v = ostr.substring(0, jj + 1);
					jj = -1;
				}
				out[i]!.t = "t";
				lasti = i;
			}
			if (jj >= 0 && lasti < out.length) {
				out[lasti]!.v = ostr.substring(0, jj + 1) + out[lasti]!.v;
			}
		} else if (decpt !== out.length && ostr.indexOf("E") === -1) {
			jj = ostr.indexOf(".") - 1;
			for (i = decpt; i >= 0; --i) {
				if (out[i] == null || "n?".indexOf(out[i]!.t) === -1) {
					continue;
				}
				j = out[i]!.v.indexOf(".") > -1 && i === decpt ? out[i]!.v.indexOf(".") - 1 : out[i]!.v.length - 1;
				vv = out[i]!.v.substring(j + 1);
				for (; j >= 0; --j) {
					if (jj >= 0 && (out[i]!.v.charAt(j) === "0" || out[i]!.v.charAt(j) === "#")) {
						vv = ostr.charAt(jj--) + vv;
					}
				}
				out[i]!.v = vv;
				out[i]!.t = "t";
				lasti = i;
			}
			if (jj >= 0 && lasti < out.length) {
				out[lasti]!.v = ostr.substring(0, jj + 1) + out[lasti]!.v;
			}
			jj = ostr.indexOf(".") + 1;
			for (i = decpt; i < out.length; ++i) {
				if (out[i] == null || ("n?(".indexOf(out[i]!.t) === -1 && i !== decpt)) {
					continue;
				}
				j = out[i]!.v.indexOf(".") > -1 && i === decpt ? out[i]!.v.indexOf(".") + 1 : 0;
				vv = out[i]!.v.substring(0, j);
				for (; j < out[i]!.v.length; ++j) {
					if (jj < ostr.length) {
						vv += ostr.charAt(jj++);
					}
				}
				out[i]!.v = vv;
				out[i]!.t = "t";
				lasti = i;
			}
		}
	}
	for (i = 0; i < out.length; ++i) {
		if (out[i] != null && "n?".indexOf(out[i]!.t) > -1) {
			myv = flen > 1 && v < 0 && i > 0 && out[i - 1]!.v === "-" ? -v : v;
			out[i]!.v = write_num(out[i]!.t, out[i]!.v, myv);
			out[i]!.t = "t";
		}
	}
	let retval = "";
	for (i = 0; i !== out.length; ++i) {
		if (out[i] != null) {
			retval += out[i]!.v;
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

function choose_fmt(f: string, v: any): [number, string] {
	let fmt = SSF_split_fmt(f);
	const l = fmt.length;
	const lat = fmt[l - 1].indexOf("@");
	let ll = l;
	if (l < 4 && lat > -1) {
		--ll;
	}
	if (fmt.length > 4) {
		throw new Error("cannot find right format for |" + fmt.join("|") + "|");
	}
	if (typeof v !== "number") {
		return [4, fmt.length === 4 || lat > -1 ? fmt[fmt.length - 1] : "@"];
	}
	if (typeof v === "number" && !isFinite(v)) {
		v = 0;
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
	const ff = v > 0 ? fmt[0] : v < 0 ? fmt[1] : fmt[2];
	if (fmt[0].indexOf("[") === -1 && fmt[1].indexOf("[") === -1) {
		return [ll, ff];
	}
	if (fmt[0].match(/\[[=<>]/) != null || fmt[1].match(/\[[=<>]/) != null) {
		const m1 = fmt[0].match(cfregex2);
		const m2 = fmt[1].match(cfregex2);
		return chkcond(v, m1)
			? [ll, fmt[0]]
			: chkcond(v, m2)
				? [ll, fmt[1]]
				: [ll, fmt[m1 != null && m2 != null ? 2 : 1]];
	}
	return [ll, ff];
}

/** Format a value using an Excel number format string */
export function SSF_format(fmt: string | number, v: any, o?: any): string {
	if (o == null) {
		o = {};
	}
	let sfmt = "";
	switch (typeof fmt) {
		case "string":
			if (fmt === "m/d/yy" && o.dateNF) {
				sfmt = o.dateNF;
			} else {
				sfmt = fmt;
			}
			break;
		case "number":
			if (fmt === 14 && o.dateNF) {
				sfmt = o.dateNF;
			} else {
				sfmt = (o.table != null ? o.table : table_fmt)[fmt];
			}
			if (sfmt == null) {
				sfmt = (o.table && o.table[SSF_default_map[fmt]]) || table_fmt[SSF_default_map[fmt]];
			}
			if (sfmt == null) {
				sfmt = SSF_default_str[fmt] || "General";
			}
			break;
	}
	if (SSF_isgeneral(sfmt, 0)) {
		return SSF_general(v, o);
	}
	if (v instanceof Date) {
		v = datenum(v, o.date1904);
	}
	const f = choose_fmt(sfmt, v);
	if (SSF_isgeneral(f[1])) {
		return SSF_general(v, o);
	}
	if (v === true) {
		v = "TRUE";
	} else if (v === false) {
		v = "FALSE";
	} else if (v === "" || v == null) {
		return "";
	} else if (isNaN(v) && f[1].indexOf("0") > -1) {
		return "#NUM!";
	} else if (!isFinite(v) && f[1].indexOf("0") > -1) {
		return "#DIV/0!";
	}
	return eval_fmt(f[1], v, o, f[0]);
}
