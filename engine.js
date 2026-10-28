/* =============================================================================
   Excel Intensivo — mini motor de fórmulas (es-AR)
   Evalúa fórmulas estilo Excel sobre una grilla en memoria, para que los
   ejercicios sean interactivos y den el mismo resultado que daría Excel.

   Soporta: números, texto "..", VERDADERO/FALSO, referencias A1 / $A$1,
   rangos A1:B3, operadores + - * / ^ % & = <> < > <= >=, separador ; o ,
   y ~30 funciones frecuentes (SUMA, SI, BUSCARV, BUSCARX, SUMAR.SI.CONJUNTO…).

   API:  XL.calc("=B2*(1+$G$1)", { data: [ [fila1...], [fila2...], ... ] })
         XL.isErr(v) · XL.fmt(v, "money"|"pct"|"int") · XL.looseNum(a,b)
   La grilla: data[0] es la FILA 1 de Excel (los encabezados), data[1] la fila 2…
   Cada celda es un valor JS (número/texto) o un string que empieza con "=".
   ========================================================================== */
window.XL = (function () {
  "use strict";

  var ERR = function (c) { return { e: c }; };
  var isErr = function (v) { return v && typeof v === "object" && "e" in v; };

  function colToIdx(s) {
    s = s.toUpperCase(); var n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }

  function toNum(x) {
    if (typeof x === "number") return x;
    if (typeof x === "boolean") return x ? 1 : 0;
    if (x == null || x === "") return NaN;
    var s = String(x).trim().replace(/[\s$%]/g, "");
    if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, "").replace(",", ".");
    else if (/^-?\d+,\d+$/.test(s)) s = s.replace(",", ".");
    var n = Number(s);
    return s !== "" && !isNaN(n) ? n : NaN;
  }
  var isNum = function (v) {
    return typeof v === "number" || (typeof v === "string" && v.trim() !== "" && !isNaN(toNum(v)));
  };
  var num0 = function (v) { var n = toNum(v); return isNaN(n) ? 0 : n; };
  function round(x, d) { var f = Math.pow(10, d || 0); return Math.round((x + Number.EPSILON) * f) / f; }
  function cmp(a, b) { return a < b ? -1 : a > b ? 1 : 0; }

  function asText(v) {
    if (isErr(v)) return v.e;
    if (v === true) return "VERDADERO";
    if (v === false) return "FALSO";
    if (v == null) return "";
    return String(v);
  }
  function truthy(v) {
    if (isErr(v)) return false;
    if (typeof v === "boolean") return v;
    if (typeof v === "number") return v !== 0;
    if (v === "" || v == null) return false;
    var n = toNum(v);
    if (!isNaN(n)) return n !== 0;
    return String(v).toLowerCase() !== "falso";
  }
  function looseEq(a, b) {
    if (isErr(a) || isErr(b)) return false;
    if (isNum(a) && isNum(b)) return toNum(a) === toNum(b);
    return String(asText(a)).trim().toLowerCase() === String(asText(b)).trim().toLowerCase();
  }
  function compareVals(op, a, b) {
    if (a === "" && isNum(b)) a = 0;
    if (b === "" && isNum(a)) b = 0;
    var r;
    if (isNum(a) && isNum(b)) r = cmp(toNum(a), toNum(b));
    else r = cmp(String(asText(a)).toLowerCase(), String(asText(b)).toLowerCase());
    return op === "=" ? r === 0 : op === "<>" ? r !== 0 : op === "<" ? r < 0 :
      op === ">" ? r > 0 : op === "<=" ? r <= 0 : r >= 0;
  }

  /* ---------------- tokenizer ---------------- */
  function lex(src) {
    var T = [], i = 0;
    var R = [
      ["ws", /^\s+/],
      ["str", /^"([^"]*)"/],
      ["num", /^\d+(?:[.,]\d+)?/],
      ["ref", /^\$?[A-Za-z]{1,3}\$?\d+(?![A-Za-z0-9_])/],
      ["name", /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ_][\wÁÉÍÓÚÜÑáéíóúüñ.]*/],
      ["op", /^(<=|>=|<>|[-+*/^&=<>():;,%])/]
    ];
    outer: while (i < src.length) {
      var s = src.slice(i);
      for (var j = 0; j < R.length; j++) {
        var m = s.match(R[j][1]);
        if (m) {
          if (R[j][0] !== "ws")
            T.push({ k: R[j][0], v: R[j][0] === "str" ? m[1] : R[j][0] === "num" ? parseFloat(m[0].replace(",", ".")) : m[0] });
          i += m[0].length;
          continue outer;
        }
      }
      throw ERR("#¡ERROR!");
    }
    T.push({ k: "end" });
    return T;
  }

  /* ---------------- parser → AST ---------------- */
  function parse(src) {
    var T = lex(src), p = 0;
    var pk = function () { return T[p]; };
    var nx = function () { return T[p++]; };
    var isOp = function (v) { return pk().k === "op" && pk().v === v; };
    var eat = function (v) { if (isOp(v)) { p++; return; } throw ERR("#¡ERROR!"); };

    function expr() { return compare(); }
    function compare() {
      var a = concat();
      while (pk().k === "op" && ["=", "<>", "<", ">", "<=", ">="].indexOf(pk().v) >= 0)
        a = { k: "bin", op: nx().v, a: a, b: concat() };
      return a;
    }
    function concat() { var a = add(); while (isOp("&")) { nx(); a = { k: "bin", op: "&", a: a, b: add() }; } return a; }
    function add() {
      var a = mul();
      while (pk().k === "op" && (pk().v === "+" || pk().v === "-")) a = { k: "bin", op: nx().v, a: a, b: mul() };
      return a;
    }
    function mul() {
      var a = un();
      while (pk().k === "op" && (pk().v === "*" || pk().v === "/")) a = { k: "bin", op: nx().v, a: a, b: un() };
      return a;
    }
    function un() {
      if (pk().k === "op" && (pk().v === "-" || pk().v === "+")) return { k: "un", op: nx().v, x: un() };
      return pow();
    }
    function pow() { var a = post(); if (isOp("^")) { nx(); return { k: "bin", op: "^", a: a, b: un() }; } return a; }
    function post() { var a = prim(); while (isOp("%")) { nx(); a = { k: "un", op: "%", x: a }; } return a; }
    function prim() {
      var t = pk();
      if (t.k === "num") { nx(); return { k: "num", v: t.v }; }
      if (t.k === "str") { nx(); return { k: "str", v: t.v }; }
      if (isOp("(")) { nx(); var e = expr(); eat(")"); return e; }
      if (t.k === "ref") {
        nx();
        if (isOp(":")) { nx(); var t2 = nx(); if (t2.k !== "ref") throw ERR("#¡ERROR!"); return { k: "range", a: t.v, b: t2.v }; }
        return { k: "cell", ref: t.v };
      }
      if (t.k === "name") {
        var nm = nx().v;
        if (isOp("(")) {
          nx();
          var args = [];
          if (!isOp(")")) {
            args.push(arg());
            while (pk().k === "op" && (pk().v === ";" || pk().v === ",")) { nx(); args.push(arg()); }
          }
          eat(")");
          return { k: "call", name: nm.toUpperCase(), args: args };
        }
        if (/^(verdadero|true)$/i.test(nm)) return { k: "bool", v: true };
        if (/^(falso|false)$/i.test(nm)) return { k: "bool", v: false };
        return { k: "nameerr" };
      }
      throw ERR("#¡ERROR!");
    }
    function arg() {
      if (pk().k === "op" && (pk().v === ";" || pk().v === "," || pk().v === ")")) return { k: "empty" };
      return expr();
    }
    var ast = expr();
    if (pk().k !== "end") throw ERR("#¡ERROR!");
    return ast;
  }

  /* ---------------- evaluador ---------------- */
  function makeEval(sheet) {
    var data = (sheet && sheet.data) || [[]];
    var cache = {}, busy = {};

    function parseRef(ref) {
      var m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
      return { c: colToIdx(m[1]), r: parseInt(m[2], 10) - 1 };
    }
    function cellRaw(c, r) {
      if (r < 0 || r >= data.length) return "";
      var row = data[r] || [];
      return row[c] == null ? "" : row[c];
    }
    function cellVal(c, r) {
      var key = c + ":" + r;
      if (key in cache) return cache[key];
      if (busy[key]) return ERR("#¡REF!");
      busy[key] = 1;
      var raw = cellRaw(c, r), out;
      if (raw === "" || raw == null) out = "";
      else if (typeof raw === "number") out = raw;
      else {
        var s = String(raw);
        if (s.charAt(0) === "=") out = evalStr(s.slice(1));
        else if (/^(verdadero|true)$/i.test(s.trim())) out = true;
        else if (/^(falso|false)$/i.test(s.trim())) out = false;
        else if (s.trim() !== "" && !isNaN(toNum(s))) out = toNum(s);
        else out = s;
      }
      delete busy[key];
      cache[key] = out;
      return out;
    }
    function rangeVal(a, b) {
      var A = parseRef(a), B = parseRef(b);
      var c0 = Math.min(A.c, B.c), c1 = Math.max(A.c, B.c), r0 = Math.min(A.r, B.r), r1 = Math.max(A.r, B.r);
      var rows = [];
      for (var r = r0; r <= r1; r++) {
        var line = [];
        for (var c = c0; c <= c1; c++) line.push(cellVal(c, r));
        rows.push(line);
      }
      return { rng: true, rows: rows };
    }
    function flat(v) {
      if (v && v.rng) return Array.prototype.concat.apply([], v.rows);
      return Array.isArray(v) ? v : [v];
    }

    function ev(n) {
      switch (n.k) {
        case "num": case "str": case "bool": return n.v;
        case "empty": return "";
        case "nameerr": return ERR("#¿NOMBRE?");
        case "cell": { var p = parseRef(n.ref); return cellVal(p.c, p.r); }
        case "range": return rangeVal(n.a, n.b);
        case "un": {
          var x = ev(n.x); if (isErr(x)) return x;
          var v = x === "" ? 0 : toNum(x); if (isNaN(v)) return ERR("#¡VALOR!");
          if (n.op === "%") return v / 100;
          return n.op === "-" ? -v : v;
        }
        case "bin": return bin(n.op, n.a, n.b);
        case "call": return call(n.name, n.args);
      }
      return ERR("#¡ERROR!");
    }
    function bin(op, an, bn) {
      var a = ev(an); if (isErr(a)) return a;
      var b = ev(bn); if (isErr(b)) return b;
      if (op === "&") return asText(a) + asText(b);
      if (["=", "<>", "<", ">", "<=", ">="].indexOf(op) >= 0) return compareVals(op, a, b);
      var x = a === "" ? 0 : toNum(a), y = b === "" ? 0 : toNum(b);
      if (isNaN(x) || isNaN(y)) return ERR("#¡VALOR!");
      if (op === "+") return x + y;
      if (op === "-") return x - y;
      if (op === "*") return x * y;
      if (op === "/") return y === 0 ? ERR("#¡DIV/0!") : x / y;
      if (op === "^") return Math.pow(x, y);
      return ERR("#¡ERROR!");
    }

    function pairs(vs, start) {
      var p = [];
      for (var i = start; i + 1 < vs.length; i += 2) p.push([flat(vs[i]), vs[i + 1]]);
      return p;
    }
    function crit(cell, c) {
      if (isErr(cell)) return false;
      if (typeof c === "number") return isNum(cell) && toNum(cell) === c;
      var s = String(c);
      var m = /^(<=|>=|<>|=|<|>)([\s\S]*)$/.exec(s);
      if (m) {
        var op = m[1], rhs = m[2];
        if (isNum(cell) && rhs !== "" && !isNaN(toNum(rhs))) return compareVals(op, toNum(cell), toNum(rhs));
        return compareVals(op, String(asText(cell)).toLowerCase(), rhs.toLowerCase());
      }
      if (isNum(cell) && s.trim() !== "" && !isNaN(toNum(s))) return toNum(cell) === toNum(s);
      var cs = String(asText(cell)).toLowerCase(), ps = s.toLowerCase();
      if (ps.indexOf("*") >= 0 || ps.indexOf("?") >= 0) {
        var re = new RegExp("^" + ps.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$");
        return re.test(cs);
      }
      return cs === ps;
    }

    function call(name, args) {
      var lazy = { "SI": 1, "SI.ERROR": 1, "SIERROR": 1, "ESERROR": 1, "Y": 1, "O": 1 };
      var vs = [], i;
      if (lazy[name]) {
        for (i = 0; i < args.length; i++) {
          try { vs.push(ev(args[i])); } catch (e) { vs.push(isErr(e) ? e : ERR("#¡ERROR!")); }
        }
      } else {
        for (i = 0; i < args.length; i++) {
          var vv;
          try { vv = ev(args[i]); } catch (e) { vv = isErr(e) ? e : ERR("#¡ERROR!"); }
          if (isErr(vv)) return vv;
          vs.push(vv);
        }
      }
      var S = function () { return Array.prototype.concat.apply([], vs.map(flat)); };
      var nums = function () { return S().filter(isNum).map(toNum); };
      var a, n, k;

      switch (name) {
        case "SUMA": return S().reduce(function (s, v) { return s + num0(v); }, 0);
        case "PROMEDIO": a = nums(); return a.length ? a.reduce(function (x, y) { return x + y; }, 0) / a.length : ERR("#¡DIV/0!");
        case "MAX": a = nums(); return a.length ? Math.max.apply(null, a) : 0;
        case "MIN": a = nums(); return a.length ? Math.min.apply(null, a) : 0;
        case "CONTAR": return S().filter(isNum).length;
        case "CONTARA": return S().filter(function (v) { return !(v === "" || v == null); }).length;
        case "CONTAR.BLANCO": return S().filter(function (v) { return v === "" || v == null; }).length;
        case "REDONDEAR": return round(toNum(vs[0]), toNum(vs[1] || 0));
        case "REDONDEAR.MENOS": return Math.floor(toNum(vs[0]) * Math.pow(10, toNum(vs[1] || 0))) / Math.pow(10, toNum(vs[1] || 0));
        case "ABS": return Math.abs(toNum(vs[0]));
        case "ENTERO": return Math.floor(toNum(vs[0]));
        case "SUMAPRODUCTO": {
          var arrs = vs.map(flat), len = arrs[0].length, s = 0;
          for (i = 0; i < len; i++) { var pr = 1; for (k = 0; k < arrs.length; k++) pr *= num0(arrs[k][i]); s += pr; }
          return s;
        }
        case "SI": {
          if (isErr(vs[0])) return vs[0];
          if (truthy(vs[0])) return vs.length > 1 ? vs[1] : true;
          return vs.length > 2 ? vs[2] : false;
        }
        case "SI.ERROR": case "SIERROR": return isErr(vs[0]) ? (vs.length > 1 ? vs[1] : "") : vs[0];
        case "ESERROR": return isErr(vs[0]);
        case "Y": return S().every(truthy);
        case "O": return S().some(truthy);
        case "NO": return !truthy(vs[0]);
        case "BUSCARV": {
          var rng = vs[1], col = Math.round(toNum(vs[2]));
          if (!rng || !rng.rng) return ERR("#¡VALOR!");
          for (i = 0; i < rng.rows.length; i++)
            if (looseEq(rng.rows[i][0], vs[0])) return rng.rows[i][col - 1] === undefined ? ERR("#¡REF!") : rng.rows[i][col - 1];
          return ERR("#N/A");
        }
        case "BUSCARX": {
          var lk = flat(vs[1]), rt = flat(vs[2]);
          for (i = 0; i < lk.length; i++) if (looseEq(lk[i], vs[0])) return rt[i];
          return vs.length > 3 ? vs[3] : ERR("#N/A");
        }
        case "COINCIDIR": {
          var arr = flat(vs[1]);
          for (i = 0; i < arr.length; i++) if (looseEq(arr[i], vs[0])) return i + 1;
          return ERR("#N/A");
        }
        case "INDICE": {
          var g = vs[0]; if (!g || !g.rng) return ERR("#¡VALOR!");
          if (vs.length >= 3 && vs[2] !== "") {
            var rr = Math.round(toNum(vs[1])) - 1, cc = Math.round(toNum(vs[2])) - 1;
            return g.rows[rr] && g.rows[rr][cc] !== undefined ? g.rows[rr][cc] : ERR("#¡REF!");
          }
          var lin = Array.prototype.concat.apply([], g.rows);
          k = Math.round(toNum(vs[1])) - 1;
          return lin[k] !== undefined ? lin[k] : ERR("#¡REF!");
        }
        case "SUMAR.SI": {
          var r0 = flat(vs[0]), sr = vs.length > 2 ? flat(vs[2]) : r0, sm = 0;
          for (i = 0; i < r0.length; i++) if (crit(r0[i], vs[1])) sm += num0(sr[i]);
          return sm;
        }
        case "CONTAR.SI": return flat(vs[0]).filter(function (x) { return crit(x, vs[1]); }).length;
        case "PROMEDIO.SI": {
          var rp = flat(vs[0]), sp = vs.length > 2 ? flat(vs[2]) : rp, tot = 0, cn = 0;
          for (i = 0; i < rp.length; i++) if (crit(rp[i], vs[1])) { tot += num0(sp[i]); cn++; }
          return cn ? tot / cn : ERR("#¡DIV/0!");
        }
        case "SUMAR.SI.CONJUNTO": {
          var srr = flat(vs[0]), P = pairs(vs, 1), acc = 0;
          for (i = 0; i < srr.length; i++)
            if (P.every(function (pp) { return crit(pp[0][i], pp[1]); })) acc += num0(srr[i]);
          return acc;
        }
        case "CONTAR.SI.CONJUNTO": {
          var P2 = pairs(vs, 0), nn = P2[0][0].length, cnt = 0;
          for (i = 0; i < nn; i++) if (P2.every(function (pp) { return crit(pp[0][i], pp[1]); })) cnt++;
          return cnt;
        }
        case "PROMEDIO.SI.CONJUNTO": {
          var srp = flat(vs[0]), P3 = pairs(vs, 1), t2 = 0, c2 = 0;
          for (i = 0; i < srp.length; i++)
            if (P3.every(function (pp) { return crit(pp[0][i], pp[1]); })) { t2 += num0(srp[i]); c2++; }
          return c2 ? t2 / c2 : ERR("#¡DIV/0!");
        }
        case "IZQUIERDA": return String(asText(vs[0])).slice(0, vs[1] === undefined ? 1 : Math.round(toNum(vs[1])));
        case "DERECHA": {
          var ss = String(asText(vs[0])), kk = vs[1] === undefined ? 1 : Math.round(toNum(vs[1]));
          return ss.slice(ss.length - kk);
        }
        case "EXTRAE": return String(asText(vs[0])).substr(Math.round(toNum(vs[1])) - 1, Math.round(toNum(vs[2])));
        case "LARGO": return String(asText(vs[0])).length;
        case "ESPACIOS": return String(asText(vs[0])).trim().replace(/\s+/g, " ");
        case "MAYUSC": return String(asText(vs[0])).toUpperCase();
        case "MINUSC": return String(asText(vs[0])).toLowerCase();
        case "CONCAT": case "CONCATENAR": return S().map(asText).join("");
        case "TEXTO": return String(asText(vs[0]));
        case "HOY": case "AHORA": return 0;
        default: return ERR("#¿NOMBRE?");
      }
    }

    function evalStr(s) {
      try { return ev(parse(s)); }
      catch (e) { return isErr(e) ? e : ERR("#¡ERROR!"); }
    }
    return { evalStr: evalStr };
  }

  /* ---------------- API pública ---------------- */
  function calc(formula, sheet) {
    var f = String(formula == null ? "" : formula).replace(/^\s*=/, "");
    if (!f.trim()) return ERR("#¡ERROR!");
    return makeEval(sheet || { data: [[]] }).evalStr(f);
  }

  function fmt(v, kind) {
    if (isErr(v)) return v.e;
    if (v === true) return "VERDADERO";
    if (v === false) return "FALSO";
    if (v === "" || v == null) return "";
    if (typeof v === "number") {
      if (kind === "pct") return (Math.round(v * 1000) / 10).toLocaleString("es-AR") + "%";
      var dec = kind === "int" ? 0 : (Math.abs(v % 1) > 1e-9 ? 2 : 0);
      var s = v.toLocaleString("es-AR", { minimumFractionDigits: dec, maximumFractionDigits: dec });
      return kind === "money" ? "$ " + s : s;
    }
    return String(v);
  }

  // comparación tolerante de números (acepta %, $, miles, y ~0,5% de error)
  function looseNum(got, exp, tol) {
    var g = toNum(got), e = toNum(exp);
    if (isNaN(g) || isNaN(e)) return String(got).trim().toLowerCase() === String(exp).trim().toLowerCase();
    var lim = tol != null ? tol : Math.max(0.01, Math.abs(e) * 0.006);
    if (Math.abs(g - e) <= lim) return true;
    // aceptar que hayan escrito 10,3 en vez de 0,103 (o viceversa)
    if (Math.abs(g / 100 - e) <= lim || Math.abs(g - e / 100) <= lim) return true;
    return false;
  }

  function normFormula(s) {
    return String(s || "").toUpperCase().replace(/\s+/g, "").replace(/,/g, ";").replace(/^=/, "");
  }

  return { calc: calc, isErr: isErr, ERR: ERR, fmt: fmt, looseNum: looseNum, normFormula: normFormula, toNum: toNum };
})();
