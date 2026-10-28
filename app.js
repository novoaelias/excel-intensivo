/* =============================================================================
   Excel Intensivo — lógica de la app (SPA, sin frameworks)
   Router por hash. Progreso persistente en localStorage.
   Ejercicios interactivos evaluados con el mini motor de fórmulas (engine.js).
   ========================================================================== */
(function () {
  "use strict";

  var C = window.COURSE;
  var XL = window.XL;
  var STORE_KEY = "excel-intensivo:v1";

  /* ---------- persistencia ---------- */
  function load() {
    try {
      var raw = localStorage.getItem(STORE_KEY);
      if (raw) {
        var o = JSON.parse(raw);
        o.modules = o.modules || {};
        o.quiz = o.quiz || {};
        o.interview = o.interview || {};
        o.ex = o.ex || {};
        return o;
      }
    } catch (e) {}
    return { modules: {}, quiz: {}, interview: {}, ex: {} };
  }
  function save() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); } catch (e) {}
  }
  var state = load();

  function modState(id) {
    if (!state.modules[id]) state.modules[id] = { theory: false, chk: {} };
    var m = state.modules[id];
    if (!m.chk) m.chk = {};
    return m;
  }
  function exId(id) { return !!state.ex[id]; }
  function setEx(id, v) { if (v) state.ex[id] = 1; else delete state.ex[id]; save(); }

  /* ---------- helpers de contenido ---------- */
  function lessonChecks(mod) {
    var out = [];
    mod.lessons.forEach(function (l) { (l.checks || []).forEach(function (c) { out.push(c); }); });
    return out;
  }

  /* ---------- cálculo de progreso ---------- */
  function moduleStats(mod) {
    var ms = modState(mod.id);
    var checks = lessonChecks(mod);
    var exs = mod.exercises || [];
    var nChk = mod.checklist.length;
    var i, chkDone = 0;
    for (i = 0; i < nChk; i++) if (ms.chk[i]) chkDone++;
    var checkDone = checks.filter(function (c) { return exId(c.id); }).length;
    var exDone = exs.filter(function (e) { return exId(e.id); }).length;
    var total = 1 + checks.length + exs.length + nChk;
    var done = (ms.theory ? 1 : 0) + checkDone + exDone + chkDone;
    var frac = total ? done / total : 0;
    var status = frac >= 1 ? "done" : done > 0 ? "progress" : "pending";
    return {
      frac: frac, status: status,
      exDone: exDone, nEx: exs.length,
      checkDone: checkDone, nCheck: checks.length,
      chkDone: chkDone, nChk: nChk, theory: ms.theory
    };
  }
  function overall() {
    var sum = 0, completed = 0;
    C.modules.forEach(function (m) {
      var s = moduleStats(m);
      sum += s.frac;
      if (s.status === "done") completed++;
    });
    return { pct: Math.round((sum / C.modules.length) * 100), completed: completed, total: C.modules.length };
  }
  function currentModule() {
    var inProgress = null, firstPending = null;
    for (var i = 0; i < C.modules.length; i++) {
      var s = moduleStats(C.modules[i]);
      if (s.status === "progress" && !inProgress) inProgress = C.modules[i];
      if (s.status === "pending" && !firstPending) firstPending = C.modules[i];
    }
    return inProgress || firstPending || C.modules[C.modules.length - 1];
  }

  /* ---------- helpers DOM ---------- */
  function el(html) {
    var d = document.createElement("div");
    d.innerHTML = html.trim();
    return d.firstChild;
  }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }
  var STATUS_LABEL = { done: "Completo", progress: "En curso", pending: "Pendiente" };

  var toastTimer;
  function toast(msg) {
    var t = document.getElementById("toast");
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove("show"); }, 1800);
  }
  function shuffle(a) {
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var t = a[i]; a[i] = a[j]; a[j] = t;
    }
    return a;
  }

  /* ---------- grillas ---------- */
  var COLS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];

  // grilla de teoría (display): { caption, cols, rows, highlight:{row,col}, formula:{cell,text,result} }
  function sheetHTML(s) {
    if (!s) return "";
    var ncol = s.cols.length, head = '<tr><th></th>', c, k, j;
    for (c = 0; c < ncol; c++) head += "<th>" + COLS[c] + "</th>";
    head += "</tr>";
    var rows = '<tr class="hdr"><td class="rownum">1</td>';
    for (k = 0; k < ncol; k++) rows += "<td>" + esc(s.cols[k]) + "</td>";
    rows += "</tr>";
    s.rows.forEach(function (r, ri) {
      var rn = ri + 2;
      rows += '<tr><td class="rownum">' + rn + "</td>";
      for (j = 0; j < ncol; j++) {
        var hl = s.highlight && s.highlight.row === rn && s.highlight.col === j + 1 ? ' class="hl"' : "";
        rows += "<td" + hl + ">" + esc(r[j] != null ? r[j] : "") + "</td>";
      }
      rows += "</tr>";
    });
    var fbar = "";
    if (s.formula) {
      fbar = '<div class="fbar"><span class="cell">' + esc(s.formula.cell) + "</span>" +
        '<span class="eq">' + esc(s.formula.text) + "</span>" +
        (s.formula.result != null ? '<span>&rarr;</span><span class="res">' + esc(s.formula.result) + "</span>" : "") +
        "</div>";
    }
    return '<div class="sheet"><div class="cap">' + esc(s.caption) + "</div>" +
      '<div class="grid-wrap"><table><thead>' + head + "</thead><tbody>" + rows + "</tbody></table></div>" + fbar + "</div>";
  }

  // grilla de ejercicio: { cols:[...], rows:[[raw,...]], money:[idx], pct:[idx], caption }
  function fmtCell(v, gi, grid) {
    if (typeof v === "string" && v.charAt(0) === "=") {
      var cv = XL.calc(v, { data: gridData(grid) });
      v = XL.isErr(cv) ? cv.e : cv;
    }
    if (typeof v !== "number") return esc(v);
    if (grid.money && grid.money.indexOf(gi) >= 0) return esc(XL.fmt(v, "money"));
    if (grid.pct && grid.pct.indexOf(gi) >= 0) return esc(XL.fmt(v, "pct"));
    return esc(XL.fmt(v));
  }
  function gridHTML(grid, hlCell) {
    if (!grid) return "";
    var hl = hlCell ? parseCellRef(hlCell) : null;
    var ncol = grid.cols.length, head = '<tr><th></th>', c, k, j;
    for (c = 0; c < ncol; c++) head += "<th>" + COLS[c] + "</th>";
    head += "</tr>";
    var body = '<tr class="hdr"><td class="rownum">1</td>';
    for (k = 0; k < ncol; k++) body += "<td>" + esc(grid.cols[k]) + "</td>";
    body += "</tr>";
    grid.rows.forEach(function (r, ri) {
      var rn = ri + 2;
      body += '<tr><td class="rownum">' + rn + "</td>";
      for (j = 0; j < ncol; j++) {
        var isHl = hl && hl.c === j && hl.r === ri + 1;
        body += "<td" + (isHl ? ' class="hl"' : "") + ">" + fmtCell(r[j] != null ? r[j] : "", j, grid) + "</td>";
      }
      body += "</tr>";
    });
    return '<div class="sheet">' + (grid.caption ? '<div class="cap">' + esc(grid.caption) + "</div>" : "") +
      '<div class="grid-wrap"><table><thead>' + head + "</thead><tbody>" + body + "</tbody></table></div></div>";
  }
  function parseCellRef(ref) {
    var m = /^\$?([A-Za-z])\$?(\d+)$/.exec(ref);
    if (!m) return null;
    return { c: m[1].toUpperCase().charCodeAt(0) - 65, r: parseInt(m[2], 10) - 1 };
  }
  function gridData(grid) { return [grid.cols].concat(grid.rows.map(function (r) { return r.slice(); })); }

  /* ==========================================================================
     MOTOR DE EJERCICIOS INTERACTIVOS
     ex = { id, type, q, explain, ... }
     ctx = { kind:"Ejercicio"|"Práctica", n, onSolve }
     ========================================================================== */
  var EX = {};

  EX.render = function (ex, ctx) {
    var solved = exId(ex.id);
    var wrap = el('<div class="exercise' + (solved ? " solved" : "") + '" id="ex-' + esc(ex.id) + '"></div>');
    var lbl = ctx.kind + (ctx.n ? " " + ctx.n : "");
    wrap.appendChild(el('<div class="lbl">' + esc(lbl) + (solved ? ' <span class="ok-tick">✓ resuelto</span>' : "") + "</div>"));
    wrap.appendChild(el('<div class="q">' + esc(ex.q) + "</div>"));

    function done() {
      if (!exId(ex.id)) {
        setEx(ex.id, true);
        toast("¡Bien! Ejercicio resuelto");
      }
      wrap.classList.add("solved");
      if (ctx.onSolve) ctx.onSolve();
    }
    function feedback(host, ok, msg) {
      var fb = host.querySelector(".ex-fb") || host.appendChild(el('<div class="ex-fb"></div>'));
      fb.className = "ex-fb " + (ok ? "ok" : "no");
      fb.innerHTML = msg;
    }
    function revealExplain(host) {
      if (ex.explain && !host.querySelector(".ex-explain"))
        host.appendChild(el('<div class="ex-explain">' + ex.explain + "</div>"));
      if (ex.model && !host.querySelector(".ex-model"))
        host.appendChild(el('<div class="ex-model"><span>Fórmula modelo:</span> <code>' + esc(ex.model) + "</code></div>"));
    }

    var body = el('<div class="ex-body"></div>');
    wrap.appendChild(body);

    (EX["t_" + ex.type] || EX.t_unknown)(ex, body, { done: done, feedback: feedback, revealExplain: revealExplain, solved: solved });

    if (solved) revealExplain(body);
    return wrap;
  };

  EX.t_unknown = function (ex, body) {
    body.appendChild(el('<div class="ex-fb no">Tipo de ejercicio no reconocido.</div>'));
  };

  /* ---- opción múltiple ---- */
  EX.t_mc = function (ex, body, h) {
    var opts = el('<div class="quiz-opts"></div>');
    var locked = h.solved;
    ex.options.forEach(function (t, i) {
      var b = el('<button class="opt"><span class="mk">' + String.fromCharCode(65 + i) + '</span><span>' + esc(t) + "</span></button>");
      if (locked && i === ex.answer) { b.classList.add("correct"); b.querySelector(".mk").textContent = "✓"; }
      b.onclick = function () {
        if (locked) return;
        if (i === ex.answer) {
          locked = true;
          b.classList.add("correct"); b.querySelector(".mk").textContent = "✓";
          Array.prototype.forEach.call(opts.children, function (x) { x.disabled = true; });
          h.revealExplain(body); h.done();
        } else {
          b.classList.add("wrong"); b.querySelector(".mk").textContent = "✕"; b.disabled = true;
          h.feedback(body, false, "Esa no. Probá otra.");
        }
      };
      opts.appendChild(b);
    });
    body.appendChild(opts);
  };

  /* ---- varias correctas ---- */
  EX.t_multi = function (ex, body, h) {
    var box = el('<div class="card" style="padding:12px 14px"></div>');
    var picks = {};
    ex.options.forEach(function (t, i) {
      var row = el('<label class="rowcheck"><input type="checkbox"> ' + esc(t) + "</label>");
      if (h.solved && ex.answer.indexOf(i) >= 0) { row.querySelector("input").checked = true; }
      row.querySelector("input").onchange = function (e) { picks[i] = e.target.checked; };
      box.appendChild(row);
    });
    body.appendChild(box);
    if (h.solved) return;
    var btn = el('<button class="btn" style="margin-top:10px">Comprobar</button>');
    btn.onclick = function () {
      var sel = Object.keys(picks).filter(function (k) { return picks[k]; }).map(Number).sort();
      var ans = ex.answer.slice().sort();
      if (JSON.stringify(sel) === JSON.stringify(ans)) {
        btn.remove(); h.revealExplain(body); h.done();
      } else {
        h.feedback(body, false, "Todavía no. Fijate de nuevo qué opciones aplican.");
      }
    };
    body.appendChild(btn);
  };

  /* ---- ordenar pasos ---- */
  EX.t_order = function (ex, body, h) {
    var cur = h.solved ? ex.steps.slice() : shuffle(ex.steps.slice());
    if (!h.solved && JSON.stringify(cur) === JSON.stringify(ex.steps)) cur.push(cur.shift());
    var list = el('<div class="order-list"></div>');
    function paint() {
      list.innerHTML = "";
      cur.forEach(function (step, i) {
        var row = el('<div class="order-row"><span class="order-num">' + (i + 1) + '</span><span class="order-txt">' + esc(step) + '</span></div>');
        if (!h.solved) {
          var ctrls = el('<span class="order-ctrls"></span>');
          var up = el('<button aria-label="subir">↑</button>');
          var dn = el('<button aria-label="bajar">↓</button>');
          up.disabled = i === 0; dn.disabled = i === cur.length - 1;
          up.onclick = function () { var t = cur[i - 1]; cur[i - 1] = cur[i]; cur[i] = t; paint(); };
          dn.onclick = function () { var t = cur[i + 1]; cur[i + 1] = cur[i]; cur[i] = t; paint(); };
          ctrls.appendChild(up); ctrls.appendChild(dn);
          row.appendChild(ctrls);
        }
        list.appendChild(row);
      });
    }
    paint();
    body.appendChild(list);
    if (h.solved) return;
    var btn = el('<button class="btn" style="margin-top:10px">Comprobar orden</button>');
    btn.onclick = function () {
      if (JSON.stringify(cur) === JSON.stringify(ex.steps)) {
        btn.remove();
        Array.prototype.forEach.call(list.children, function (r) { r.classList.add("ok"); });
        h.revealExplain(body); h.done();
      } else {
        Array.prototype.forEach.call(list.children, function (r, i) {
          r.classList.toggle("bad", cur[i] !== ex.steps[i]);
        });
        h.feedback(body, false, "Ese orden no es. Las filas en rojo están fuera de lugar.");
      }
    };
    body.appendChild(btn);
  };

  /* ---- clasificar en categorías ---- */
  EX.t_classify = function (ex, body, h) {
    var assign = {};
    var list = el('<div class="classify-list"></div>');
    ex.items.forEach(function (it, i) {
      if (h.solved) assign[i] = it.c;
      var row = el('<div class="classify-row"><span class="classify-txt"><code>' + esc(it.t) + "</code></span></div>");
      var cats = el('<span class="classify-cats"></span>');
      ex.cats.forEach(function (cat, ci) {
        var b = el('<button>' + esc(cat) + "</button>");
        if (h.solved && it.c === ci) b.classList.add("sel");
        b.onclick = function () {
          if (h.solved) return;
          assign[i] = ci;
          Array.prototype.forEach.call(cats.children, function (x, xi) { x.classList.toggle("sel", xi === ci); });
        };
        cats.appendChild(b);
      });
      row.appendChild(cats);
      list.appendChild(row);
    });
    body.appendChild(list);
    if (h.solved) return;
    var btn = el('<button class="btn" style="margin-top:10px">Comprobar</button>');
    btn.onclick = function () {
      var allRight = true;
      ex.items.forEach(function (it, i) {
        var row = list.children[i];
        var ok = assign[i] === it.c;
        row.classList.toggle("ok", ok);
        row.classList.toggle("bad", assign[i] != null && !ok);
        if (!ok) allRight = false;
      });
      if (allRight && Object.keys(assign).length === ex.items.length) {
        btn.remove(); h.revealExplain(body); h.done();
      } else {
        h.feedback(body, false, "Todavía hay items mal clasificados (en rojo) o sin clasificar.");
      }
    };
    body.appendChild(btn);
  };

  /* ---- armar fórmula tocando fichas ---- */
  EX.t_build = function (ex, body, h) {
    var answer = h.solved ? ex.answer.slice() : [];
    var line = el('<div class="build-line"></div>');
    var bank = el('<div class="build-bank"></div>');
    function paint() {
      line.innerHTML = answer.length ? "" : '<span class="build-ph">Tocá las fichas para armar la fórmula…</span>';
      answer.forEach(function (tok, i) {
        var chip = el('<button class="build-chip">' + esc(tok) + "</button>");
        if (!h.solved) chip.onclick = function () { answer.splice(i, 1); paint(); };
        line.appendChild(chip);
      });
    }
    bank.innerHTML = "";
    ex.bank.forEach(function (tok) {
      var b = el('<button class="build-tok">' + esc(tok) + "</button>");
      if (h.solved) b.disabled = true;
      b.onclick = function () { answer.push(tok); paint(); };
      bank.appendChild(b);
    });
    paint();
    body.appendChild(line);
    body.appendChild(bank);
    if (h.solved) return;
    var btn = el('<button class="btn" style="margin-top:10px">Comprobar</button>');
    var clr = el('<button class="btn ghost" style="margin-top:10px;margin-left:8px">Limpiar</button>');
    clr.onclick = function () { answer.length = 0; paint(); };
    btn.onclick = function () {
      if (answer.join("") === ex.answer.join("")) {
        btn.remove(); clr.remove();
        line.classList.add("ok");
        h.revealExplain(body); h.done();
      } else {
        h.feedback(body, false, "Esa combinación no da. Revisá el orden y los paréntesis.");
      }
    };
    body.appendChild(btn); body.appendChild(clr);
  };

  /* ---- escribir la fórmula (se evalúa en la grilla) ---- */
  EX.t_formula = function (ex, body, h) {
    var grid = ex.grid;
    var expect = ex.expect;
    if (expect == null && ex.model && grid) {
      var mv = XL.calc(ex.model, { data: gridData(grid) });
      if (!XL.isErr(mv)) expect = mv;
    }
    if (grid) body.appendChild(el(gridHTML(grid, ex.cell)));
    if (ex.cell) body.appendChild(el('<div class="ex-hint">Escribí la fórmula que iría en la celda <code>' + esc(ex.cell) + "</code>:</div>"));

    var inp = el('<input class="ex-input" type="text" autocapitalize="off" autocorrect="off" spellcheck="false" placeholder="=…">');
    if (h.solved) { inp.value = ex.model || ex.accept && ex.accept[0] || ""; inp.disabled = true; }
    body.appendChild(inp);
    if (h.solved) return;

    var btn = el('<button class="btn" style="margin-top:10px">Comprobar</button>');
    function grade() {
      var raw = inp.value.trim();
      if (!raw) { h.feedback(body, false, "Escribí una fórmula (empieza con <code>=</code>)."); return; }
      // 1) comparación por texto contra respuestas aceptadas
      if (ex.accept && ex.accept.length) {
        var n = XL.normFormula(raw);
        for (var i = 0; i < ex.accept.length; i++) {
          if (XL.normFormula(ex.accept[i]) === n) { pass(); return; }
        }
      }
      // 2) comparación por resultado calculado sobre la grilla
      if (grid && expect != null) {
        var v = XL.calc(raw, { data: gridData(grid) });
        if (XL.isErr(v)) { h.feedback(body, false, "La fórmula devuelve <code>" + esc(v.e) + "</code>. Revisá referencias y funciones."); return; }
        var ok = (typeof expect === "number") ? XL.looseNum(v, expect)
          : String(v).trim().toLowerCase() === String(expect).trim().toLowerCase();
        if (ok) { pass(v); return; }
        h.feedback(body, false, "Tu fórmula da <code>" + esc(XL.fmt(v)) + "</code>, que no es el resultado esperado. Seguí probando.");
        return;
      }
      // 3) sin grilla ni accept: no se puede verificar solo → aceptar si "parece" una fórmula razonable
      if (!grid && !ex.accept) { pass(); return; }
      h.feedback(body, false, "Esa no coincide con la respuesta esperada.");
    }
    function pass(v) {
      inp.disabled = true; btn.remove();
      if (v != null) h.feedback(body, true, "Correcto → <code>" + esc(XL.fmt(v, ex.fmt)) + "</code>");
      h.revealExplain(body); h.done();
    }
    btn.onclick = grade;
    inp.addEventListener("keydown", function (e) { if (e.key === "Enter") grade(); });
    body.appendChild(btn);

    var hint = el('<details class="ex-peek"><summary>Ver la fórmula modelo</summary></details>');
    hint.appendChild(el('<div><code>' + esc(ex.model || (ex.accept && ex.accept[0]) || "—") + "</code></div>"));
    body.appendChild(hint);
  };

  /* ---- predecir el resultado de una fórmula dada ---- */
  EX.t_predict = function (ex, body, h) {
    var grid = ex.grid;
    var expect = ex.expect;
    if (expect == null && grid) {
      var v = XL.calc(ex.formula, { data: gridData(grid) });
      if (!XL.isErr(v)) expect = v;
    }
    if (grid) body.appendChild(el(gridHTML(grid, ex.cell)));
    body.appendChild(el('<div class="ex-hint">' + (ex.cell ? "En <code>" + esc(ex.cell) + "</code>: " : "") +
      "<code>" + esc(ex.formula) + "</code> &nbsp;→ ¿qué devuelve?</div>"));

    var inp = el('<input class="ex-input" type="text" autocapitalize="off" spellcheck="false" placeholder="resultado…">');
    if (h.solved) { inp.value = typeof expect === "number" ? XL.fmt(expect, ex.fmt) : String(expect); inp.disabled = true; }
    body.appendChild(inp);
    if (h.solved) return;

    var btn = el('<button class="btn" style="margin-top:10px">Comprobar</button>');
    function grade() {
      var raw = inp.value.trim();
      if (!raw) return;
      var ok = (typeof expect === "number")
        ? XL.looseNum(raw, expect)
        : String(raw).trim().toLowerCase() === String(expect).trim().toLowerCase();
      if (ok) {
        inp.disabled = true; btn.remove();
        h.feedback(body, true, "Correcto. El resultado es <code>" +
          esc(typeof expect === "number" ? XL.fmt(expect, ex.fmt) : expect) + "</code>");
        h.revealExplain(body); h.done();
      } else {
        h.feedback(body, false, "No es ese. Calculalo paso a paso y volvé a probar.");
      }
    }
    btn.onclick = grade;
    inp.addEventListener("keydown", function (e) { if (e.key === "Enter") grade(); });
    body.appendChild(btn);
  };

  /* ---------- router ---------- */
  var routes = {};
  function route(path) { return path.replace(/^#\/?/, "").split("/").filter(Boolean); }
  function go(hash) { window.location.hash = hash; }
  function render() {
    var parts = route(window.location.hash);
    var view = document.getElementById("view");
    view.innerHTML = "";
    var name = parts[0] || "";
    var fn = routes[name] || routes[""];
    fn(view, parts.slice(1));
    renderNav(name);
    renderSidebar(name, parts);
    try { window.scrollTo(0, 0); } catch (e) {}
  }
  window.addEventListener("hashchange", render);

  /* ---------- nav + sidebar ---------- */
  var NAV = [
    { key: "", label: "Inicio", hash: "#/" },
    { key: "m", label: "Módulos", hash: "#/m/1" },
    { key: "entrevista", label: "Entrevista", hash: "#/entrevista" },
    { key: "quiz", label: "Quiz", hash: "#/quiz" },
    { key: "cheatsheet", label: "Cheat Sheet", hash: "#/cheatsheet" }
  ];
  function renderNav(active) {
    var wrap = document.getElementById("nav-links");
    wrap.innerHTML = "";
    NAV.forEach(function (n) {
      var a = el('<a href="' + n.hash + '">' + n.label + "</a>");
      if (n.key === active) a.className = "active";
      wrap.appendChild(a);
    });
  }
  function renderSidebar(active, parts) {
    var sb = document.getElementById("sidebar");
    sb.innerHTML = "";
    var o = overall();
    sb.appendChild(el(
      '<div class="side-progress"><div class="num">' + o.pct + "%</div>" +
      '<div class="cap">avance total · ' + o.completed + "/" + o.total + " módulos</div>" +
      '<div class="bar" style="margin-top:10px"><i style="width:' + o.pct + '%"></i></div></div>'
    ));
    sb.appendChild(el('<div class="side-label">Navegación</div>'));
    [
      { label: "Inicio", hash: "#/", on: active === "" },
      { label: "Entrevista", hash: "#/entrevista", on: active === "entrevista" },
      { label: "Quiz", hash: "#/quiz", on: active === "quiz" },
      { label: "Cheat Sheet", hash: "#/cheatsheet", on: active === "cheatsheet" }
    ].forEach(function (l) {
      var d = el('<div class="side-link' + (l.on ? " active" : "") + '"><span class="dot"></span>' + l.label + "</div>");
      d.onclick = function () { go(l.hash); };
      sb.appendChild(d);
    });
    sb.appendChild(el('<div class="side-label">Módulos</div>'));
    var activeId = active === "m" ? Number(parts[1]) : null;
    C.modules.forEach(function (m) {
      var s = moduleStats(m);
      var d = el('<div class="side-link' + (m.id === activeId ? " active" : "") + '">' +
        '<span class="dot ' + s.status + '"></span>' +
        '<span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' +
        m.id + ". " + esc(m.title) + "</span></div>");
      d.onclick = function () { go("#/m/" + m.id); };
      sb.appendChild(d);
    });
  }

  /* ========================================================================
     DASHBOARD
     ===================================================================== */
  routes[""] = function (view) {
    var o = overall();
    var cur = currentModule();
    var cs = moduleStats(cur);
    view.appendChild(el('<div class="hello">Hola, ' + esc(C.meta.student) + "</div>"));
    view.appendChild(el('<div class="h">Seguí donde quedaste</div>'));

    var pct = Math.round(cs.frac * 100);
    var resume = el(
      '<div class="resume-card">' +
      '<div class="top"><span>' + (cs.status === "done" ? "Repaso" : "En curso") + " · Módulo " + cur.id + "</span><span>" + pct + "%</span></div>" +
      '<div class="ttl">' + esc(cur.title) + "</div>" +
      '<div class="sub">' + esc(cur.short) + "</div>" +
      '<div class="bar"><i style="width:' + pct + '%"></i></div>' +
      '<button class="btn">' + (cs.status === "pending" ? "Empezar módulo" : "Continuar módulo") + " →</button></div>"
    );
    resume.querySelector("button").onclick = function () { go("#/m/" + cur.id); };
    view.appendChild(resume);

    view.appendChild(el(
      '<div class="stats">' +
      '<div class="stat"><div class="n">' + o.completed + "/" + o.total + '</div><div class="c">módulos completos</div></div>' +
      '<div class="stat"><div class="n">' + o.pct + '%</div><div class="c">avance total del curso</div></div></div>'
    ));

    view.appendChild(el('<div class="eyebrow" style="margin-bottom:12px;display:block">Todos los módulos</div>'));
    var grid = el('<div class="mod-list"></div>');
    C.modules.forEach(function (m) {
      var s = moduleStats(m);
      var row = el(
        '<button class="mod-row ' + s.status + '">' +
        '<span class="idx">' + m.id + "</span>" +
        '<span class="body"><span class="t">' + esc(m.title) + '</span>' +
        '<span class="s ' + s.status + '">' + STATUS_LABEL[s.status] +
        (s.status !== "pending" ? " · " + Math.round(s.frac * 100) + "%" : "") + "</span></span>" +
        '<span class="mini"><span class="bar"><i style="width:' + Math.round(s.frac * 100) + '%"></i></span></span></button>'
      );
      row.onclick = function () { go("#/m/" + m.id); };
      grid.appendChild(row);
    });
    view.appendChild(grid);

    var pr = C.meta.priorityOrder.slice(0, 4).join(" · ");
    view.appendChild(el(
      '<div class="card quick" style="margin-top:26px">' +
      '<div class="eyebrow" style="display:block;margin-bottom:6px">Prioridad si tenés poco tiempo</div>' +
      '<p style="font-size:13px;color:var(--ink-soft);margin:0 0 8px">Estudiá primero los módulos <b>' + pr + '</b> (stock, consignación, liquidaciones, licitaciones y simulacro): es lo más específico del puesto.</p>' +
      '<a href="#/entrevista">Ir al simulacro de entrevista <span>→</span></a>' +
      '<a href="#/quiz">Banco de quiz <span>→</span></a>' +
      '<a href="#/cheatsheet">Cheat sheet de repaso final <span>→</span></a></div>'
    ));
    view.appendChild(el('<div class="footer-note">Herramienta de estudio personal · el progreso se guarda en este dispositivo</div>'));
  };

  /* ========================================================================
     MÓDULO
     ===================================================================== */
  routes["m"] = function (view, args) {
    var id = Number(args[0]) || 1;
    var mod = C.modules.filter(function (m) { return m.id === id; })[0];
    if (!mod) { go("#/m/1"); return; }
    var ms = modState(id);
    var tab = args[1] || "teoria";
    var st = moduleStats(mod);

    view.appendChild(el(
      '<div class="mod-head"><div class="kick">Módulo ' + id + " de " + C.modules.length + "</div>" +
      "<h1>" + esc(mod.title) + "</h1><p class=\"goal\">" + esc(mod.goal) + "</p></div>"
    ));

    var tabs = el('<div class="tabs"></div>');
    [
      { k: "teoria", label: "Teoría", badge: st.nCheck ? st.checkDone + "/" + st.nCheck : "" },
      { k: "ejercicios", label: "Ejercicios", badge: st.nEx ? st.exDone + "/" + st.nEx : "" },
      { k: "checklist", label: "Checklist", badge: st.nChk ? st.chkDone + "/" + st.nChk : "" },
      { k: "repaso", label: "Repaso" }
    ].forEach(function (t) {
      var b = el('<button class="tab' + (t.k === tab ? " active" : "") + '">' + t.label +
        (t.badge ? ' <span class="badge">' + t.badge + "</span>" : "") + "</button>");
      b.onclick = function () { go("#/m/" + id + "/" + t.k); };
      tabs.appendChild(b);
    });
    view.appendChild(tabs);
    var body = el("<div></div>");
    view.appendChild(body);

    if (tab === "teoria") {
      var practiceN = 0;
      mod.lessons.forEach(function (ls) {
        body.appendChild(el('<div class="lesson"><h3>' + esc(ls.h) + "</h3>" + ls.body + sheetHTML(ls.sheet) + "</div>"));
        (ls.checks || []).forEach(function (chk) {
          practiceN++;
          body.appendChild(EX.render(chk, { kind: "Práctica", n: practiceN, onSolve: function () { softRefreshTabs(); } }));
        });
      });
      var mark = el('<button class="btn ' + (ms.theory ? "ghost" : "") + '" style="margin-top:14px">' +
        (ms.theory ? "✓ Teoría marcada como leída" : "Marcar teoría como leída") + "</button>");
      mark.onclick = function () {
        ms.theory = !ms.theory; save();
        toast(ms.theory ? "Teoría completada" : "Marca quitada");
        render();
      };
      body.appendChild(mark);
      if (st.nEx) {
        var goEx = el('<a href="#/m/' + id + '/ejercicios" class="btn dark" style="margin-top:14px;margin-left:8px;display:inline-flex">Ir a los ejercicios →</a>');
        body.appendChild(goEx);
      }
    }

    else if (tab === "ejercicios") {
      if (!mod.exercises || !mod.exercises.length) {
        body.appendChild(el('<div class="card" style="color:var(--ink-soft)">Este módulo se practica en el <a href="#/entrevista">simulacro de entrevista</a>.</div>'));
      } else {
        body.appendChild(el('<p style="font-size:13.5px;color:var(--ink-soft);margin:0 0 16px">Resolvé cada uno en la app. Se corrigen solos con el mismo criterio que usaría Excel.</p>'));
        mod.exercises.forEach(function (ex, i) {
          body.appendChild(EX.render(ex, { kind: "Ejercicio", n: i + 1, onSolve: function () { softRefreshTabs(); } }));
        });
      }
    }

    else if (tab === "checklist") {
      body.appendChild(el('<p style="font-size:13.5px;color:var(--ink-soft);margin:0 0 14px">Tildá cada punto solo cuando puedas explicarlo con tus palabras y resolver un ejercicio sin ayuda.</p>'));
      var box = el('<div class="card"></div>');
      mod.checklist.forEach(function (item, i) {
        var row = el('<label class="rowcheck' + (ms.chk[i] ? " on" : "") + '"><input type="checkbox"' + (ms.chk[i] ? " checked" : "") + "> " + esc(item) + "</label>");
        row.querySelector("input").onchange = function (e) {
          ms.chk[i] = e.target.checked; save();
          if (moduleStats(mod).status === "done") toast("Módulo " + id + " completo ✓");
          render();
        };
        box.appendChild(row);
      });
      body.appendChild(box);
    }

    else if (tab === "repaso") {
      body.appendChild(el('<div class="eyebrow" style="display:block;margin-bottom:12px">Puntos clave del módulo</div>'));
      var kg = el('<div class="keys-grid"></div>');
      mod.keys.forEach(function (kv) {
        kg.appendChild(el('<div class="key-item"><div class="k">' + esc(kv.k) + '</div><div class="v">' + esc(kv.v) + "</div></div>"));
      });
      body.appendChild(kg);
      var hasQuiz = C.quiz.some(function (q) { return q.module === id; });
      body.appendChild(el(hasQuiz
        ? '<a href="#/quiz/' + id + '" class="btn dark" style="margin-top:20px;display:inline-flex">Hacer el quiz de este módulo →</a>'
        : '<a href="#/entrevista" class="btn dark" style="margin-top:20px;display:inline-flex">Ir al simulacro de entrevista →</a>'));
    }

    function softRefreshTabs() {
      var s2 = moduleStats(mod);
      var badges = tabs.querySelectorAll(".badge");
      if (badges[0]) badges[0].textContent = s2.checkDone + "/" + s2.nCheck;
      if (badges[1]) badges[1].textContent = s2.exDone + "/" + s2.nEx;
      renderSidebar("m", ["m", String(id)]);
    }

    var nav = el('<div class="mod-nav"></div>');
    var prev = el('<button class="btn ghost"' + (id <= 1 ? " disabled" : "") + ">← Módulo " + (id - 1) + "</button>");
    var next = el('<button class="btn ghost"' + (id >= C.modules.length ? " disabled" : "") + ">Módulo " + (id + 1) + " →</button>");
    prev.onclick = function () { if (id > 1) go("#/m/" + (id - 1)); };
    next.onclick = function () { if (id < C.modules.length) go("#/m/" + (id + 1)); };
    nav.appendChild(prev); nav.appendChild(next);
    view.appendChild(nav);
  };

  /* ========================================================================
     QUIZ
     ===================================================================== */
  var quizSession = null;
  routes["quiz"] = function (view, args) {
    var moduleFilter = args[0] ? Number(args[0]) : null;
    view.appendChild(el(
      '<div class="page-head"><span class="eyebrow">Autoevaluación</span><h1>Quiz</h1>' +
      "<p>Elegí un módulo o hacé el integrador. Feedback inmediato con la explicación de cada respuesta.</p></div>"
    ));
    var pickers = el('<div class="quiz-pickers"></div>');
    var allChip = el('<button class="chip' + (!moduleFilter ? " active" : "") + '">Integrador (todos)</button>');
    allChip.onclick = function () { quizSession = null; go("#/quiz"); };
    pickers.appendChild(allChip);
    C.modules.forEach(function (m) {
      var n = C.quiz.filter(function (q) { return q.module === m.id; }).length;
      if (!n) return;
      var ch = el('<button class="chip' + (moduleFilter === m.id ? " active" : "") + '">M' + m.id + " · " + n + "</button>");
      ch.title = m.title;
      ch.onclick = function () { quizSession = null; go("#/quiz/" + m.id); };
      pickers.appendChild(ch);
    });
    view.appendChild(pickers);

    if (!quizSession || quizSession.moduleFilter !== moduleFilter) {
      var pool = C.quiz.filter(function (q) { return !moduleFilter || q.module === moduleFilter; });
      quizSession = { pool: shuffle(pool.slice()), idx: 0, answered: {}, picked: {}, moduleFilter: moduleFilter };
    }
    var s = quizSession;
    var body = el("<div></div>");
    view.appendChild(body);

    var dots = el('<div class="quiz-dots"></div>');
    s.pool.forEach(function (q, i) {
      var d = el('<i>' + (i + 1) + "</i>");
      d.onclick = function () { if (s.idx < s.pool.length) { s.idx = i; draw(); } };
      dots.appendChild(d);
    });
    body.appendChild(dots);
    var qHost = el("<div></div>");
    body.appendChild(qHost);

    function draw() {
      Array.prototype.forEach.call(dots.children, function (d, i) {
        var q = s.pool[i];
        d.className = i === s.idx ? "now" : s.answered[q.id] === true ? "ok" : s.answered[q.id] === false ? "no" : "";
      });
      qHost.innerHTML = "";

      if (!s.pool.length) {
        qHost.appendChild(el('<div class="card" style="text-align:center;color:var(--ink-soft)">Sin preguntas para este filtro. Practicá en el <a href="#/entrevista">simulacro</a>.</div>'));
        return;
      }
      if (s.idx >= s.pool.length) {
        var correct = 0;
        s.pool.forEach(function (q) { if (s.answered[q.id]) correct++; });
        var p = Math.round((correct / s.pool.length) * 100);
        var msg = p >= 80 ? "Muy sólido. Estás para la entrevista." : p >= 60 ? "Bien encaminado. Repasá las que fallaste." : "Volvé a la teoría de estos temas y reintentá.";
        var card = el('<div class="card score"><div class="big">' + correct + "/" + s.pool.length + '</div><div class="lbl">' + p + "% · " + msg + '</div>' +
          '<div style="margin-top:20px;display:flex;gap:10px;justify-content:center;flex-wrap:wrap">' +
          '<button class="btn" id="q-retry">Reintentar</button><button class="btn ghost" id="q-review">Ver todas</button></div></div>');
        qHost.appendChild(card);
        card.querySelector("#q-retry").onclick = function () { quizSession = null; render(); };
        card.querySelector("#q-review").onclick = function () { s.idx = 0; draw(); };
        return;
      }

      var q = s.pool[s.idx];
      var mod = C.modules.filter(function (m) { return m.id === q.module; })[0];
      var already = s.answered.hasOwnProperty(q.id);
      var pick = s.picked[q.id];

      qHost.appendChild(el('<div class="quiz-top"><span class="count">Pregunta ' + (s.idx + 1) + " de " + s.pool.length +
        '</span><span class="pill">Módulo ' + q.module + " · " + esc(mod ? mod.short.split(",")[0] : "") + "</span></div>"));
      qHost.appendChild(el('<div class="bar" style="margin-bottom:4px"><i style="width:' + Math.round((s.idx / s.pool.length) * 100) + '%"></i></div>'));
      qHost.appendChild(el('<div class="quiz-q">' + esc(q.q) + "</div>"));

      var opts = el('<div class="quiz-opts"></div>');
      q.options.forEach(function (o, oi) {
        var cls = "opt";
        if (already) { if (o.correct) cls += " correct"; else if (pick === oi) cls += " wrong"; }
        var b = el('<button class="' + cls + '"' + (already ? " disabled" : "") + "><span class=\"mk\">" +
          (already && o.correct ? "✓" : already && pick === oi ? "✕" : String.fromCharCode(65 + oi)) + "</span><span>" + esc(o.t) + "</span></button>");
        b.onclick = function () {
          if (s.answered.hasOwnProperty(q.id)) return;
          s.picked[q.id] = oi;
          s.answered[q.id] = !!o.correct;
          state.quiz[q.id] = { correct: !!o.correct };
          save();
          draw();
        };
        opts.appendChild(b);
      });
      qHost.appendChild(opts);

      if (already) {
        var ok = s.answered[q.id];
        qHost.appendChild(el('<div class="quiz-fb ' + (ok ? "ok" : "no") + '"><b>' + (ok ? "Correcto" : "Incorrecto") + "</b>" + esc(q.explain) + "</div>"));
      }
      var actions = el('<div class="quiz-actions"></div>');
      var back = el('<button class="btn ghost"' + (s.idx === 0 ? " disabled" : "") + ">← Anterior</button>");
      back.onclick = function () { if (s.idx > 0) { s.idx--; draw(); } };
      var fwd = el('<button class="btn dark"' + (!already ? " disabled" : "") + ">" + (s.idx === s.pool.length - 1 ? "Ver resultado" : "Siguiente →") + "</button>");
      fwd.onclick = function () { s.idx++; draw(); };
      actions.appendChild(back); actions.appendChild(fwd);
      qHost.appendChild(actions);
    }
    draw();
  };

  /* ========================================================================
     ENTREVISTA
     ===================================================================== */
  routes["entrevista"] = function (view) {
    var iv = C.interview;
    view.appendChild(el(
      '<div class="page-head"><span class="eyebrow">Módulo 7 · Simulacro</span>' +
      "<h1>Simulacro de entrevista técnica</h1>" +
      "<p>Escribí tu respuesta <b>antes</b> de abrir la respuesta modelo. Se guarda sola en este dispositivo.</p></div>"
    ));

    iv.groups.forEach(function (g) {
      var gh = el('<div class="iv-group"><h2>' + esc(g.title) + "</h2></div>");
      view.appendChild(gh);
      g.items.forEach(function (it) {
        var saved = state.interview[it.id] || "";
        var item = el(
          '<div class="iv-item">' +
          '<div class="iv-q">' + esc(it.q) + "</div>" +
          '<div class="iv-body">' +
          '<div class="iv-aim"><b>Qué buscan:</b> ' + esc(it.aim) + "</div>" +
          '<textarea placeholder="Tu respuesta…">' + esc(saved) + "</textarea>" +
          '<div class="iv-saved"></div>' +
          '<details class="iv-model"><summary></summary><div class="txt">' + esc(it.model) + "</div></details>" +
          "</div></div>"
        );
        item.querySelector(".iv-q").onclick = function () { item.classList.toggle("open"); };
        var ta = item.querySelector("textarea");
        var savedLbl = item.querySelector(".iv-saved");
        var t;
        ta.addEventListener("input", function () {
          clearTimeout(t);
          t = setTimeout(function () {
            state.interview[it.id] = ta.value; save();
            savedLbl.textContent = "Guardado";
            setTimeout(function () { savedLbl.textContent = ""; }, 1200);
          }, 400);
        });
        gh.appendChild(item);
      });
    });

    if (iv.cases && iv.cases.length) {
      view.appendChild(el('<div class="iv-group"><h2>«Resolvé esto en Excel» — casos interactivos</h2></div>'));
      var host = view.lastChild;
      iv.cases.forEach(function (cs, i) {
        host.appendChild(EX.render(cs, { kind: "Caso", n: i + 1, onSolve: function () {} }));
      });
    }
  };

  /* ========================================================================
     CHEAT SHEET
     ===================================================================== */
  routes["cheatsheet"] = function (view) {
    view.appendChild(el(
      '<div class="page-head"><span class="eyebrow">Repaso final</span><h1>Cheat sheet</h1>' +
      "<p>Todo lo esencial en una vista, para repasar 2–3 minutos antes de entrar. Podés imprimirlo o guardarlo como PDF.</p></div>"
    ));
    var grid = el('<div class="cheat-grid"></div>');
    C.cheat.forEach(function (g) {
      var col = el('<div class="cheat-col"><h3>' + esc(g.title) + "</h3></div>");
      g.items.forEach(function (it) {
        col.appendChild(el('<div class="cheat-item"><div class="k">' + esc(it.k) + '</div><div class="v">' + esc(it.v) + "</div></div>"));
      });
      grid.appendChild(col);
    });
    view.appendChild(grid);
    var hint = el('<div class="print-hint"><span class="l">Antes de entrar</span><button class="btn ghost" id="print-btn">Imprimir / Guardar PDF</button></div>');
    hint.querySelector("#print-btn").onclick = function () { window.print(); };
    view.appendChild(hint);
  };

  /* ---------- arranque ---------- */
  if (!window.location.hash) window.location.hash = "#/";
  render();
})();
