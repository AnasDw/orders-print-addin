/* global Office, Excel, parseDaySheet, fmtQty */
(function () {
  var current = null, selected = {};
  var PER_TABLE = 10;
  var OUT_SHEET = "הדפסה";
  var $ = function (id) { return document.getElementById(id); };

  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function msg(t, cls) { var m = $("msg"); m.textContent = t || ""; m.className = cls || ""; }

  // ---------- Excel access (the only part that touches Office.js) ----------
  function readActiveSheet() {
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      var rg = ws.getRange("A1:Z700");
      rg.load("values");
      return ctx.sync().then(function () { return { name: ws.name, values: rg.values }; });
    });
  }

  function colName(i) { // 0-based -> letters
    var s = ""; i++;
    while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }

  function writeOutput(blocks) {
    return Excel.run(function (ctx) {
      var wbk = ctx.workbook;
      var old = wbk.worksheets.getItemOrNullObject(OUT_SHEET);
      return ctx.sync().then(function () {
        if (!old.isNullObject) { old.delete(); }
        return ctx.sync();
      }).then(function () {
        var ws = wbk.worksheets.add(OUT_SHEET);
        var row = 0;
        var breaks = [];
        blocks.forEach(function (b, bi) {
          var width = b.width;
          var data = b.rows.map(function (r) { while (r.length < width) r.push(""); return r; });
          var rg = ws.getRangeByIndexes(row, 0, data.length, width);
          rg.values = data;
          if (bi > 0) breaks.push(row);
          // title
          var t = ws.getRangeByIndexes(row, 0, 1, 1);
          t.format.font.bold = true; t.format.font.size = 14;
          // header
          var h = ws.getRangeByIndexes(row + 1, 0, 1, width);
          h.format.font.bold = true; h.format.fill.color = "#E5E7EB";
          h.format.horizontalAlignment = "Center"; h.format.wrapText = true;
          // table body
          var tbl = ws.getRangeByIndexes(row + 1, 0, b.tableRows, width);
          ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"].forEach(function (e) {
            var bd = tbl.format.borders.getItem(e); bd.style = "Continuous"; bd.color = "#6B7280";
          });
          if (b.tableRows > 1) {
            var q = ws.getRangeByIndexes(row + 2, 2, b.tableRows - 1, width - 2);
            q.format.horizontalAlignment = "Center"; q.format.font.bold = true; q.format.font.size = 14;
            ws.getRangeByIndexes(row + 2, 2, b.tableRows - 1, 1).format.fill.color = "#F3F4F6";
            ws.getRangeByIndexes(row + 2, 1, b.tableRows - 1, 1).format.font.size = 9;
          }
          if (b.notesAt >= 0) {
            ws.getRangeByIndexes(row + b.notesAt, 0, 1, 1).format.font.bold = true;
          }
          row += data.length + 1;
        });
        ws.getRange("A:A").format.columnWidth = 190;
        ws.getRange("B:B").format.columnWidth = 75;
        ws.getRange("C:C").format.columnWidth = 50;
        ws.getRange("D:M").format.columnWidth = 75;
        ws.activate();
        return ctx.sync().then(function () {
          // Optional niceties: each in its own sync so a failure doesn't block the rest.
          function tryDo(fn) { try { fn(); } catch (e) {} return ctx.sync().catch(function () {}); }
          return tryDo(function () { ws.rightToLeft = true; })
            .then(function () { return tryDo(function () {
              ws.pageLayout.orientation = "Landscape";
              ws.pageLayout.zoom = { scale: 80 };
            }); })
            .then(function () { return tryDo(function () {
              breaks.forEach(function (r) { ws.horizontalPageBreaks.add(ws.getRangeByIndexes(r, 0, 1, 1)); });
            }); });
        });
      });
    });
  }

  // ---------- Build print blocks ----------
  function buildBlocks(chosen, includeGeneral) {
    var blocks = [];
    for (var start = 0; start < chosen.length; start += PER_TABLE) {
      var cs = chosen.slice(start, start + PER_TABLE);
      var isLast = start + PER_TABLE >= chosen.length;
      var rowMap = {};
      cs.forEach(function (c, k) {
        c.items.forEach(function (it) {
          var r = rowMap[it.pos] || (rowMap[it.pos] = { pos: it.pos, name: it.name, supplier: it.supplier, q: {} });
          r.q[k] = it.qty;
        });
      });
      var rows = Object.keys(rowMap).map(function (p) { return rowMap[p]; }).sort(function (a, b) { return a.pos - b.pos; });
      var width = 3 + cs.length;
      var out = [];
      out.push(["טופס הזמנה ליום: " + current.day + "     תאריך: " + current.date]);
      out.push(["", "", 'סה"כ'].concat(cs.map(function (c) { return c.name; })));
      if (!rows.length) out.push(["אין פריטים"]);
      rows.forEach(function (r) {
        var tot = 0, line = [r.name, r.supplier, 0];
        cs.forEach(function (c, k) { var v = r.q[k] || 0; tot += v; line.push(v || ""); });
        line[2] = tot;
        out.push(line);
      });
      var tableRows = out.length - 1;
      var notes = [];
      cs.forEach(function (c) { c.notes.forEach(function (t) { if (notes.indexOf(t) < 0) notes.push(t); }); });
      if (includeGeneral && isLast) current.generalNotes.forEach(function (t) { if (notes.indexOf(t) < 0) notes.push(t); });
      var notesAt = -1;
      if (notes.length) {
        out.push([""]);
        notesAt = out.length;
        out.push(["הערות:"]);
        notes.forEach(function (t) { out.push([t]); });
      }
      blocks.push({ rows: out, width: width, tableRows: tableRows, notesAt: notesAt });
    }
    return blocks;
  }

  // ---------- UI ----------
  function load() {
    msg("טוען...");
    var p = (window.__readSheet || readActiveSheet)();
    p.then(function (res) {
      current = parseDaySheet(res.values);
      selected = {};
      $("sheetInfo").textContent = "גיליון: " + res.name + (current.date ? "  |  " + current.date : "");
      if (!current.customers.length) {
        msg("לא נמצאו לקוחות בגיליון הזה. פתח גיליון הזמנה ולחץ רענן.", "err");
      } else { msg(""); }
      renderGeneral(); renderList(); updateBtn();
    }).catch(function (e) { msg("שגיאה בקריאת הגיליון: " + (e && e.message || e), "err"); });
  }

  function renderGeneral() {
    var g = current.generalNotes, w = $("gen");
    if (g.length) {
      w.classList.remove("hidden"); $("genLbl").classList.remove("hidden");
      w.innerHTML = "<b>הערות כלליות:</b><br>" + g.map(esc).join("<br>");
    } else { w.classList.add("hidden"); $("genLbl").classList.add("hidden"); $("genPrint").checked = false; }
  }

  function renderList() {
    var q = $("search").value.trim(), list = $("list");
    list.innerHTML = "";
    current.customers.forEach(function (c) {
      if (q && c.name.indexOf(q) === -1) return;
      var d = document.createElement("label");
      d.className = "c" + (selected[c.id] ? " on" : "") + (c.items.length ? "" : " z");
      var cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !!selected[c.id];
      cb.onchange = function () {
        if (cb.checked) selected[c.id] = true; else delete selected[c.id];
        d.classList.toggle("on", cb.checked); updateBtn();
      };
      var n = document.createElement("span"); n.className = "n"; n.textContent = c.name;
      var k = document.createElement("span"); k.className = "k"; k.textContent = c.items.length + (c.notes.length ? " + הערה" : "");
      d.appendChild(cb); d.appendChild(n); d.appendChild(k); list.appendChild(d);
    });
  }

  function chosenList() { return current.customers.filter(function (c) { return selected[c.id]; }); }
  function updateBtn() {
    var n = chosenList().length;
    $("go").disabled = !n;
    $("go").textContent = n ? "צור הדפסה (" + n + " לקוחות)" : "צור הדפסה";
  }

  function go() {
    var chosen = chosenList();
    if (!chosen.length) return;
    var blocks = buildBlocks(chosen, $("genPrint").checked);
    if (window.__writeOutput) { window.__writeOutput(blocks); msg("נוצר", "ok"); return; }
    msg("יוצר גיליון הדפסה...");
    $("go").disabled = true;
    writeOutput(blocks).then(function () {
      msg('נוצר הגיליון "' + OUT_SHEET + '". לחץ Ctrl+P להדפסה.', "ok");
      updateBtn();
    }).catch(function (e) { msg("שגיאה: " + (e && e.message || e), "err"); updateBtn(); });
  }

  function init() {
    $("reload").onclick = load;
    $("search").oninput = renderList;
    $("selAll").onclick = function () {
      var q = $("search").value.trim();
      current.customers.forEach(function (c) {
        if (q && c.name.indexOf(q) === -1) return;
        if (c.items.length || c.notes.length) selected[c.id] = true;
      });
      renderList(); updateBtn();
    };
    $("selNone").onclick = function () { selected = {}; renderList(); updateBtn(); };
    $("go").onclick = go;
    load();
  }

  window.__buildBlocks = buildBlocks;
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function () { init(); });
  } else {
    window.addEventListener("load", init); // test mode, no Office.js
  }
})();
