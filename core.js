// Core parsing logic. Works in the browser and in node (for testing).
// Input: array of rows (array of arrays) for one day sheet.
// Output: { day, date, customers: [{id, name, items:[{name, supplier, qty}], notes:[]}], generalNotes:[] }

var FIRST_CUST_COL = 3;   // column D
var LAST_CUST_COL = 22;   // column W
var NOTES_COL = 24;       // column Y

function cellStr(v) {
  return v === undefined || v === null ? "" : String(v).trim();
}

function isHeaderRow(row) {
  return cellStr(row && row[0]).indexOf("טופס הזמנה") === 0;
}

function normName(s) {
  return cellStr(s).replace(/[*$]/g, "").replace(/\s+/g, " ").trim();
}

function parseDaySheet(rows) {
  var headerRows = [];
  for (var r = 0; r < rows.length; r++) if (isHeaderRow(rows[r])) headerRows.push(r);

  var day = headerRows.length ? cellStr(rows[headerRows[0]][1]) : "";
  var date = headerRows.length ? cellStr(rows[headerRows[0]][17]) : "";

  // Pages that share the same customer header belong to one customer group.
  var groups = [];
  headerRows.forEach(function (h, i) {
    var end = i + 1 < headerRows.length ? headerRows[i + 1] : rows.length;
    var hdr = rows[h + 1] || [];
    var names = [];
    for (var c = FIRST_CUST_COL; c <= LAST_CUST_COL; c++) names.push(cellStr(hdr[c]));
    var sig = names.join("|");
    var last = groups[groups.length - 1];
    if (last && last.sig === sig) last.ranges.push([h + 2, end]);
    else groups.push({ sig: sig, names: names, ranges: [[h + 2, end]] });
  });

  var customers = [];
  groups.forEach(function (g, gi) {
    g.names.forEach(function (name, idx) {
      var col = FIRST_CUST_COL + idx;
      var unnamed = !name;
      var items = [];
      var pos = 0;
      g.ranges.forEach(function (rg) {
        for (var r = rg[0]; r < rg[1]; r++) {
          var row = rows[r] || [];
          var prod = cellStr(row[0]);
          if (!prod || isHeaderRow(row)) continue;
          pos++;
          var q = row[col];
          if (typeof q === "string" && q.trim() !== "" && !isNaN(Number(q))) q = Number(q);
          if (typeof q === "number" && q > 0) {
            items.push({ pos: pos, name: prod, supplier: cellStr(row[1]), qty: q });
          }
        }
      });
      if (unnamed) {
        // Quantities under a column with no customer name: keep them so nothing is lost.
        if (!items.length) return;
        name = "(ללא שם לקוח) עמודה " + String.fromCharCode(65 + col);
      }
      customers.push({ id: gi + "-" + col, name: name, norm: unnamed ? "" : normName(name), unnamed: unnamed, items: items, notes: [] });
    });
  });

  // Notes: column Y, below the "הערות" title.
  var notes = [];
  var start = -1;
  for (var r2 = 0; r2 < rows.length; r2++) {
    if (cellStr((rows[r2] || [])[NOTES_COL]) === "הערות") { start = r2; break; }
  }
  if (start >= 0) {
    for (var r3 = start + 1; r3 < rows.length; r3++) {
      var t = cellStr((rows[r3] || [])[NOTES_COL]);
      if (t) notes.push(t);
    }
  }

  var generalNotes = [];
  notes.forEach(function (note) {
    var n = normName(note);
    var hit = false;
    customers.forEach(function (c) {
      if (c.norm.length >= 3 && n.indexOf(c.norm) !== -1) {
        c.notes.push(note);
        hit = true;
      }
    });
    if (!hit) generalNotes.push(note);
  });

  return { day: day, date: date, customers: customers, generalNotes: generalNotes };
}

function fmtQty(q) {
  return Number.isInteger(q) ? String(q) : String(Math.round(q * 100) / 100);
}

if (typeof module !== "undefined") module.exports = { parseDaySheet: parseDaySheet, fmtQty: fmtQty, normName: normName };
