/* AnEspresso assignment-sheet parser + late board. Unshipped. */
(function (g) {
  "use strict";

  var SHIFT_LATE = { M: 1, S: 1 };
  var SHIFT_DINNER = { Q: 1, W: 1, E: 1 };
  var ROOM_INDEX = null;

  function buildRoomIndex() {
    ROOM_INDEX = {};
    if (typeof CATEGORIES === "undefined") return;
    CATEGORIES.forEach(function (c) {
      (c.rooms || []).forEach(function (r) {
        ROOM_INDEX[String(r).toUpperCase()] = { cat: c.id, room: r };
      });
    });
  }

  function u16(v, i) { return v.getUint16(i, true); }
  function u32(v, i) { return v.getUint32(i, true); }

  async function inflateRaw(data) {
    if (typeof DecompressionStream === "undefined") {
      throw new Error("This phone cannot unpack Excel. Use Files in a current Safari.");
    }
    var ds = new DecompressionStream("deflate-raw");
    var stream = new Blob([data]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function unzip(buf) {
    var u = new Uint8Array(buf);
    var v = new DataView(buf);
    var eocd = -1;
    var min = Math.max(0, u.length - 22 - 65557);
    for (var i = u.length - 22; i >= min; i--) {
      if (u32(v, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("Not an Excel workbook (.xlsx)");
    var n = u16(v, eocd + 10);
    var cdOff = u32(v, eocd + 16);
    var p = cdOff;
    var files = {};
    for (var k = 0; k < n; k++) {
      if (u32(v, p) !== 0x02014b50) break;
      var method = u16(v, p + 10);
      var comp = u32(v, p + 20);
      var nameLen = u16(v, p + 28);
      var extra = u16(v, p + 30);
      var comment = u16(v, p + 32);
      var localOff = u32(v, p + 42);
      var name = new TextDecoder().decode(u.slice(p + 46, p + 46 + nameLen));
      var localNameLen = u16(v, localOff + 26);
      var localExtra = u16(v, localOff + 28);
      var dataStart = localOff + 30 + localNameLen + localExtra;
      files[name] = { method: method, data: u.slice(dataStart, dataStart + comp) };
      p += 46 + nameLen + extra + comment;
    }
    var out = {};
    var names = Object.keys(files);
    for (var j = 0; j < names.length; j++) {
      var f = files[names[j]];
      if (f.method === 0) out[names[j]] = f.data;
      else if (f.method === 8) out[names[j]] = await inflateRaw(f.data);
    }
    return out;
  }

  function xmlText(bytes) {
    return new TextDecoder("utf-8").decode(bytes);
  }

  function loadSharedStrings(xml) {
    var ss = [];
    if (!xml) return ss;
    var doc = new DOMParser().parseFromString(xml, "text/xml");
    var nodes = doc.getElementsByTagName("si");
    for (var i = 0; i < nodes.length; i++) {
      var ts = nodes[i].getElementsByTagName("t");
      var s = "";
      for (var j = 0; j < ts.length; j++) s += ts[j].textContent || "";
      ss.push(s);
    }
    return ss;
  }

  function loadCells(sheetXml, ss) {
    var doc = new DOMParser().parseFromString(sheetXml, "text/xml");
    var cells = {};
    var nodes = doc.getElementsByTagName("c");
    for (var i = 0; i < nodes.length; i++) {
      var c = nodes[i];
      var ref = c.getAttribute("r");
      if (!ref) continue;
      var t = c.getAttribute("t");
      var v = c.getElementsByTagName("v")[0];
      var isEl = c.getElementsByTagName("is")[0];
      var val = "";
      if (t === "s" && v && v.textContent) val = ss[parseInt(v.textContent, 10)] || "";
      else if (t === "inlineStr" && isEl) {
        var ts = isEl.getElementsByTagName("t");
        for (var j = 0; j < ts.length; j++) val += ts[j].textContent || "";
      } else if (v) val = v.textContent || "";
      if (val && String(val).trim()) cells[ref] = String(val).trim();
    }
    return cells;
  }

  function cell(cells, col, row) {
    return (cells[col + row] || "").trim();
  }

  function excelDate(val) {
    val = String(val || "").trim();
    var m = val.match(/(\d{1,2})[./](\d{1,2})[./](\d{2,4})/);
    if (m) {
      var mo = parseInt(m[1], 10), d = parseInt(m[2], 10), y = parseInt(m[3], 10);
      if (y < 100) y += 2000;
      return y + "-" + mo + "-" + d;
    }
    if (/^\d+(\.0+)?$/.test(val)) {
      var serial = parseInt(val, 10);
      var dt = new Date(Date.UTC(1899, 11, 30));
      dt.setUTCDate(dt.getUTCDate() + serial);
      return dt.getUTCFullYear() + "-" + (dt.getUTCMonth() + 1) + "-" + dt.getUTCDate();
    }
    return null;
  }

  function cleanName(s) {
    s = String(s || "").replace(/\*/g, " ");
    s = s.replace(/\s*\+.*$/, "");
    s = s.replace(/\s*\(.*\)$/, "");
    s = s.replace(/\s*REQ\s*$/i, "");
    return s.replace(/\s+/g, " ").replace(/^[\s-]+|[\s-]+$/g, "");
  }

  function lastName(name) {
    var p = String(name || "").trim().split(/\s+/);
    return p.length ? p[p.length - 1] : "";
  }

  function normalizeShift(sh) {
    return String(sh || "").split("/").map(function (p) {
      p = p.trim();
      if (p === "d") return "d";
      if (p.toLowerCase() === "s") return "S";
      return p.toUpperCase();
    }).join("/");
  }

  function coreShift(shift) {
    return String(shift || "").replace(/^o\//i, "").replace(/^\*/, "").split("/")[0];
  }

  function isEarlyShift(shift) {
    var s = String(shift || "").replace(/^o\//i, "");
    return s.charAt(0) === "*";
  }

  function fmtShift(letter, early, orient) {
    if (!letter || letter === "Dr") return letter || "";
    return (orient ? "o/" : "") + (early ? "*" : "") + letter;
  }

  function breakKind(shift) {
    var day = coreShift(shift);
    if (!day || day === "Dr" || /^\d/.test(day)) return "none";
    if (SHIFT_LATE[day]) return "late";
    if (SHIFT_DINNER[day]) return "dinner";
    return "none";
  }

  function parseStaff(raw) {
    var s = String(raw || "").replace(/\s+/g, " ").trim();
    if (!s) return null;
    if (/^closed$/i.test(s)) return { closed: true, shift: "", name: "", kind: "none", early: false, orient: false };
    if (/^#ref/i.test(s)) return null;
    if (/rotate/i.test(s) && !/[A-Za-z]{4,}\s+[A-Za-z]{3,}/.test(s)) return null;

    var rest = s;
    var orient = false;
    var early = false;
    for (var i = 0; i < 6; i++) {
      if (/^o\//i.test(rest)) { orient = true; rest = rest.replace(/^o\//i, "").trim(); continue; }
      if (/^\*/.test(rest)) { early = true; rest = rest.replace(/^\*+/, "").trim(); continue; }
      break;
    }
    if (/[DdMmSsQqWwEeNnTt]\*/.test(s)) early = true;

    var m = rest.match(/^(\d{1,2}a-\d{1,2}p)\s+(.+)$/i);
    if (m) {
      return { closed: false, shift: m[1].toLowerCase(), name: cleanName(m[2]), kind: "none", early: early, orient: orient };
    }
    m = rest.match(/^(Dr\.?)\s+(.+)$/i);
    if (m) {
      return { closed: false, shift: "Dr", name: cleanName(m[2]), kind: "none", early: false, orient: orient };
    }
    m = rest.match(/^([DdMmSsQqWwEeNnTt](?:\/[DdMmSsQqWwEeNnTt])?)\*?\s*(.+)$/);
    if (m && /[A-Za-z]/.test(m[2])) {
      var sh = normalizeShift(m[1]);
      return {
        closed: false,
        shift: fmtShift(sh, early, orient),
        name: cleanName(m[2]),
        kind: breakKind(sh),
        early: early,
        orient: orient
      };
    }
    m = rest.match(/^([DdMmSsQqWwEeNnTt])\*?(?=[A-Z])([A-Z].+)$/);
    if (m) {
      sh = normalizeShift(m[1]);
      return {
        closed: false,
        shift: fmtShift(sh, early, orient),
        name: cleanName(m[2]),
        kind: breakKind(sh),
        early: early,
        orient: orient
      };
    }
    return { closed: false, shift: "", name: cleanName(rest), kind: "none", early: false, orient: orient };
  }

  function skipRoomLabel(s) {
    return /^(NORTH TOWER|SOUTH TOWER|ENDO|EP|OFFSITE|CCS|FBC|M\/N|STE 100|OB|WBF|ENDO A|ENDO B|RESOURCE STAFF|POC TESTING|CCS POC|2N POC|2S POC|STE POC|OB POC|CHECK HEMACUE|TRAUMA RM|SHIFT|CRNA|ASSIGN|MIDNIGHT|LATE STAY|#\d|MN CALL|HEART CALL|NT BREAKERS|ST BREAKERS|POS|CV|3N|2N|STE 1|STE 2|ST 1|ST 2)/i.test(String(s || "").trim());
  }

  function normalizeRoom(raw) {
    if (!ROOM_INDEX) buildRoomIndex();
    var s = String(raw || "").replace(/\s+/g, " ").trim();
    if (!s || skipRoomLabel(s)) return null;
    var su = s.toUpperCase();
    su = su.replace(/^STE\.?\s+/, "OR ");
    su = su.replace(/^ENDO\s*/, "ENDO ");
    su = su.replace("OB1", "OB 1").replace("OB2", "OB 2");
    su = su.replace("1ST MRI", "MRI 1ST");
    su = su.replace(/U\/S.*/, "US");
    if (su === "PET") su = "PET SCAN";
    su = su.replace(/^IR\s*(\d+).*/, "IR $1");
    su = su.replace(/^BMBX?\d*.*/, "BMB");
    su = su.replace(/^CT\b.*/, "CT");
    su = su.replace("OR39/40", "OR 38-39").replace("OR 39/40", "OR 38-39");
    su = su.replace(/^MRI IC\b.*/, "MRI IC 1");
    if (su === "OB") su = "OB 1";
    var num = su.match(/^(\d{1,3})(?:\s*\(.*\))?$/);
    if (num) {
      var n = parseInt(num[1], 10);
      if ((n >= 1 && n <= 36) || (n >= 51 && n <= 66) || (n >= 71 && n <= 76) || n >= 101) su = "OR " + n;
    }
    var em = su.match(/^ENDO\s*(\d)$/);
    if (em) su = "ENDO " + em[1];
    var epm = su.match(/^EP\s*(\d)$/);
    if (epm) su = "EP " + epm[1];
    if (ROOM_INDEX[su]) return ROOM_INDEX[su];
    var su2 = su.replace(/\s*[\(@].*$/, "").trim();
    if (ROOM_INDEX[su2]) return ROOM_INDEX[su2];
    return null;
  }

  function nameKey(n) {
    return String(n || "").toLowerCase().replace(/[^a-z]/g, "");
  }

  function stillInHouse(shift) {
    var ENDH = { D: 15, d: 17, M: 18, S: 19, Q: 21, W: 23, E: 23, N: 31, t: 31 };
    var letter = coreShift(shift);
    var end = ENDH[letter];
    if (end == null) return true;
    if (isEarlyShift(shift)) end -= 1;
    return new Date().getHours() < end;
  }

  function parseWeekday(cells) {
    var date = excelDate(cell(cells, "H", 3));
    var ntR = parseStaff(cell(cells, "E", 2));
    var stR = parseStaff(cell(cells, "E", 3));
    var rooms = [];
    var unmatched = [];
    var closed = [];
    var onDeck = [];
    var filled = {};
    var mriIcN = 0;
    function addDeck(st, lastRoom, role) {
      if (!st || !st.name || st.closed) return;
      if (/^\d{3,4}\s*-\s*\d/.test(st.name) || /^2200/.test(st.name)) return;
      onDeck.push({
        name: st.name, shift: st.shift || "", kind: st.kind || "none",
        lastRoom: lastRoom || "", role: role || ""
      });
    }
    function addPair(roomRaw, staffRaw, src) {
      roomRaw = (roomRaw || "").trim();
      staffRaw = (staffRaw || "").trim();
      if (!roomRaw) return;
      if (skipRoomLabel(roomRaw)) return;
      if (!staffRaw) return;
      var locRoom = roomRaw;
      if (/^MRI IC\b/i.test(roomRaw)) {
        mriIcN += 1;
        locRoom = mriIcN === 1 ? "MRI IC 1" : "MRI IC 2";
      }
      var loc = normalizeRoom(locRoom);
      var st = parseStaff(staffRaw);
      if (!loc) {
        if (st && st.name) addDeck(st, roomRaw, "unplaced");
        unmatched.push({ room: roomRaw, staff: staffRaw });
        return;
      }
      if (!st) return;
      var key = loc.cat + "|" + loc.room;
      if (filled[key] && !st.closed) {
        addDeck(st, loc.room, src === "cd" ? "offsite" : "extra");
        return;
      }
      var rec = {
        cat: loc.cat, room: loc.room, shift: st.shift || "", name: st.name || "",
        closed: !!st.closed, kind: st.kind || "none", early: !!st.early, orient: !!st.orient
      };
      filled[key] = rec;
      rooms.push(rec);
      if (rec.closed) closed.push({ cat: loc.cat, room: loc.room });
    }
    for (var r = 4; r <= 44; r++) {
      addPair(cell(cells, "A", r), cell(cells, "B", r), "ab");
      addPair(cell(cells, "E", r), cell(cells, "F", r), "ef");
    }
    for (r = 4; r <= 44; r++) {
      addPair(cell(cells, "C", r), cell(cells, "D", r), "cd");
    }
    function addLabelled(raw, lastRoom, role) {
      var st = parseStaff(raw);
      if (st && st.name) addDeck(st, lastRoom, role);
    }
    for (r = 5; r <= 43; r++) {
      var g = cell(cells, "G", r);
      var h = cell(cells, "H", r);
      if (/^MIDNIGHT/i.test(g)) addLabelled(h, g, "midnight");
      else if (/^#\d/.test(g)) { /* late stay = already assigned; can stay 2 hrs past shift */ }
      else if (/^(CV|CCS|3N|2N|ENDO|ST 1|ST 2|STE 1|STE 2)$/i.test(g)) addLabelled(h, g, "breaker");
      else if (/^(MN CALL|HEART CALL)$/i.test(g)) {
        var who = h && !/^\d{3,4}/.test(h) ? h : cell(cells, "G", r + 1);
        if (who && !/^(MN CALL|HEART CALL|NT BREAKERS|ST BREAKERS|LATE STAY)/i.test(who)) {
          addLabelled(who, g, "call");
        }
      }
    }
    addLabelled(cell(cells, "D", 28), "Resource", "resource");
    var placed = {};
    rooms.forEach(function (x) {
      if (x.name && !x.closed) {
        placed[nameKey(x.name)] = 1;
        placed[nameKey(lastName(x.name))] = 1;
      }
    });
    onDeck = onDeck.filter(function (p) {
      return p.name && !placed[nameKey(p.name)] && !placed[nameKey(lastName(p.name))];
    });
    var seen = {};
    onDeck = onDeck.filter(function (p) {
      var k = nameKey(p.name);
      if (seen[k]) return false;
      seen[k] = 1;
      return true;
    });
    var pos = {};
    for (r = 36; r <= 41; r++) {
      var lab = cell(cells, "G", r);
      var val = cell(cells, "H", r);
      var blob = lab + " " + val;
      function take(key, re) {
        var mm = blob.match(re);
        if (mm) pos[key] = parseInt(mm[1], 10);
      }
      take("1530-1730", /1530-1730\s*=?\s*(\d+)/);
      take("1730-1930", /1730-1930\s*=?\s*(\d+)/);
      take("1930-2130", /1930-2130\s*=?\s*(\d+)/);
      take("2130-2300", /2130-2300\s*=?\s*(\d+)/);
      take("2300-0700", /2300-0700\s*=\s*(\d+)/);
      if (/1530/.test(lab) && !pos["1530-1730"]) {
        var mm = val.match(/(\d+)/);
        if (mm) pos["1530-1730"] = parseInt(mm[1], 10);
      }
    }
    return {
      kind: "weekday", date: date,
      runners: { nt: ntR && ntR.name, st: stR && stR.name },
      rooms: rooms, closed: closed, unmatched: unmatched, pos: pos, people: [], onDeck: onDeck
    };
  }

  function parseWeekend(cells) {
    var people = [];
    var currentDate = null, currentTower = "";
    var eveningTotals = [];
    var r;
    for (r = 1; r <= 60; r++) {
      var a = cell(cells, "A", r), b = cell(cells, "B", r), c = cell(cells, "C", r);
      var e = cell(cells, "E", r), f = cell(cells, "F", r);
      var dt = excelDate(f) || excelDate(e);
      if (dt && /North Tower/i.test(b)) { currentDate = dt; currentTower = "NT"; continue; }
      if (/South Tower/i.test(b)) { currentTower = "ST"; continue; }
      if (/^evenings$/i.test(e)) {
        for (var k = 1; k <= 8; k++) {
          var t = cell(cells, "E", r + k);
          if (/^total:/i.test(t)) {
            var tm = t.match(/(\d+)/);
            if (tm) eveningTotals.push(parseInt(tm[1], 10));
            break;
          }
        }
      }
      if (/^shift$/i.test(a)) continue;
      if (/^(check hemacue|trauma rm)/i.test(a)) continue;
      if (!a || !/^(?:o\/)?\*?[DdMmSsQqWwEeNnTt]/i.test(a)) continue;
      if (!b) continue;
      var stA = parseStaff(a + " " + b);
      if (!stA || !stA.name) continue;
      var loc = c ? normalizeRoom(c) : null;
      people.push({
        date: currentDate, tower: currentTower, shift: stA.shift, name: stA.name,
        assign: c, cat: loc && loc.cat, room: loc && loc.room,
        closed: false, kind: stA.kind, early: stA.early, orient: stA.orient,
        role: /BR/i.test(c || "") ? "breaker" : ""
      });
    }
    var rooms = [];
    var unmatched = [];
    people.forEach(function (p) {
      if (p.cat && p.room) {
        rooms.push({ cat: p.cat, room: p.room, shift: p.shift, name: p.name, closed: false, kind: p.kind });
      } else if (p.assign && !/^BR/i.test(p.assign)) {
        unmatched.push({ room: p.assign, staff: p.name });
      }
    });
    var dates = [];
    people.forEach(function (p) { if (p.date && dates.indexOf(p.date) < 0) dates.push(p.date); });
    return {
      kind: "weekend", date: dates[0] || null, dates: dates, runners: {},
      rooms: rooms, closed: [], unmatched: unmatched, people: people,
      pos: { evening: eveningTotals.length ? eveningTotals[0] : 4, "2300-0700": 4 },
      onDeck: people.filter(function (p) { return p.name && !p.room; }).map(function (p) {
        return { name: p.name, shift: p.shift, kind: p.kind, lastRoom: p.assign || p.role || "", role: p.role || "float" };
      })
    };
  }

  async function parseAssignmentWorkbook(arrayBuffer, fileName) {
    buildRoomIndex();
    var zip = await unzip(arrayBuffer);
    var ss = loadSharedStrings(zip["xl/sharedStrings.xml"] ? xmlText(zip["xl/sharedStrings.xml"]) : "");
    var sheet = zip["xl/worksheets/sheet1.xml"];
    if (!sheet) throw new Error("No sheet1 in workbook");
    var cells = loadCells(xmlText(sheet), ss);
    var blob = Object.keys(cells).map(function (k) { return cells[k]; }).join(" ").toUpperCase();
    var out;
    if (blob.indexOf("CRNA SCHEDULE") >= 0 && blob.indexOf("NORTH TOWER") >= 0) out = parseWeekday(cells);
    else out = parseWeekend(cells);
    out.file = fileName || "";
    out.onDeck = out.onDeck || [];
    out.late = (out.rooms || []).filter(function (x) { return x.kind === "late" && !x.closed && x.name; });
    out.dinner = (out.rooms || []).filter(function (x) { return x.kind === "dinner" && !x.closed && x.name; });
    if (out.kind === "weekend" && out.people) {
      out.late = out.people.filter(function (x) { return x.kind === "late"; });
      out.dinner = out.people.filter(function (x) { return x.kind === "dinner"; });
    }
    return out;
  }

  function shiftPillHtml(shift) {
    var sh = String(shift || "");
    if (!sh || /^\d/.test(sh)) return "";
    var extra = " ";
    if (sh === "Dr") extra += "dr";
    if (sh.length > 2) extra += " wide";
    return '<span class="shift-pill' + extra + '">' + sh + "</span>";
  }

  function deckTag(p) {
    if (p.role === "freed" && p.lastRoom) return "last " + p.lastRoom;
    if (p.role === "breaker") return "breaker";
    if (p.role === "call") return "call";
    if (p.role === "midnight") return "night";
    return "";
  }

  function staffChipHtml(catId, room) {
    var rec = g.roomStaff && g.roomStaff[catId] && g.roomStaff[catId][room];
    if (!rec || rec.closed || !rec.name) return "";
    var nm = lastName(rec.name);
    var size = nm.length > 11 ? " tiny" : nm.length > 8 ? " long" : "";
    return '<span class="staff-line">' + shiftPillHtml(rec.shift) + '<span class="staff-name' + size + '">' + nm + "</span></span>";
  }

  function applyAssignmentResult(res) {
    if (!g.roomStaff) g.roomStaff = {};
    var listed = {};
    (res.rooms || []).forEach(function (r) {
      if (!r.cat || !r.room) return;
      listed[r.cat + "|" + r.room] = r;
    });
    if (typeof CATEGORIES !== "undefined") {
      CATEGORIES.forEach(function (c) {
        if (!g.roomStaff[c.id]) g.roomStaff[c.id] = {};
        (c.rooms || []).forEach(function (room) {
          var r = listed[c.id + "|" + room];
          var active = r && r.name && !r.closed;
          if (active) {
            g.roomStaff[c.id][room] = {
              name: r.name, shift: r.shift, kind: r.kind, closed: false,
              early: !!r.early, orient: !!r.orient
            };
            if (typeof catEditState !== "undefined") catEditState[c.id].deletedRooms.delete(room);
          } else {
            g.roomStaff[c.id][room] = { name: "", shift: "", kind: "none", closed: true };
            if (typeof catEditState !== "undefined") catEditState[c.id].deletedRooms.add(room);
          }
        });
      });
    } else {
      (res.rooms || []).forEach(function (r) {
        if (!r.cat || !r.room) return;
        if (!g.roomStaff[r.cat]) g.roomStaff[r.cat] = {};
        g.roomStaff[r.cat][r.room] = { name: r.name, shift: r.shift, kind: r.kind, closed: r.closed };
      });
    }
    g.onDeck = (res.onDeck || []).map(function (p) {
      return { name: p.name, shift: p.shift || "", kind: p.kind || "none", lastRoom: p.lastRoom || "", role: p.role || "" };
    });
    g.assignmentMeta = {
      date: res.date, kind: res.kind, file: res.file, pos: res.pos,
      lateN: (res.late || []).length, dinnerN: (res.dinner || []).length,
      roomsN: (res.rooms || []).length, closedN: (res.closed || []).length,
      unmatchedN: (res.unmatched || []).length, deckN: g.onDeck.length, appliedAt: Date.now()
    };
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try {
      if (typeof buildBoard === "function" && !document.getElementById("card-nt")) buildBoard();
      else if (typeof refreshBoard === "function") refreshBoard();
    } catch (e) {}
    renderLateBoardBar();
    renderOnDeck();
  }

  function renderLateBoardBar() {
    var bar = document.getElementById("late-board-bar");
    if (!bar) {
      bar = document.createElement("div");
      bar.id = "late-board-bar";
      var host = document.getElementById("evening-runner-bar") || document.getElementById("edit-active-bar");
      if (host && host.parentNode) host.parentNode.insertBefore(bar, host.nextSibling);
      else return;
    }
    var evening = typeof currentWindow === "number" && currentWindow === 3;
    var weekend = typeof isWeekend === "function" && isWeekend();
    var show = (evening || weekend) && (g.assignmentMeta || (g.roomStaff && Object.keys(g.roomStaff).length));
    bar.style.display = show ? "block" : "none";
    if (!show) return;
    var late = [];
    var dinner = [];
    var posOpen = 0;
    Object.keys(g.roomStaff || {}).forEach(function (cat) {
      Object.keys(g.roomStaff[cat] || {}).forEach(function (room) {
        var rec = g.roomStaff[cat][room];
        if (!rec || rec.closed) return;
        var deleted = typeof catEditState !== "undefined" && catEditState[cat] && catEditState[cat].deletedRooms.has(room);
        if (deleted) return;
        posOpen++;
        if (rec.kind === "late") late.push(rec.shift + " " + lastName(rec.name));
        if (rec.kind === "dinner") dinner.push(rec.shift + " " + lastName(rec.name));
      });
    });
    var posNote = "";
    if (g.assignmentMeta && g.assignmentMeta.pos) {
      if (weekend && g.assignmentMeta.pos.evening != null) posNote = "Sheet evening POS " + g.assignmentMeta.pos.evening;
      else if (g.assignmentMeta.pos["1530-1730"] != null) posNote = "Sheet 15:30 POS " + g.assignmentMeta.pos["1530-1730"];
    }
    bar.innerHTML =
      '<div class="late-board-title">Late board</div>' +
      '<div class="late-board-row"><strong>POS open</strong> ' + posOpen + (posNote ? " · " + posNote : "") + "</div>" +
      '<div class="late-board-row"><strong>Late (M+S)</strong> ' + late.length + (late.length ? " — " + late.slice(0, 12).join(", ") : "") + "</div>" +
      '<div class="late-board-row"><strong>Dinner (Q+W+E)</strong> ' + dinner.length + (dinner.length ? " — " + dinner.slice(0, 12).join(", ") : "") + "</div>";
  }

  function renderOnDeck() {
    var board = document.getElementById("board");
    if (!board) return;
    var card = document.getElementById("card-ondeck");
    var list = g.onDeck || [];
    if (!list.length) {
      if (card) card.remove();
      return;
    }
    if (!card) {
      card = document.createElement("div");
      card.className = "cat-card ondeck-card";
      card.id = "card-ondeck";
      board.insertBefore(card, board.firstChild);
    }
    var afternoon = (typeof currentWindow === "number" && currentWindow >= 2) || new Date().getHours() >= 15;
    var sel = g.selectedDeck || null;
    var chips = list.map(function (p) {
      var key = nameKey(p.name);
      var gone = afternoon && !stillInHouse(p.shift);
      var tag = deckTag(p);
      var last = tag ? '<span class="staff-last">' + tag + "</span>" : "";
      var on = sel === key ? " selected" : "";
      return '<button type="button" class="ondeck-chip' + (gone ? " out" : "") + on + '" data-deck="' + key + '">' +
        shiftPillHtml(p.shift) + '<span class="staff-name">' + lastName(p.name) + "</span>" + last +
        (gone ? '<span class="staff-last">out</span>' : "") + "</button>";
    }).join("");
    var hint = sel
      ? '<div class="ondeck-hint">Tap a room to place them · tap the chip again to cancel</div>'
      : '<div class="ondeck-hint">Tap someone, then tap a room</div>';
    card.innerHTML =
      '<div class="cat-header-row"><div class="cup-indicator">☕</div><div class="cat-info">' +
      '<div class="cat-name">On deck</div>' +
      '<div class="cat-full-name">Not in a room · tap to place</div></div>' +
      '<div class="cat-progress-label"><div class="cat-pct">' + list.length + '</div>' +
      '<div class="cat-count">available</div></div></div>' +
      hint +
      '<div class="ondeck-grid">' + chips + "</div>";
    card.querySelectorAll("[data-deck]").forEach(function (el) {
      el.onclick = function (ev) {
        ev.stopPropagation();
        var k = el.getAttribute("data-deck");
        g.selectedDeck = g.selectedDeck === k ? null : k;
        document.body.classList.toggle("assigning", !!g.selectedDeck);
        renderOnDeck();
      };
    });
    document.body.classList.toggle("assigning", !!g.selectedDeck);
  }

  function pushOnDeckFromRoom(catId, room) {
    var rec = g.roomStaff && g.roomStaff[catId] && g.roomStaff[catId][room];
    if (!rec || !rec.name) return;
    g.onDeck = g.onDeck || [];
    var k = nameKey(rec.name);
    g.onDeck = g.onDeck.filter(function (p) { return nameKey(p.name) !== k; });
    g.onDeck.unshift({ name: rec.name, shift: rec.shift || "", kind: rec.kind || "none", lastRoom: room, role: "freed" });
    delete g.roomStaff[catId][room];
    renderOnDeck();
  }

  function assignSelectedTo(catId, room) {
    var key = g.selectedDeck;
    if (!key || !catId || !room) return false;
    var list = g.onDeck || [];
    var idx = -1;
    for (var i = 0; i < list.length; i++) {
      if (nameKey(list[i].name) === key) { idx = i; break; }
    }
    if (idx < 0) { g.selectedDeck = null; return false; }
    var person = list[idx];
    g.roomStaff = g.roomStaff || {};
    g.roomStaff[catId] = g.roomStaff[catId] || {};
    var occ = g.roomStaff[catId][room];
    list.splice(idx, 1);
    if (occ && occ.name) {
      list.unshift({ name: occ.name, shift: occ.shift || "", kind: occ.kind || "none", lastRoom: room, role: "freed" });
    }
    g.roomStaff[catId][room] = { name: person.name, shift: person.shift || "", kind: person.kind || "none", closed: false };
    if (typeof catEditState !== "undefined" && catEditState[catId]) {
      catEditState[catId].deletedRooms.delete(room);
    }
    g.selectedDeck = null;
    document.body.classList.remove("assigning");
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    renderOnDeck();
    try { if (typeof showToast === "function") showToast(lastName(person.name) + " → " + room); } catch (e) {}
    return true;
  }

  function wrapDeactivate() {
    if (typeof toggleDeleteRoom === "function" && !g._deckWrapped) {
      g._deckWrapped = true;
      var orig = toggleDeleteRoom;
      window.toggleDeleteRoom = function (catId, room) {
        if (g.selectedDeck && assignSelectedTo(catId, room)) return;
        var was = typeof catEditState !== "undefined" && catEditState[catId] && catEditState[catId].deletedRooms.has(room);
        orig.apply(this, arguments);
        var now = typeof catEditState !== "undefined" && catEditState[catId] && catEditState[catId].deletedRooms.has(room);
        if (!was && now) pushOnDeckFromRoom(catId, room);
        else renderOnDeck();
      };
    }
    if (typeof toggleRoom === "function" && !g._toggleRoomWrapped) {
      g._toggleRoomWrapped = true;
      var origToggle = toggleRoom;
      window.toggleRoom = function (catId, room) {
        if (g.selectedDeck && assignSelectedTo(catId, room)) return;
        origToggle.apply(this, arguments);
      };
    }
    if (typeof setWindow === "function" && !g._setWindowWrapped) {
      g._setWindowWrapped = true;
      var origSet = setWindow;
      window.setWindow = function () {
        origSet.apply(this, arguments);
        renderOnDeck();
      };
    }
  }

  function showAssignSummary(res, file) {
    var existing = document.getElementById("assign-upload-overlay");
    if (existing) existing.remove();
    var ov = document.createElement("div");
    ov.id = "assign-upload-overlay";
    ov.className = "assign-upload-overlay";
    var today = typeof todayStr === "function" ? todayStr() : "";
    var dateWarn = res.date && today && res.date !== today
      ? '<div class="assign-warn">Sheet date is ' + res.date + " (today is " + today + "). Apply anyway?</div>"
      : "";
    var um = (res.unmatched || []).slice(0, 6).map(function (u) {
      return "<li>" + (u.room || "") + " — " + (u.staff || "") + "</li>";
    }).join("");
    var hideN = 0;
    if (typeof CATEGORIES !== "undefined") {
      var named = {};
      (res.rooms || []).forEach(function (r) {
        if (r && r.cat && r.room && r.name && !r.closed) named[r.cat + "|" + r.room] = 1;
      });
      CATEGORIES.forEach(function (c) {
        (c.rooms || []).forEach(function (room) {
          if (!named[c.id + "|" + room]) hideN++;
        });
      });
    }
    ov.innerHTML =
      '<div class="assign-upload-card">' +
      "<h3>Assignment sheet</h3>" +
      "<p class='assign-meta'>" + (res.kind || "") + (res.date ? " · " + res.date : "") + (file ? " · " + file.name : "") + "</p>" +
      dateWarn +
      "<ul class='assign-stats'>" +
      "<li>" + (res.rooms || []).length + " rooms with a name</li>" +
      "<li>" + (res.closed || []).length + " marked CLOSED</li>" +
      (hideN ? "<li>" + hideN + " not on sheet (will hide)</li>" : "") +
      "<li>" + (res.late || []).length + " late (M+S)</li>" +
      "<li>" + (res.dinner || []).length + " dinner (Q+W+E)</li>" +
      "<li>" + ((res.onDeck || []).length) + " on deck (not in a room)</li>" +
      (res.unmatched && res.unmatched.length ? "<li>" + res.unmatched.length + " unmatched labels</li>" : "") +
      "</ul>" +
      (um ? "<ul class='assign-unmatched'>" + um + "</ul>" : "") +
      '<div class="assign-actions">' +
      '<button type="button" class="assign-cancel">Cancel</button>' +
      '<button type="button" class="assign-apply">Apply to board</button>' +
      "</div></div>";
    document.body.appendChild(ov);
    ov.querySelector(".assign-cancel").onclick = function () { ov.remove(); };
    ov.querySelector(".assign-apply").onclick = function () {
      applyAssignmentResult(res);
      ov.remove();
      try { if (typeof showToast === "function") showToast("Assignment sheet applied"); } catch (e) {}
    };
  }

  async function handleAssignmentFile(file) {
    if (!file) return;
    try {
      var buf = await file.arrayBuffer();
      var res = await parseAssignmentWorkbook(buf, file.name);
      showAssignSummary(res, file);
    } catch (e) {
      try { if (typeof showToast === "function") showToast("Could not read that sheet"); } catch (x) {}
      console.warn(e);
    }
  }

  function ensureUploadUi() {
    if (document.getElementById("assign-file-input")) return;
    var input = document.createElement("input");
    input.type = "file";
    input.id = "assign-file-input";
    input.accept = ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    input.style.display = "none";
    input.addEventListener("change", function () {
      var f = input.files && input.files[0];
      input.value = "";
      handleAssignmentFile(f);
    });
    document.body.appendChild(input);
    if (!document.getElementById("assign-upload-css")) {
      var st = document.createElement("style");
      st.id = "assign-upload-css";
      st.textContent =
        ".staff-chip{display:none;}" +
        ".staff-line{display:flex;align-items:center;justify-content:center;gap:3px;margin-top:2px;white-space:nowrap;overflow:hidden;max-width:100%;}" +
        ".shift-pill{flex:0 0 auto;font-size:8px;font-weight:800;letter-spacing:.02em;padding:1px 4px;border-radius:5px;background:rgba(122,78,45,.14);color:#7A4E2D;line-height:1.2;}" +
        ".shift-pill.dr{background:transparent;border:1px solid rgba(122,78,45,.3);font-weight:700;}" +
        ".shift-pill.wide{font-size:7px;padding:1px 3px;letter-spacing:0;}" +
        ".ondeck-chip .shift-pill{position:static;}" +
        ".staff-name{font-size:10px;font-weight:600;color:#1E0E04;letter-spacing:-0.03em;}" +
        ".staff-name.long{font-size:9px;}" +
        ".staff-name.tiny{font-size:8px;letter-spacing:-0.04em;}" +
        ".staff-last{font-size:10px;font-weight:500;color:#9A6A38;}" +
        ".staff-last::before{content:'·';margin:0 3px;color:rgba(122,78,45,.45);}" +
        ".ondeck-hint{padding:0 12px 6px;font-size:11px;color:#9A6A38;}" +
        "body.assigning .room-btn{box-shadow:inset 0 0 0 1.5px rgba(122,78,45,.35);}" +
        "#late-board-bar{display:none;padding:8px 14px 10px;background:#F7F0E6;border-bottom:1px solid rgba(160,98,42,.15);font-size:12px;color:#1E0E04;}" +
        "#late-board-bar .late-board-title{font-weight:800;font-size:11px;letter-spacing:.04em;text-transform:uppercase;color:#7A4E2D;margin-bottom:4px;}" +
        "#late-board-bar .late-board-row{margin:2px 0;line-height:1.35;}" +
        ".assign-upload-overlay{position:fixed;inset:0;background:rgba(30,14,4,.55);z-index:800;display:flex;align-items:center;justify-content:center;padding:18px;}" +
        ".assign-upload-card{background:#fff;border-radius:16px;padding:18px 16px;max-width:360px;width:100%;color:#1E0E04;}" +
        ".assign-upload-card h3{margin:0 0 6px;font-size:18px;}" +
        ".assign-meta{font-size:12px;color:#7A4E2D;margin:0 0 8px;}" +
        ".assign-warn{background:#FDEBD0;padding:8px;border-radius:8px;font-size:12px;margin-bottom:8px;}" +
        ".assign-stats{margin:0 0 8px;padding-left:18px;font-size:13px;}" +
        ".assign-unmatched{font-size:11px;color:#6b5344;max-height:90px;overflow:auto;}" +
        ".assign-actions{display:flex;gap:8px;margin-top:12px;}" +
        ".assign-actions button{flex:1;padding:12px;border:none;border-radius:10px;font-weight:700;}" +
        ".assign-cancel{background:#eee;color:#1E0E04;-webkit-appearance:none;appearance:none;}" +
        ".assign-apply{background:linear-gradient(135deg,#7A4E2D,#9A6A38);color:#fff;}" +
        "#edit-active-bar .assign-upload-btn{margin-top:6px;width:100%;font-size:11px;font-weight:700;padding:7px;border-radius:7px;border:1px solid rgba(122,78,45,.35);background:#fff;color:#7A4E2D;cursor:pointer;}" +
        ".ondeck-card{border:1px dashed rgba(122,78,45,.28);}" +
        ".ondeck-grid{display:flex;flex-wrap:wrap;gap:6px;padding:4px 12px 12px;}" +
        ".ondeck-chip{display:inline-flex;align-items:center;gap:3px;white-space:nowrap;max-width:100%;background:#FDF6EC;border:1px solid rgba(160,98,42,.22);border-radius:8px;padding:5px 8px;cursor:pointer;-webkit-appearance:none;appearance:none;font:inherit;color:inherit;}" +
        ".ondeck-chip.selected{border-color:#7A4E2D;background:#F5E6D0;box-shadow:inset 0 0 0 1px #7A4E2D;}" +
        ".ondeck-chip.out{opacity:.42;}";
      document.head.appendChild(st);
    }
  }

  function mountUploadButton() {
    ensureUploadUi();
    var bar = document.getElementById("edit-active-bar");
    if (!bar || bar.querySelector(".assign-upload-btn")) return;
    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "assign-upload-btn";
    btn.textContent = "Upload CRNA assignment sheet";
    btn.onclick = function () {
      var el = document.getElementById("assign-file-input");
      if (el) el.click();
    };
    bar.appendChild(btn);
  }

  g.roomStaff = g.roomStaff || {};
  g.onDeck = g.onDeck || [];
  g.assignmentMeta = g.assignmentMeta || null;
  g.parseAssignmentWorkbook = parseAssignmentWorkbook;
  g.applyAssignmentResult = applyAssignmentResult;
  g.handleAssignmentFile = handleAssignmentFile;
  g.staffChipHtml = staffChipHtml;
  g.renderLateBoardBar = renderLateBoardBar;
  g.renderOnDeck = renderOnDeck;
  g.mountUploadButton = mountUploadButton;
  g.ensureUploadUi = ensureUploadUi;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () { ensureUploadUi(); wrapDeactivate(); });
  } else { ensureUploadUi(); wrapDeactivate(); }
  setTimeout(wrapDeactivate, 500);
})(window);
