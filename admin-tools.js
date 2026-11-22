/* Trang quản trị – mô-đun công cụ: danh sách người lái đò, in thẻ QR & poster, báo cáo Excel.
 * Thay thế toàn bộ công cụ Python: mọi thao tác thực hiện trên trình duyệt bằng tài khoản quản trị. */
(function () {
  'use strict';
  var T = window.TA;
  if (!T) return;
  var $ = T.$, esc = T.esc, api = T.api;
  var LIB = {
    qr: { url: 'https://cdn.jsdelivr.net/npm/qrcode-generator@1.4.4/qrcode.js', sri: 'sha384-8FWZA6BGMXhsfO+BLtrJK0We6gg5o1JyO8xQm6peWDEUs17ACA5ziE/NIAkl9z2k', test: function () { return window.qrcode; } },
    xlsx: { url: 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.core.min.js', sri: 'sha384-+Vs0lbfRt7mOFk7nllUwVGaNHhoqQ8RjqUF8Dl2CrJxmL5CfKbQPdTymQkQym+bK', test: function () { return window.XLSX; } }
  };
  var loading = {};
  function need(name) {
    var L = LIB[name];
    if (L.test()) return Promise.resolve(L.test());
    if (loading[name]) return loading[name];
    loading[name] = new Promise(function (ok, fail) {
      var s = document.createElement('script');
      s.src = L.url; s.integrity = L.sri; s.crossOrigin = 'anonymous';
      s.onload = function () { L.test() ? ok(L.test()) : fail(new Error('Không tải được thư viện ' + name)); };
      s.onerror = function () { delete loading[name]; fail(new Error('Không tải được thư viện ' + name + '. Kiểm tra kết nối mạng.')); };
      document.head.appendChild(s);
    });
    return loading[name];
  }
  var D = { list: null, ver: null, base: '', q: '', team: '', st: '' };   // trạng thái danh sách
  var TEAM_SORT = function (a, b) { return String(a).localeCompare(String(b), 'vi', { numeric: true }); };
  function siteBase() { return D.base || (location.origin + location.pathname.replace(/[^/]*$/, '')); }
  function linkOf(id) { return siteBase() + '?id=' + encodeURIComponent(id); }
  function dirLink() { return siteBase() + 'danhba.html'; }
  function busy(el, on) { if (el) { el.disabled = on; } }
  function loadDrivers(force) {
    if (D.list && !force) return Promise.resolve(D);
    return api({ action: 'adm_drivers' }).then(function (j) { D.list = j.drivers; D.ver = j.ver; D.base = j.siteUrl || ''; return D; });
  }
  function saveDrivers(list, msg) {
    return api({ action: 'adm_drivers_save', drivers: list, ver: D.ver }).then(function (j) {
      D.list = j.drivers; D.ver = j.ver; D.base = j.siteUrl || D.base;
      T.toast(msg || ('Đã lưu: ' + j.count + ' người lái đò (' + j.active + ' đang hoạt động).'));
      return j;
    });
  }

  /* =============================== MÃ QR =============================== */
  function qrMatrix(text) {
    var q = window.qrcode(0, 'Q');            // mức sửa lỗi Q: chịu được ~25% bề mặt bị bẩn, xước
    q.addData(text); q.make();
    var n = q.getModuleCount(), m = [];
    for (var r = 0; r < n; r++) { var row = []; for (var c = 0; c < n; c++) row.push(q.isDark(r, c)); m.push(row); }
    return m;
  }
  function qrSvg(text, color, quiet) {
    var m = qrMatrix(text), n = m.length, z = quiet == null ? 0 : quiet, d = '';
    for (var r = 0; r < n; r++) {
      for (var c = 0; c < n; c++) {
        if (!m[r][c]) continue;
        var k = c; while (k < n && m[r][k]) k++;
        d += 'M' + (c + z) + ' ' + (r + z) + 'h' + (k - c) + 'v1h' + (c - k) + 'z';
        c = k;
      }
    }
    var s = n + 2 * z;
    return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + s + ' ' + s + '" shape-rendering="crispEdges">' +
      (z ? '<rect width="' + s + '" height="' + s + '" fill="#fff"/>' : '') + '<path d="' + d + '" fill="' + (color || '#1f4d3a') + '"/></svg>';
  }
  function qrPng(text, px) {
    var m = qrMatrix(text), n = m.length, z = 4, cell = Math.max(4, Math.floor((px || 800) / (n + 2 * z)));
    var cv = document.createElement('canvas'); cv.width = cv.height = (n + 2 * z) * cell;
    var g = cv.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, cv.width, cv.height); g.fillStyle = '#000';
    for (var r = 0; r < n; r++) for (var c = 0; c < n; c++) if (m[r][c]) g.fillRect((c + z) * cell, (r + z) * cell, cell, cell);
    return cv.toDataURL('image/png');
  }
  function download(name, href) { var a = document.createElement('a'); a.href = href; a.download = name; document.body.appendChild(a); a.click(); a.remove(); }
  function showQr(title, sub, url, file) {
    need('qr').then(function () {
      T.modal('<h3>' + esc(title) + '</h3><p>' + esc(sub) + '</p><div style="width:240px;margin:6px auto 10px">' + qrSvg(url, '#000', 2) + '</div>' +
        '<div class="row" style="gap:6px;margin-bottom:8px"><input class="i" id="qUrl" readonly value="' + esc(url) + '" style="font-size:13px;flex:1;min-width:0"><button class="btn sm" id="qCopy">Sao chép</button></div>' +
        '<div class="row" style="justify-content:flex-end"><button class="btn" data-x>Đóng</button><button class="btn" id="qSvg">Tải SVG</button><button class="btn pri" id="qPng">Tải ảnh PNG</button></div>');
      $('qCopy').onclick = function () { copy(url); };
      $('qPng').onclick = function () { download(file + '.png', qrPng(url, 1200)); };
      $('qSvg').onclick = function () { download(file + '.svg', 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(qrSvg(url, '#000', 4))); };
    }).catch(function (e) { T.toast(e.message); });
  }
  function copy(t) {
    (navigator.clipboard ? navigator.clipboard.writeText(t) : Promise.reject()).then(function () { T.toast('Đã sao chép.'); },
      function () { var i = $('qUrl'); if (i) { i.select(); document.execCommand('copy'); T.toast('Đã sao chép.'); } });
  }

  /* =============================== NGƯỜI LÁI ĐÒ =============================== */
  function paneDr() {
    var p = $('pDr'); p.innerHTML = '<div class="loading"><span class="spin"></span>Đang tải danh sách…</div>';
    loadDrivers(true).then(renderDr).catch(function (e) { p.innerHTML = '<div class="err">' + esc(e.message) + '</div>'; });
  }
  function renderDr() {
    var adm = T.can('admin'), L = D.list;
    var teams = []; L.forEach(function (d) { if (d.team && teams.indexOf(d.team) < 0) teams.push(d.team); }); teams.sort(TEAM_SORT);
    var act = L.filter(function (d) { return d.status !== 'Nghỉ'; }).length;
    $('pDr').innerHTML =
      '<div class="kpis"><div class="kpi"><div class="l">Tổng số người lái đò</div><div class="v">' + L.length + '</div></div>' +
      '<div class="kpi"><div class="l">Đang hoạt động</div><div class="v">' + act + '</div></div>' +
      '<div class="kpi"><div class="l">Nghỉ</div><div class="v">' + (L.length - act) + '</div></div>' +
      '<div class="kpi"><div class="l">Số tổ</div><div class="v">' + teams.length + '</div></div></div>' +
      '<div class="card"><div class="row">' +
        '<input class="i" id="dQ" placeholder="Tìm mã, tên, số đò" style="max-width:220px;padding:8px 10px;font-size:14px" value="' + esc(D.q) + '">' +
        '<select class="i" id="dTeam" style="max-width:160px;padding:8px 10px;font-size:14px"><option value="">Tất cả tổ</option>' + teams.map(function (t) { return '<option' + (D.team === t ? ' selected' : '') + '>' + esc(t) + '</option>'; }).join('') + '</select>' +
        '<select class="i" id="dSt" style="max-width:160px;padding:8px 10px;font-size:14px"><option value="">Mọi trạng thái</option><option' + (D.st === 'Hoạt động' ? ' selected' : '') + '>Hoạt động</option><option' + (D.st === 'Nghỉ' ? ' selected' : '') + '>Nghỉ</option></select>' +
        '<span class="sp"></span>' +
        (adm ? '<button class="btn sm pri" id="dAdd">+ Thêm người</button><button class="btn sm" id="dImp">Nhập từ Excel</button><button class="btn sm" id="dTpl">Tải file mẫu</button>' : '') +
        '<button class="btn sm" id="dExp">Xuất Excel</button><button class="btn sm" id="dDir">QR danh bạ</button>' +
      '</div></div>' +
      '<div class="card"><div class="muted" id="dCount" style="margin-bottom:6px"></div><div class="tw"><table><thead><tr><th>Mã số</th><th>Họ và tên</th><th>Số đò</th><th>Tổ</th><th>Trạng thái</th><th></th></tr></thead><tbody id="dBody"></tbody></table></div>' +
      (adm ? '<p class="muted" style="font-size:12px;margin:10px 0 0">Mỗi thay đổi được lưu ngay lên hệ thống và ghi nhật ký. Người nghỉ việc nên chuyển sang trạng thái <b>Nghỉ</b> thay vì xóa để giữ lịch sử đánh giá. Đổi mã số đồng nghĩa với cấp mã QR mới.</p>' : '') + '</div>';
    drawDrRows();
    $('dQ').oninput = function () { D.q = this.value; drawDrRows(); };
    $('dTeam').onchange = function () { D.team = this.value; drawDrRows(); };
    $('dSt').onchange = function () { D.st = this.value; drawDrRows(); };
    $('dExp').onclick = exportList;
    $('dDir').onclick = function () { showQr('Mã QR tổng (danh bạ)', 'Đặt tại quầy vé, bến xuất phát – du khách quét trước chuyến đi.', dirLink(), 'qr_danh_ba'); };
    if (adm) {
      $('dAdd').onclick = function () { driverForm(null); };
      $('dImp').onclick = importForm;
      $('dTpl').onclick = template;
    }
    $('dBody').onclick = function (e) {
      var b = e.target.closest('[data-a]'); if (!b) return;
      var id = b.closest('tr').dataset.id, d = D.list.filter(function (x) { return x.id === id; })[0];
      if (b.dataset.a === 'qr') showQr(d.id + (d.name ? ' – ' + d.name : ''), 'Mã QR đánh giá riêng của người lái đò' + (d.boat ? ' – đò số ' + d.boat : '') + '.', linkOf(d.id), 'qr_' + d.id);
      if (b.dataset.a === 'ed') driverForm(d);
      if (b.dataset.a === 'tg') {
        var nl = D.list.map(function (x) { return x.id === id ? Object.assign({}, x, { status: x.status === 'Nghỉ' ? 'Hoạt động' : 'Nghỉ' }) : x; });
        b.disabled = true;
        saveDrivers(nl, d.id + (d.status === 'Nghỉ' ? ' đã hoạt động trở lại.' : ' đã chuyển sang Nghỉ.')).then(renderDr).catch(function (er) { T.toast(er.message); b.disabled = false; });
      }
      if (b.dataset.a === 'rm') T.confirmBox('Xóa ' + d.id + ' khỏi danh sách?', 'Người này sẽ không còn trong danh bạ, mã QR của họ ngừng hoạt động và các đánh giá cũ không còn được tổng hợp vào báo cáo. Nếu chỉ nghỉ việc, nên chọn "Cho nghỉ" thay vì xóa.', 'Xóa', function () {
        return saveDrivers(D.list.filter(function (x) { return x.id !== id; }), 'Đã xóa ' + d.id + '.').then(renderDr); });
    };
  }
  function filtered() {
    var q = T.fold(D.q);
    return D.list.filter(function (d) {
      return (!D.team || d.team === D.team) && (!D.st || d.status === D.st) && (!q || T.fold([d.id, d.name, d.boat].join(' ')).indexOf(q) >= 0);
    });
  }
  function drawDrRows() {
    var adm = T.can('admin'), rows = filtered();
    $('dCount').textContent = 'Hiển thị ' + rows.length + '/' + D.list.length + ' người';
    $('dBody').innerHTML = rows.length ? rows.map(function (d) {
      return '<tr data-id="' + esc(d.id) + '"><td><b>' + esc(d.id) + '</b></td><td>' + esc(d.name) + '</td><td>' + esc(d.boat) + '</td><td>' + esc(d.team) + '</td>' +
        '<td>' + (d.status === 'Nghỉ' ? '<span class="pill k">Nghỉ</span>' : '<span class="pill g">Hoạt động</span>') + '</td>' +
        '<td class="n"><button class="btn sm" data-a="qr">QR</button> ' + (adm ? '<button class="btn sm" data-a="ed">Sửa</button> <button class="btn sm" data-a="tg">' + (d.status === 'Nghỉ' ? 'Kích hoạt' : 'Cho nghỉ') + '</button> <button class="btn sm danger" data-a="rm">Xóa</button>' : '') + '</td></tr>';
    }).join('') : '<tr><td colspan="6" class="empty">Chưa có người lái đò nào' + (adm && !D.list.length ? ' – bấm "+ Thêm người" hoặc "Nhập từ Excel".' : '.') + '</td></tr>';
  }
  function cleanId(v) { return String(v == null ? '' : v).trim().toUpperCase(); }
  function driverForm(d) {
    var isNew = !d; d = d || { id: '', name: '', boat: '', team: '', status: 'Hoạt động' };
    var teams = []; D.list.forEach(function (x) { if (x.team && teams.indexOf(x.team) < 0) teams.push(x.team); }); teams.sort(TEAM_SORT);
    T.modal('<h3>' + (isNew ? 'Thêm người lái đò' : 'Sửa ' + esc(d.id)) + '</h3>' +
      '<label class="f">Mã số (in trên thẻ QR – chữ không dấu, số, dấu gạch ngang)</label><input class="i" id="fdId" value="' + esc(d.id) + '" placeholder="VD: TA-0006" autocapitalize="characters">' +
      (isNew ? '' : '<div class="muted" style="font-size:12px;margin-top:4px">Nếu đổi mã số, cần in lại thẻ QR cho người này.</div>') +
      '<label class="f">Họ và tên</label><input class="i" id="fdName" value="' + esc(d.name) + '">' +
      '<div class="row" style="gap:10px"><div style="flex:1"><label class="f">Số đò</label><input class="i" id="fdBoat" value="' + esc(d.boat) + '"></div>' +
      '<div style="flex:1"><label class="f">Tổ</label><input class="i" id="fdTeam" list="fdTeams" value="' + esc(d.team) + '"><datalist id="fdTeams">' + teams.map(function (t) { return '<option value="' + esc(t) + '">'; }).join('') + '</datalist></div></div>' +
      '<label class="f">Trạng thái</label><select class="i" id="fdSt"><option' + (d.status !== 'Nghỉ' ? ' selected' : '') + '>Hoạt động</option><option' + (d.status === 'Nghỉ' ? ' selected' : '') + '>Nghỉ</option></select>' +
      '<div class="err hidden" id="mErr"></div><div class="row" style="justify-content:flex-end;margin-top:14px"><button class="btn" data-x>Hủy</button><button class="btn pri" id="mOk">Lưu</button></div>');
    if (isNew) {   // gợi ý mã tiếp theo theo mẫu mã hiện có
      var max = 0, pre = 'TA-', w = 4;
      D.list.forEach(function (x) { var m = /^(.*?)(\d+)$/.exec(x.id); if (m && +m[2] >= max) { max = +m[2]; pre = m[1]; w = m[2].length; } });
      $('fdId').value = pre + String(max + 1).padStart(w, '0');
    }
    $('mOk').onclick = function () {
      var x = { id: cleanId($('fdId').value), name: $('fdName').value.trim(), boat: $('fdBoat').value.trim(), team: $('fdTeam').value.trim(), status: $('fdSt').value };
      if (!/^[A-Z0-9-]{1,32}$/.test(x.id)) { T.showErr('Mã số chỉ gồm chữ không dấu, số và dấu gạch ngang (tối đa 32 ký tự).'); return; }
      if (D.list.some(function (o) { return o.id === x.id && o.id !== d.id; })) { T.showErr('Mã số ' + x.id + ' đã có trong danh sách.'); return; }
      if (!x.name) { T.showErr('Vui lòng nhập họ tên.'); return; }
      var nl = isNew ? D.list.concat([x]) : D.list.map(function (o) { return o.id === d.id ? x : o; });
      var b = this; b.disabled = true;
      saveDrivers(nl, (isNew ? 'Đã thêm ' : 'Đã lưu ') + x.id + '.').then(function () { T.closeModal(); renderDr(); })
        .catch(function (e) { T.showErr(e.message); b.disabled = false; });
    };
  }

  /* ----- Nhập từ Excel / CSV (đọc ngay trên trình duyệt, không tải file lên máy chủ) ----- */
  var COLS = {
    id: ['ma so', 'ma', 'ma lai do', 'ma nguoi lai do', 'ma dinh danh', 'id'],
    name: ['ho va ten', 'ho ten', 'ten', 'ten nguoi lai do'],
    boat: ['so do', 'do so', 'so thuyen', 'bien so do'],
    team: ['to', 'doi', 'to/doi', 'to doi', 'ten to'],
    status: ['trang thai', 'tinh trang']
  };
  function norm(v) { return T.fold(String(v == null ? '' : v)).replace(/\s+/g, ' ').trim(); }
  function cellStr(v) { if (v == null) return ''; if (typeof v === 'number' && Math.floor(v) === v) return String(v); return String(v).trim(); }
  function parseSheet(rows) {
    var head = -1, col = {};
    for (var i = 0; i < Math.min(15, rows.length) && head < 0; i++) {
      var t = (rows[i] || []).map(norm), f = {};
      Object.keys(COLS).forEach(function (k) { t.forEach(function (x, j) { if (COLS[k].indexOf(x) >= 0 && !(k in f)) f[k] = j; }); });
      if ('id' in f) { head = i; col = f; }
    }
    if (head < 0) throw new Error('Không tìm thấy cột "Mã số" trong 15 dòng đầu. Hãy dùng file mẫu (nút "Tải file mẫu").');
    var out = [], err = [], warn = [], seen = {};
    for (var r = head + 1; r < rows.length; r++) {
      var row = rows[r] || [], get = function (k) { return k in col ? cellStr(row[col[k]]) : ''; };
      var x = { id: get('id').toUpperCase(), name: get('name'), boat: get('boat'), team: get('team'), status: norm(get('status')) === 'nghi' ? 'Nghỉ' : 'Hoạt động' };
      if (!x.id && !x.name && !x.boat && !x.team) continue;
      var ln = r + 1;
      if (!x.id) { err.push('Dòng ' + ln + ': thiếu mã số (' + (x.name || x.boat) + ')'); continue; }
      if (!/^[A-Z0-9-]{1,32}$/.test(x.id)) { err.push('Dòng ' + ln + ': mã số "' + x.id + '" chỉ được dùng chữ không dấu, số, dấu gạch ngang'); continue; }
      if (seen[x.id]) { err.push('Dòng ' + ln + ': mã số ' + x.id + ' trùng với dòng ' + seen[x.id]); continue; }
      seen[x.id] = ln;
      if (!x.name) warn.push('Dòng ' + ln + ' (' + x.id + '): chưa có họ tên');
      out.push(x);
    }
    return { list: out, err: err, warn: warn };
  }
  function importForm() {
    T.modal('<h3>Nhập danh sách từ Excel</h3><p>Chọn file Excel (.xlsx, .xls) hoặc .csv có các cột: <b>Mã số, Họ và tên, Số đò, Tổ, Trạng thái</b> (tên cột viết có dấu hay không dấu đều được). File chỉ được đọc trên máy này, không tải lên đâu cả.</p>' +
      '<input type="file" id="iFile" accept=".xlsx,.xls,.xlsm,.csv" class="i" style="padding:8px">' +
      '<div id="iRes" style="margin-top:10px"></div><div class="err hidden" id="mErr"></div>' +
      '<div class="row" style="justify-content:flex-end;margin-top:14px"><button class="btn" data-x>Hủy</button><button class="btn pri" id="mOk" disabled>Cập nhật lên hệ thống</button></div>');
    var parsed = null;
    $('iFile').onchange = function () {
      var f = this.files[0]; if (!f) return;
      $('iRes').innerHTML = '<div class="loading"><span class="spin"></span>Đang đọc file…</div>'; $('mOk').disabled = true; $('mErr').classList.add('hidden');
      var csv = /\.csv$/i.test(f.name);
      Promise.all([need('xlsx'), csv ? f.text() : f.arrayBuffer()]).then(function (a) {
        var wb = csv ? a[0].read(a[1].replace(/^\ufeff/, ''), { type: 'string' }) : a[0].read(new Uint8Array(a[1]), { type: 'array' });
        var ws = wb.Sheets.DanhSach || wb.Sheets[wb.SheetNames[0]];
        parsed = parseSheet(a[0].utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' }));
        var old = {}; D.list.forEach(function (d) { old[d.id] = d; });
        var add = 0, upd = 0, same = 0;
        parsed.list.forEach(function (x) { var o = old[x.id]; if (!o) add++; else if (o.name !== x.name || o.boat !== x.boat || o.team !== x.team || o.status !== x.status) upd++; else same++; });
        var inFile = {}; parsed.list.forEach(function (x) { inFile[x.id] = 1; });
        var miss = D.list.filter(function (d) { return !inFile[d.id]; }).length;
        $('iRes').innerHTML = parsed.err.length
          ? '<div class="err">File có ' + parsed.err.length + ' lỗi, cần sửa rồi chọn lại:<br>' + parsed.err.slice(0, 15).map(esc).join('<br>') + (parsed.err.length > 15 ? '<br>…' : '') + '</div>'
          : '<div class="card" style="margin:0"><b>' + parsed.list.length + ' người</b> trong file: thêm mới <b>' + add + '</b>, thay đổi <b>' + upd + '</b>, giữ nguyên ' + same + '.' +
            (parsed.warn.length ? '<div class="muted" style="font-size:12px;margin-top:6px">Lưu ý: ' + parsed.warn.slice(0, 8).map(esc).join('; ') + (parsed.warn.length > 8 ? '…' : '') + '</div>' : '') +
            (miss ? '<div style="margin-top:10px">Có <b>' + miss + '</b> người đang trên hệ thống nhưng không có trong file:<br>' +
              '<label class="chk" style="margin-top:6px"><input type="radio" name="iMode" value="merge" checked> Giữ nguyên những người này (chỉ cập nhật / thêm mới)</label><br>' +
              '<label class="chk" style="margin-top:4px"><input type="radio" name="iMode" value="replace"> Xóa họ – thay toàn bộ danh sách bằng file</label></div>' : '') + '</div>';
        $('mOk').disabled = !!parsed.err.length || !parsed.list.length;
      }).catch(function (e) { $('iRes').innerHTML = ''; T.showErr('Không đọc được file: ' + e.message); });
    };
    $('mOk').onclick = function () {
      var mode = (document.querySelector('input[name=iMode]:checked') || { value: 'merge' }).value, nl;
      if (mode === 'replace') nl = parsed.list;
      else {
        var map = {}; parsed.list.forEach(function (x) { map[x.id] = x; });
        nl = D.list.map(function (d) { var x = map[d.id]; if (x) delete map[d.id]; return x || d; });
        parsed.list.forEach(function (x) { if (map[x.id]) nl.push(x); });
      }
      var b = this; b.disabled = true;
      saveDrivers(nl).then(function (j) { T.closeModal(); renderDr(); T.toast('Đã cập nhật: thêm ' + j.added + ', sửa ' + j.updated + ', xóa ' + j.removed + '. Tổng ' + j.count + ' người.'); })
        .catch(function (e) { T.showErr(e.message); b.disabled = false; });
    };
  }
  function sheetFrom(X, aoa, widths) { var ws = X.utils.aoa_to_sheet(aoa, { cellDates: true }); ws['!cols'] = widths.map(function (w) { return { wch: w }; }); return ws; }
  function template() {
    need('xlsx').then(function (X) {
      var wb = X.utils.book_new();
      X.utils.book_append_sheet(wb, sheetFrom(X, [['Mã số', 'Họ và tên', 'Số đò', 'Tổ', 'Trạng thái'],
        ['TA-0001', 'Nguyễn Văn A', '001', 'Tổ 1', 'Hoạt động'], ['TA-0002', 'Trần Thị B', '002', 'Tổ 1', 'Hoạt động'], ['TA-0003', 'Lê Văn C', '015', 'Tổ 2', 'Nghỉ']], [12, 28, 10, 12, 14]), 'DanhSach');
      X.utils.book_append_sheet(wb, sheetFrom(X, [['HƯỚNG DẪN'], ['1. Mỗi người một dòng ở trang DanhSach, xóa các dòng ví dụ.'],
        ['2. Mã số: duy nhất, chỉ gồm chữ không dấu, số, dấu gạch ngang (VD: TA-0001). Mã số được in trên thẻ QR.'],
        ['3. Số đò nên định dạng Văn bản để giữ số 0 ở đầu (001).'], ['4. Trạng thái: "Hoạt động" hoặc "Nghỉ" (người nghỉ không hiện trên danh bạ, mã QR tạm ngừng).'],
        ['5. Trên trang quản trị: Người lái đò → Nhập từ Excel → chọn file này.']], [100]), 'HuongDan');
      X.writeFile(wb, 'mau_danh_sach_lai_do.xlsx');
    }).catch(function (e) { T.toast(e.message); });
  }
  function exportList() {
    need('xlsx').then(function (X) {
      var wb = X.utils.book_new();
      X.utils.book_append_sheet(wb, sheetFrom(X, [['Mã số', 'Họ và tên', 'Số đò', 'Tổ', 'Trạng thái', 'Link đánh giá']].concat(D.list.map(function (d) {
        return [d.id, d.name, d.boat, d.team, d.status, linkOf(d.id)]; })), [12, 28, 10, 12, 14, 50]), 'DanhSach');
      X.writeFile(wb, 'danh_sach_lai_do_' + ymd(new Date()) + '.xlsx');
    }).catch(function (e) { T.toast(e.message); });
  }
  function ymd(d) { return d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2); }

  /* =============================== IN THẺ QR & POSTER =============================== */
  var CARD = { lon: { w: 90, h: 125, ten: 'Thẻ gắn đò 9 × 12,5 cm – 4 thẻ/trang A4' }, deo: { w: 54, h: 86, ten: 'Thẻ đeo 5,4 × 8,6 cm – 9 thẻ/trang A4' } };
  var P = { type: 'lon', name: false, scope: 'active', team: '', ids: '', size: 'A4', hot: null };
  function panePr() {
    var p = $('pPr'); p.innerHTML = '<div class="loading"><span class="spin"></span>Đang tải…</div>';
    Promise.all([loadDrivers(true), P.hot === null ? api({ action: 'pub_dir' }).then(function (j) { return j.data.settings.hotline || ''; }).catch(function () { return ''; }) : Promise.resolve(P.hot)])
      .then(function (a) { P.hot = a[1]; renderPr(); }).catch(function (e) { p.innerHTML = '<div class="err">' + esc(e.message) + '</div>'; });
  }
  function renderPr() {
    var teams = []; D.list.forEach(function (x) { if (x.team && teams.indexOf(x.team) < 0) teams.push(x.team); }); teams.sort(TEAM_SORT);
    function rd(n, v, l, on) { return '<label class="chk" style="margin:4px 14px 4px 0"><input type="radio" name="' + n + '" value="' + v + '"' + (on ? ' checked' : '') + '> ' + l + '</label>'; }
    $('pPr').innerHTML =
      '<div class="card"><h3>Thẻ QR đánh giá cho người lái đò</h3>' +
        '<div class="row">' + rd('pType', 'lon', CARD.lon.ten, P.type === 'lon') + rd('pType', 'deo', CARD.deo.ten, P.type === 'deo') + '</div>' +
        '<label class="chk" style="margin:6px 0"><input type="checkbox" id="pName"' + (P.name ? ' checked' : '') + '> In họ tên trên thẻ</label>' +
        '<label class="f">In cho</label><div class="row">' + rd('pScope', 'active', 'Tất cả người đang hoạt động', P.scope === 'active') + rd('pScope', 'team', 'Một tổ', P.scope === 'team') + rd('pScope', 'ids', 'Một số mã cụ thể', P.scope === 'ids') + '</div>' +
        '<div id="pTeamW" class="' + (P.scope === 'team' ? '' : 'hidden') + '"><select class="i" id="pTeam" style="max-width:240px">' + teams.map(function (t) { return '<option' + (P.team === t ? ' selected' : '') + '>' + esc(t) + '</option>'; }).join('') + '</select></div>' +
        '<div id="pIdsW" class="' + (P.scope === 'ids' ? '' : 'hidden') + '"><input class="i" id="pIds" placeholder="VD: TA-0001, TA-0005, TA-0012" value="' + esc(P.ids) + '"></div>' +
        '<div class="row" style="margin-top:12px"><span class="muted" id="pCount"></span><span class="sp"></span><button class="btn pri" id="pGo">Xem trước &amp; in thẻ</button></div>' +
        '<p class="muted" style="font-size:12px;margin:10px 0 0">Trang in mở ở tab mới. Chọn máy in, hoặc chọn <b>Lưu dưới dạng PDF</b> để gửi đi in ấn. Khổ A4, tỷ lệ 100% (không chọn "Vừa trang"), bật "Đồ họa nền". Mã QR dạng vector, in sắc nét ở mọi kích thước.</p></div>' +
      '<div class="card"><h3>Poster mã QR tổng (danh bạ người lái đò)</h3><p class="muted" style="font-size:13px;margin:0 0 8px">Đặt tại quầy vé, bến xuất phát để du khách quét trước chuyến đi.</p>' +
        '<div class="row">' + rd('pSize', 'A4', 'Khổ A4', P.size === 'A4') + rd('pSize', 'A3', 'Khổ A3', P.size === 'A3') + '</div>' +
        '<label class="f">Đường dây nóng in trên poster (để trống nếu không in)</label><input class="i" id="pHot" value="' + esc(P.hot) + '" style="max-width:360px">' +
        '<div class="row" style="margin-top:12px"><span class="sp"></span><button class="btn" id="pDirQr">Tải riêng mã QR</button><button class="btn pri" id="pPoster">Xem trước &amp; in poster</button></div></div>';
    function upd() {
      P.type = document.querySelector('input[name=pType]:checked').value; P.name = $('pName').checked;
      P.scope = document.querySelector('input[name=pScope]:checked').value; P.team = $('pTeam').value; P.ids = $('pIds').value;
      P.size = document.querySelector('input[name=pSize]:checked').value; P.hot = $('pHot').value;
      $('pTeamW').classList.toggle('hidden', P.scope !== 'team'); $('pIdsW').classList.toggle('hidden', P.scope !== 'ids');
      var sel = pick(), per = P.type === 'lon' ? 4 : 9;
      $('pCount').textContent = sel.err || (sel.list.length + ' thẻ · ' + Math.ceil(sel.list.length / per) + ' trang A4');
    }
    $('pPr').oninput = upd; $('pPr').onchange = upd; upd();
    $('pGo').onclick = function () { var sel = pick(); if (sel.err || !sel.list.length) { T.toast(sel.err || 'Không có người nào để in.'); return; } printWin(function () { return cardsHtml(sel.list); }, 'Thẻ QR người lái đò'); };
    $('pPoster').onclick = function () { printWin(function () { return posterHtml(); }, 'Poster mã QR tổng'); };
    $('pDirQr').onclick = function () { showQr('Mã QR tổng (danh bạ)', 'Dùng để chèn vào ấn phẩm, biển bảng riêng.', dirLink(), 'qr_danh_ba'); };
  }
  function pick() {
    if (P.scope === 'team') return { list: D.list.filter(function (d) { return d.team === P.team && d.status !== 'Nghỉ'; }) };
    if (P.scope === 'ids') {
      var ids = P.ids.toUpperCase().split(/[\s,;]+/).filter(Boolean), list = [], miss = [];
      ids.forEach(function (id) { var d = D.list.filter(function (x) { return x.id === id; })[0]; if (d) { if (list.indexOf(d) < 0) list.push(d); } else miss.push(id); });
      return miss.length ? { list: list, err: 'Không có mã: ' + miss.join(', ') } : { list: list };
    }
    return { list: D.list.filter(function (d) { return d.status !== 'Nghỉ'; }) };
  }
  var STAR = '<svg viewBox="0 0 24 24"><path fill="#e3a51f" d="M12 1.8l3.1 6.6 7.2.9-5.3 5 1.4 7.1L12 17.9l-6.4 3.5 1.4-7.1-5.3-5 7.2-.9z"/></svg>';
  var PRINT_CSS = '@page{size:A4;margin:0}*{box-sizing:border-box;margin:0;padding:0}html,body{background:#888}' +
    'body{font-family:"Be Vietnam Pro",Arial,sans-serif;color:#1d2a24;-webkit-print-color-adjust:exact;print-color-adjust:exact}' +
    '.page{width:210mm;height:297mm;background:#fff;margin:8mm auto;position:relative;overflow:hidden;display:grid;align-content:center;justify-content:center;gap:5mm}' +
    '.note{position:absolute;left:10mm;bottom:4mm;font-size:7pt;color:#888}' +
    '.card{position:relative;border:.35mm solid #1f4d3a;border-radius:var(--r);overflow:hidden;display:flex;flex-direction:column;background:#fff}' +
    '.hd{background:#1f4d3a;color:#fff;text-align:center;display:flex;flex-direction:column;justify-content:center;align-items:center}' +
    '.ft{background:#f2efe6;text-align:center;display:flex;flex-direction:column;justify-content:center}' +
    '.mid{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-start}' +
    '.fit{white-space:nowrap;max-width:100%;overflow:hidden}.stars{display:flex;justify-content:center}.stars svg{display:block}' +
    '.bar{position:sticky;top:0;z-index:9;background:#1f4d3a;color:#fff;padding:10px 16px;display:flex;gap:12px;align-items:center;font:14px Arial}' +
    '.bar button{background:#fff;color:#1f4d3a;border:0;border-radius:8px;padding:8px 14px;font-weight:bold;cursor:pointer}' +
    '@media print{html,body{background:#fff}.page{margin:0;break-after:page}.bar{display:none}}';
  function printWin(build, title) {
    var w = window.open('', '_blank');
    if (!w) { T.toast('Trình duyệt đã chặn cửa sổ mới. Hãy cho phép cửa sổ bật lên (pop-up) cho trang này.'); return; }
    try { w.document.title = 'Đang tạo…'; w.document.body.innerHTML = '<p style="font:16px Arial;padding:30px">Đang tạo trang in…</p>'; } catch (e) {}
    need('qr').then(function () {
      var body = build();
      w.document.open();
      w.document.write('<!DOCTYPE html><html lang="vi"><head><meta charset="utf-8"><title>' + esc(title) + '</title>' +
        '<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;600;700&display=swap" rel="stylesheet">' +
        '<style>' + PRINT_CSS + '</style></head><body><div class="bar"><b>' + esc(title) + '</b><span style="flex:1"></span><span>Khổ giấy A4/A3 theo mẫu, tỷ lệ 100%, bật "Đồ họa nền"</span><button id="bPrint">In / Lưu PDF</button></div>' + body + '</body></html>');
      w.document.close();
      var doc = w.document;
      doc.getElementById('bPrint').onclick = function () { w.print(); };
      function fit() {      // co chữ dài (họ tên, tiêu đề) cho vừa khổ thẻ
        Array.prototype.forEach.call(doc.querySelectorAll('.fit'), function (e) {
          var sz = parseFloat(w.getComputedStyle(e).fontSize);
          while (e.scrollWidth > e.clientWidth + 0.5 && sz > 3) { sz -= 0.25; e.style.fontSize = sz + 'px'; }
        });
      }
      fit();
      Promise.race([doc.fonts ? doc.fonts.ready : Promise.resolve(), new Promise(function (r) { setTimeout(r, 2500); })])
        .then(function () { fit(); setTimeout(function () { w.focus(); w.print(); }, 300); });
    }).catch(function (e) { w.close(); T.toast(e.message); });
  }
  function cardsHtml(list) {
    var c = CARD[P.type], W = c.w, H = c.h, gap = 5, le = 10;
    var cols = Math.floor((210 - 2 * le + gap) / (W + gap)), rows = Math.floor((297 - 2 * le + gap) / (H + gap)), per = cols * rows;
    var pages = Math.ceil(list.length / per), html = '';
    var hh = H * 0.215, fh = H * 0.12, cao = H - hh - fh;
    var fn = W * 0.058, fi = W * 0.12, fm = W * 0.05, kh = W * 0.035;
    for (var pg = 0; pg < pages; pg++) {
      html += '<div class="page" style="grid-template-columns:repeat(' + cols + ',' + W + 'mm);grid-auto-rows:' + H + 'mm">';
      list.slice(pg * per, pg * per + per).forEach(function (d) {
        var hasName = P.name && d.name;
        var textH = kh + (hasName ? fn * 1.15 : 0) + fi * 0.95 + fm * 1.5 + W * 0.025;
        var qs = Math.min(W * 0.64, cao - cao * 0.05 - textH);
        var meta = [d.boat ? 'Đò số ' + d.boat : '', d.team].filter(Boolean).join(' · ');
        html += '<div class="card" style="width:' + W + 'mm;height:' + H + 'mm;--r:' + (W * 0.06) + 'mm">' +
          '<div class="hd" style="height:' + hh + 'mm;padding:0 ' + (W * 0.07) + 'mm">' +
            '<div class="fit" style="font-size:' + (W * 0.045) + 'mm;color:#d7e6df;letter-spacing:.03em">KHU DU LỊCH SINH THÁI TRÀNG AN</div>' +
            '<div class="fit" style="font-size:' + (W * 0.085) + 'mm;font-weight:700;margin:' + (H * 0.008) + 'mm 0">QUÉT MÃ ĐỂ ĐÁNH GIÁ</div>' +
            '<div class="stars" style="gap:' + (hh * 0.04) + 'mm">' + new Array(6).join('<span style="width:' + (hh * 0.17) + 'mm;height:' + (hh * 0.17) + 'mm;display:block">' + STAR + '</span>') + '</div></div>' +
          '<div class="mid" style="padding:' + (cao * 0.05) + 'mm ' + (W * 0.07) + 'mm 0">' +
            '<div style="width:' + qs + 'mm;height:' + qs + 'mm">' + qrSvg(linkOf(d.id)) + '</div>' +
            '<div style="height:' + kh + 'mm"></div>' +
            (hasName ? '<div class="fit" style="font-size:' + fn + 'mm;font-weight:700;line-height:1.15">' + esc(d.name) + '</div>' : '') +
            '<div class="fit" style="font-size:' + fi + 'mm;font-weight:700;color:#1f4d3a;line-height:1">' + esc(d.id) + '</div>' +
            (meta ? '<div class="fit" style="font-size:' + fm + 'mm;color:#5d6b64;margin-top:' + (fm * 0.4) + 'mm">' + esc(meta) + '</div>' : '') + '</div>' +
          '<div class="ft" style="height:' + fh + 'mm;padding:0 ' + (W * 0.07) + 'mm">' +
            '<div class="fit" style="font-size:' + (W * 0.042) + 'mm">Ý kiến của bạn giúp chúng tôi phục vụ tốt hơn</div>' +
            '<div class="fit" style="font-size:' + (W * 0.036) + 'mm;color:#5d6b64;margin-top:' + (fh * 0.05) + 'mm">Scan to rate your boat trip · Ẩn danh · Miễn phí</div></div></div>';
      });
      html += '<div class="note">Thẻ QR đánh giá người lái đò Tràng An – trang ' + (pg + 1) + '/' + pages + ' – cắt theo viền thẻ – in ngày ' + new Date().toLocaleDateString('vi-VN') + '</div></div>';
    }
    return html;
  }
  function posterHtml() {
    var A3 = P.size === 'A3', pw = A3 ? 297 : 210, ph = A3 ? 420 : 297, k = pw / 210;
    function t(top, size, txt, o) { o = o || {};
      return '<div class="fit" style="position:absolute;left:' + (o.l || 15 * k) + 'mm;right:' + (o.r || 15 * k) + 'mm;top:' + top + 'mm;text-align:center;font-size:' + size + 'pt;' + (o.b ? 'font-weight:700;' : '') + 'color:' + (o.c || '#1d2a24') + ';line-height:1.2">' + txt + '</div>'; }
    var qs = 105 * k, qt = 78 * k, html = '<div class="page" style="width:' + pw + 'mm;height:' + ph + 'mm;display:block">' +
      '<style>@page{size:' + P.size + '}</style>' +
      '<div style="position:absolute;left:0;right:0;top:0;height:' + (62 * k) + 'mm;background:#1f4d3a"></div>' +
      t(12 * k, 13 * k, 'KHU DU LỊCH SINH THÁI TRÀNG AN', { c: '#d7e6df' }) +
      t(22 * k, 34 * k, 'ĐỘI NGŨ NGƯỜI LÁI ĐÒ', { b: 1, c: '#fff' }) +
      t(38 * k, 17 * k, 'Thông tin &amp; đánh giá của du khách', { c: '#fff' }) +
      '<div class="stars" style="position:absolute;left:0;right:0;top:' + (49.5 * k) + 'mm;gap:' + (2.6 * k) + 'mm">' + new Array(6).join('<span style="width:' + (6.4 * k) + 'mm;height:' + (6.4 * k) + 'mm;display:block">' + STAR + '</span>') + '</div>' +
      '<div style="position:absolute;left:' + (pw / 2 - qs / 2 - 8 * k) + 'mm;top:' + (qt - 8 * k) + 'mm;width:' + (qs + 16 * k) + 'mm;height:' + (qs + 16 * k) + 'mm;border:' + (0.7 * k) + 'mm solid #1f4d3a;border-radius:' + (6 * k) + 'mm"></div>' +
      '<div style="position:absolute;left:' + (pw / 2 - qs / 2) + 'mm;top:' + qt + 'mm;width:' + qs + 'mm;height:' + qs + 'mm">' + qrSvg(dirLink()) + '</div>' +
      t(qt + qs + 14 * k, 22 * k, 'QUÉT MÃ TRƯỚC CHUYẾN ĐI', { b: 1, c: '#1f4d3a' }) +
      t(qt + qs + 24 * k, 14 * k, 'để xem thông tin và điểm đánh giá của từng người lái đò') +
      t(qt + qs + 31.5 * k, 12 * k, 'Scan to view our boat rowers and their visitor ratings', { c: '#5d6b64' });
    var by = qt + qs + 42 * k, bw = 82 * k, bh = 26 * k;
    [['1. Trước chuyến đi', 'Quét mã này để xem đội ngũ', 'người lái đò và điểm đánh giá'], ['2. Sau chuyến đi', 'Quét mã trên thẻ người lái đò', 'của bạn để đánh giá (30 giây)']].forEach(function (b, i) {
      var bx = i ? pw / 2 + 4 * k : pw / 2 - bw - 4 * k;
      html += '<div style="position:absolute;left:' + bx + 'mm;top:' + by + 'mm;width:' + bw + 'mm;height:' + bh + 'mm;background:#f2efe6;border-radius:' + (4 * k) + 'mm"></div>' +
        t(by + 3.5 * k, 13 * k, b[0], { b: 1, c: '#1f4d3a', l: bx + 3 * k, r: pw - bx - bw + 3 * k }) +
        t(by + 11 * k, 10.5 * k, b[1], { l: bx + 3 * k, r: pw - bx - bw + 3 * k }) +
        t(by + 16.5 * k, 10.5 * k, b[2], { l: bx + 3 * k, r: pw - bx - bw + 3 * k });
    });
    if (P.hot) html += t(ph - 20 * k, 13 * k, 'Phản ánh, góp ý: ' + esc(P.hot), { b: 1, c: '#b3402f' });
    return html + '</div>';
  }

  /* =============================== BÁO CÁO EXCEL =============================== */
  var R = null;
  function paneRp() {
    var now = new Date(), first = new Date(now.getFullYear(), now.getMonth(), 1);
    function iso(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
    $('pRp').innerHTML = '<div class="card" style="max-width:720px"><h3>Báo cáo kết quả đánh giá (Excel)</h3>' +
      '<p class="muted" style="font-size:13px;margin:0 0 8px">Gồm 5 trang: Tổng hợp, Xếp hạng từng người, Theo tổ, Phản ánh cần xác minh (có cột ghi kết quả xác minh), Dữ liệu gốc. Đã loại trừ các đánh giá bị ẩn/xóa.</p>' +
      '<div class="row" style="align-items:flex-end"><div><label class="f">Từ ngày</label><input class="i" type="date" id="rFrom" value="' + iso(first) + '"></div>' +
      '<div><label class="f">Đến ngày</label><input class="i" type="date" id="rTo" value="' + iso(now) + '"></div>' +
      '<div><label class="f">Số lượt tối thiểu để xếp loại</label><input class="i" type="number" id="rMin" value="20" min="1" max="1000" style="width:120px"></div></div>' +
      '<div class="row" id="rQuick" style="margin-top:8px"><button class="btn sm" data-q="m">Tháng này</button><button class="btn sm" data-q="pm">Tháng trước</button><button class="btn sm" data-q="q">Quý này</button><button class="btn sm" data-q="y">Năm nay</button>' +
      '<span class="sp"></span><button class="btn pri" id="rGo">Xuất báo cáo Excel</button></div><div class="err hidden" id="rErr"></div><div id="rOut"></div></div>';
    $('rQuick').onclick = function (e) {
      var b = e.target.closest('[data-q]'); if (!b) return;
      var n = new Date(), y = n.getFullYear(), m = n.getMonth(), f, t = n;
      if (b.dataset.q === 'm') f = new Date(y, m, 1);
      if (b.dataset.q === 'pm') { f = new Date(y, m - 1, 1); t = new Date(y, m, 0); }
      if (b.dataset.q === 'q') f = new Date(y, m - m % 3, 1);
      if (b.dataset.q === 'y') f = new Date(y, 0, 1);
      $('rFrom').value = iso(f); $('rTo').value = iso(t);
    };
    $('rGo').onclick = function () {
      var f = $('rFrom').value, t = $('rTo').value, mn = Math.max(1, parseInt($('rMin').value, 10) || 20), b = this;
      $('rErr').classList.add('hidden');
      if (!f || !t || f > t) { $('rErr').textContent = 'Chọn khoảng ngày hợp lệ ("Từ ngày" trước "Đến ngày").'; $('rErr').classList.remove('hidden'); return; }
      b.disabled = true; $('rOut').innerHTML = '<div class="loading"><span class="spin"></span>Đang tổng hợp dữ liệu…</div>';
      Promise.all([api({ action: 'adm_export', from: f, to: t }), need('xlsx')]).then(function (a) {
        var s = buildReport(a[1], a[0], f, t, mn);
        $('rOut').innerHTML = '<div class="kpis" style="margin-top:12px">' +
          kpi('Lượt đánh giá', s.tong.toLocaleString('vi-VN')) + kpi('Điểm TB toàn khu', s.tb == null ? '–' : T.fmt(s.tb)) +
          kpi('Phản ánh xin tiền', s.boa) + kpi('Đánh giá thấp', s.thap) + '</div><p class="muted" style="font-size:13px">Đã tải file <b>' + esc(s.file) + '</b>. Mở bằng Excel hoặc Google Sheets.</p>';
      }).catch(function (e) { $('rOut').innerHTML = ''; $('rErr').textContent = e.message; $('rErr').classList.remove('hidden'); })
        .then(function () { b.disabled = false; });
    };
  }
  function kpi(l, v) { return '<div class="kpi"><div class="l">' + l + '</div><div class="v">' + v + '</div></div>'; }
  function xepLoai(d, n, mn) { if (n < mn || d == null) return 'Chưa đủ lượt'; return d >= 4.5 ? 'Xuất sắc' : d >= 4 ? 'Tốt' : d >= 3.5 ? 'Đạt' : 'Cần cải thiện'; }
  function r2(x) { return Math.round(x * 100) / 100; }
  function vnDate(ms) { var d = new Date(ms + 7 * 3600000); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes()); }
  function buildReport(X, data, f, t, mn) {
    var TC = data.criteria, rows = data.rows, drivers = data.drivers, nc = TC.length;
    var map = {}; drivers.forEach(function (d) { map[d.id] = { d: d, n: 0, sum: TC.map(function () { return 0; }), tip: 0, low: 0, nx: 0 }; });
    rows.forEach(function (r) { var m = map[r[1]]; if (!m) return; m.n++; for (var i = 0; i < nc; i++) m.sum[i] += +r[2 + i] || 0;
      if (r[3 + nc] === 'Có') m.tip++; if ((+r[2 + nc] || 0) <= 2.5) m.low++; if (r[4 + nc]) m.nx++; });
    var kq = Object.keys(map).map(function (k) { var m = map[k], avgs = m.sum.map(function (s) { return m.n ? r2(s / m.n) : null; });
      var diem = m.n ? r2(avgs.reduce(function (a, b) { return a + b; }, 0) / nc) : null;
      return { d: m.d, n: m.n, avgs: avgs, diem: diem, tip: m.tip, low: m.low, nx: m.nx, xl: xepLoai(diem, m.n, mn) }; });
    kq.filter(function (x) { return x.n >= mn && x.diem != null; }).sort(function (a, b) { return b.diem - a.diem || b.n - a.n; }).forEach(function (x, i) { x.hang = i + 1; });
    var tong = rows.length, tb = tong ? r2(rows.reduce(function (s, r) { return s + (+r[2 + nc]); }, 0) / tong) : null;
    var boa = rows.filter(function (r) { return r[3 + nc] === 'Có'; }).length, thap = rows.filter(function (r) { return (+r[2 + nc] || 0) <= 2.5; }).length;
    var fd = function (s) { return s.split('-').reverse().join('/'); }, ky = fd(f) + ' – ' + fd(t);
    var wb = X.utils.book_new();
    var th = [['BÁO CÁO KẾT QUẢ ĐÁNH GIÁ NGƯỜI LÁI ĐÒ – KHU DU LỊCH SINH THÁI TRÀNG AN'], ['Kỳ báo cáo: ' + ky + '     |     Ngày xuất: ' + T.dt(Date.now()) + '     |     Người xuất: ' + T.S.user.name], [],
      ['CHỈ SỐ CHUNG', ''], ['Tổng số lượt đánh giá', tong], ['Điểm trung bình toàn khu (thang 5)', tb == null ? '–' : tb],
      ['Số người lái đò được đánh giá', kq.filter(function (x) { return x.n; }).length], ['Số người lái đò trong danh sách', drivers.length],
      ['Số phản ánh xin tiền boa / ép mua hàng', boa], ['Số lượt đánh giá thấp (≤ 2,5 sao)', thap], ['Số lượt có nhận xét, góp ý', rows.filter(function (r) { return r[4 + nc]; }).length], [], ['ĐIỂM THEO TIÊU CHÍ', '']];
    TC.forEach(function (c, i) { th.push([c, tong ? r2(rows.reduce(function (s, r) { return s + (+r[2 + i]); }, 0) / tong) : '–']); });
    th.push([], ['XẾP LOẠI (người có từ ' + mn + ' lượt trở lên)', '']);
    ['Xuất sắc', 'Tốt', 'Đạt', 'Cần cải thiện', 'Chưa đủ lượt'].forEach(function (l) { th.push([l, kq.filter(function (x) { return x.xl === l; }).length]); });
    th.push([], ['Ngưỡng xếp loại: Xuất sắc ≥ 4,50; Tốt 4,00–<4,50; Đạt 3,50–<4,00; Cần cải thiện < 3,50. Đã loại trừ đánh giá bị ẩn/xóa trên trang quản trị.']);
    X.utils.book_append_sheet(wb, sheetFrom(X, th, [50, 16]), 'TongHop');
    var xh = [['Hạng', 'Mã số', 'Họ và tên', 'Số đò', 'Tổ', 'Số lượt', 'Điểm TB'].concat(TC).concat(['Phản ánh xin boa', 'Lượt đánh giá thấp', 'Lượt có nhận xét', 'Xếp loại', 'Trạng thái'])];
    kq.slice().sort(function (a, b) { return (a.hang || 1e9) - (b.hang || 1e9) || b.n - a.n || (a.d.id < b.d.id ? -1 : 1); }).forEach(function (x) {
      xh.push([x.hang || '', x.d.id, x.d.name, x.d.boat, x.d.team, x.n, x.diem].concat(x.avgs).concat([x.tip, x.low, x.nx, x.xl, x.d.active === false ? 'Nghỉ' : ''])); });
    var ws = sheetFrom(X, xh, [6, 11, 26, 8, 10, 8, 9, 11, 11, 11, 11, 11, 11, 10, 10, 10, 15, 10]); ws['!autofilter'] = { ref: ws['!ref'] };
    X.utils.book_append_sheet(wb, ws, 'XepHang');
    var to = {};
    kq.forEach(function (x) { var k = x.d.team || '(Chưa ghi tổ)', a = to[k] = to[k] || { ng: 0, n: 0, s: 0, tip: 0, low: 0 }; a.ng++; a.n += x.n; a.s += (x.diem || 0) * x.n; a.tip += x.tip; a.low += x.low; });
    var tt = [['Tổ', 'Số người lái đò', 'Số lượt đánh giá', 'Điểm TB', 'Phản ánh xin boa', 'Lượt đánh giá thấp']];
    Object.keys(to).sort(TEAM_SORT).forEach(function (k) { var a = to[k]; tt.push([k, a.ng, a.n, a.n ? r2(a.s / a.n) : '', a.tip, a.low]); });
    X.utils.book_append_sheet(wb, sheetFrom(X, tt, [18, 14, 14, 10, 14, 14]), 'TheoTo');
    var dmap = {}; drivers.forEach(function (d) { dmap[d.id] = d; });
    var pa = [['Thời gian', 'Mã số', 'Họ và tên', 'Số đò', 'Tổ', 'Điểm', 'Loại', 'Nhận xét của du khách', 'Kết quả xác minh', 'Người xác minh']];
    rows.slice().sort(function (a, b) { return b[0] - a[0]; }).forEach(function (r) {
      var low = (+r[2 + nc] || 0) <= 2.5, tip = r[3 + nc] === 'Có', c = r[4 + nc];
      if (!(tip || low || c)) return;
      var d = dmap[r[1]] || {};
      pa.push([vnDate(r[0]), r[1], d.name || '', d.boat || '', d.team || '', +r[2 + nc], tip ? 'Xin boa/ép mua' : low ? 'Đánh giá thấp' : 'Nhận xét', c || '', '', '']); });
    ws = sheetFrom(X, pa, [17, 11, 24, 8, 10, 7, 15, 60, 30, 18]); fmtCol(X, ws, 0, 'dd/mm/yyyy hh:mm'); ws['!autofilter'] = { ref: ws['!ref'] };
    X.utils.book_append_sheet(wb, ws, 'PhanAnh');
    var g = [['Thời gian', 'Mã số'].concat(TC).concat(['Điểm TB', 'Xin boa', 'Nhận xét', 'Ngôn ngữ'])];
    rows.slice().sort(function (a, b) { return a[0] - b[0]; }).forEach(function (r) { g.push([vnDate(r[0]), r[1]].concat(r.slice(2, 2 + nc).map(Number)).concat([+r[2 + nc], r[3 + nc], r[4 + nc], r[5 + nc]])); });
    ws = sheetFrom(X, g, [17, 11].concat(TC.map(function () { return 10; })).concat([9, 9, 60, 9])); fmtCol(X, ws, 0, 'dd/mm/yyyy hh:mm'); ws['!autofilter'] = { ref: ws['!ref'] };
    X.utils.book_append_sheet(wb, ws, 'DuLieuGoc');
    var file = 'bao_cao_' + f.replace(/-/g, '') + '_den_' + t.replace(/-/g, '') + '.xlsx';
    X.writeFile(wb, file, { cellDates: true });
    return { tong: tong, tb: tb, boa: boa, thap: thap, file: file };
  }
  function fmtCol(X, ws, c, z) {
    var rg = X.utils.decode_range(ws['!ref']);
    for (var r = 1; r <= rg.e.r; r++) { var cell = ws[X.utils.encode_cell({ r: r, c: c })]; if (cell) cell.z = z; }
  }

  window.TA_TOOLS = { dr: paneDr, pr: panePr, rp: paneRp };
  window.TA_TOOLS._test = { parseSheet: parseSheet, qrSvg: qrSvg, need: need };
})();
