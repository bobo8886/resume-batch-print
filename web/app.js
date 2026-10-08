/* 简历一键打印 —— 前端逻辑 */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };

  var state = {
    info: null,
    printers: [],
    folder: '',
    files: [],
    selected: {},
    result: {},
    printing: false,
    stopFlag: false,
    filter: '',
    onlyUnprinted: false,
    pkPath: '',
    tested: false
  };

  /* ---------------- 小工具 ---------------- */

  function api(path, body) {
    return fetch(path, {
      method: body === undefined ? 'GET' : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: body === undefined ? undefined : JSON.stringify(body)
    }).then(function (r) {
      return r.json().catch(function () { return { ok: false, error: '服务返回异常 (HTTP ' + r.status + ')' }; });
    });
  }

  var toastTimer = null;
  function toast(msg, isErr) {
    var t = $('toast');
    t.textContent = msg;
    t.className = 'toast' + (isErr ? ' err' : '');
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.hidden = true; }, isErr ? 7000 : 2800);
  }

  function fmtSize(n) {
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(0) + ' KB';
    return (n / 1024 / 1024).toFixed(1) + ' MB';
  }
  function extLabel(e) { return e.replace('.', '').toUpperCase(); }
  function badgeClass(e) {
    if (e === '.pdf') return 'badge pdf';
    if (e === '.doc' || e === '.docx') return 'badge doc';
    return 'badge';
  }
  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  function asArray(v) { return Array.isArray(v) ? v : (v === null || v === undefined || v === '' ? [] : [v]); }
  function printerByName(name) {
    for (var i = 0; i < state.printers.length; i++) {
      if (state.printers[i].name === name) return state.printers[i];
    }
    return null;
  }
  function isWordExt(e) { return e === '.doc' || e === '.docx' || e === '.rtf' || e === '.odt' || e === '.wps' || e === '.txt'; }

  /* ---------------- 引擎状态 ---------------- */

  function renderEngines(busy) {
    var box = $('engines');
    var e = (state.info && state.info.engines) || {};
    var html = '';

    if (busy) {
      html += '<span class="pill busy"><span class="dot"></span>正在检测打印引擎…</span>';
      box.innerHTML = html;
      return;
    }

    html += e.sumatra
      ? '<span class="pill ok"><span class="dot"></span>PDF <b>已就绪</b></span>'
      : '<span class="pill bad"><span class="dot"></span>PDF <b>缺失</b></span>';

    if (e.word) html += '<span class="pill ok"><span class="dot"></span>Word <b>可用</b></span>';
    else if (e.wordPresent) html += '<span class="pill bad"><span class="dot"></span>Word <b>不可用</b></span>';
    else html += '<span class="pill"><span class="dot"></span>Word <b>未安装</b></span>';

    if (e.wps) html += '<span class="pill ok"><span class="dot"></span>WPS <b>可用</b></span>';
    else if (e.wpsPresent) html += '<span class="pill bad"><span class="dot"></span>WPS <b>不可用</b></span>';
    else html += '<span class="pill"><span class="dot"></span>WPS <b>未安装</b></span>';

    box.innerHTML = html;
  }

  function runSelfTest(silent) {
    if (!state.tested) { renderEngines(true); }
    return api('/api/selftest', {}).then(function (d) {
      state.tested = true;
      state.info = state.info || {};
      state.info.engines = d.engines;
      renderEngines(false);
      var e = d.engines || {};
      var okList = [];
      if (e.sumatra) okList.push('PDF');
      if (e.word) okList.push('Word');
      if (e.wps) okList.push('WPS');
      if (!silent) {
        toast(okList.length ? ('检测完成，可用引擎：' + okList.join(' / ')) : '没有检测到可用的打印引擎', okList.length === 0);
      }
      updateSummary();
    }).catch(function (err) {
      state.tested = true;
      renderEngines(false);
      if (!silent) toast('检测失败：' + err.message, true);
    });
  }

  /* ---------------- 初始化 ---------------- */

  function loadInfo(skipScan) {
    return api('/api/info').then(function (d) {
      state.info = d;
      state.printers = asArray(d.printers);
      $('folder').value = d.lastFolder || '';
      $('recursive').checked = d.recursive !== false;
      $('copies').value = d.copies || 1;
      var enh = $('enhance');
      if (enh) {
        enh.value = d.enhance || 'auto';
        if (d.enhanceAvailable === false) {
          enh.querySelector('option[value="dark"]').disabled = true;
          enh.querySelector('option[value="darker"]').disabled = true;
        }
        updateEnhanceInfo();
      }

      var sel = $('printer');
      sel.innerHTML = '';
      state.printers.forEach(function (p) {
        var o = document.createElement('option');
        o.value = p.name;
        o.textContent = p.name + (p.kind === 'virtual' ? '（虚拟打印机）' : (p.default ? '（系统默认）' : ''));
        sel.appendChild(o);
      });
      if (state.printers.length === 0) {
        var o2 = document.createElement('option');
        o2.value = '';
        o2.textContent = '（本机没有检测到打印机）';
        sel.appendChild(o2);
      }
      var want = d.printer || d.defaultPrinter || '';
      if (want && printerByName(want)) sel.value = want;
      else if (d.defaultPrinter && printerByName(d.defaultPrinter)) sel.value = d.defaultPrinter;

      var dl = $('folderList');
      if (dl) {
        dl.innerHTML = '';
        asArray(d.recentFolders).forEach(function (f) {
          var o = document.createElement('option');
          o.value = f;
          dl.appendChild(o);
        });
      }

      renderEngines(!state.tested);
      updatePrinterInfo();
      updateSummary();
      var banner = $('engineBanner');
      if (banner) banner.hidden = !d.engineMissing;
      if (d.lastFolder && !skipScan) scan(true);
    });
  }

  /* ---------------- 打印浓度 ---------------- */

  var ENH_TEXT = {
    auto: '自动判断：扫描件/图片版简历会自动加深，纯文字版简历原样打印。',
    normal: '原样打印，不做任何处理。',
    dark: '把浅灰色正文压成实心黑，适合底色偏浅、打印发虚的简历。',
    darker: '更强力的加深，适合整体很淡的扫描件或照片版简历。'
  };

  function updateEnhanceInfo() {
    var sel = $('enhance');
    if (!sel) return;
    var v = sel.value || 'auto';
    var box = $('enhanceInfo');
    box.textContent = ENH_TEXT[v] || '';
    var avail = !state.info || state.info.enhanceAvailable !== false;
    if (!avail && (v === 'dark' || v === 'darker')) {
      box.textContent = '本机不支持内置 PDF 渲染（需 Windows 10 及以上），将退回原样打印。';
    }
  }

  /* ---------------- 打印机信息 / 警告 ---------------- */

  function updatePrinterInfo() {
    var p = printerByName($('printer').value);
    var info = $('printerInfo');
    var warn = $('printerWarn');

    if (!p) {
      info.textContent = state.printers.length ? '' : '本机没有检测到打印机';
      warn.hidden = true;
      return;
    }

    var bits = [];
    if (p.kind === 'virtual') bits.push('虚拟打印机');
    else bits.push('可出纸');
    if (p.port) bits.push(p.port);
    if (p.network) bits.push('网络');
    if (p.default) bits.push('系统默认');
    info.textContent = bits.join(' · ');

    if (p.kind === 'virtual') {
      var real = state.printers.filter(function (x) { return x.kind === 'real'; });
      warn.innerHTML = '⚠️ <b>' + esc(p.name) + '</b> 是虚拟打印机，任务会被直接丢弃，一张纸都不会出。'
        + (real.length ? '请改选 <b>' + esc(real[0].name) + '</b>。' : '本机没有可出纸的打印机。');
      warn.hidden = false;
    } else {
      warn.hidden = true;
    }
  }

  /* ---------------- 扫描与渲染 ---------------- */

  function scan(silent) {
    var folder = $('folder').value.trim();
    if (!folder) { toast('请先选择或填写简历文件夹', true); return Promise.resolve(); }
    $('btnScan').disabled = true;
    $('scanHint').textContent = '正在扫描…';
    return api('/api/scan', { folder: folder, recursive: $('recursive').checked })
      .then(function (d) {
        if (d.ok === false || d.error) throw new Error(d.error || '扫描失败');
        state.folder = d.folder;
        $('folder').value = d.folder;
        state.files = asArray(d.files);
        state.selected = {};
        state.result = {};
        state.files.forEach(function (f) { state.selected[f.path] = true; });
        $('filesCard').hidden = false;
        $('placeholder').hidden = true;
        $('scanHint').textContent = '支持 PDF / Word / RTF / TXT';
        render();
        if (state.files.length === 0) toast('这个文件夹里没有找到简历文件');
        else if (!silent) toast('扫描完成，共 ' + state.files.length + ' 份，已全部勾选');
      })
      .catch(function (err) {
        toast(err.message, true);
        $('scanHint').textContent = '扫描失败，请检查路径';
      })
      .then(function () { $('btnScan').disabled = false; });
  }

  function visibleFiles() {
    var kw = state.filter.toLowerCase();
    return state.files.filter(function (f) {
      if (state.onlyUnprinted && f.printed) return false;
      if (kw && f.name.toLowerCase().indexOf(kw) < 0) return false;
      return true;
    });
  }

  function render() {
    var list = visibleFiles();
    var html = '';
    list.forEach(function (f, i) {
      var sel = !!state.selected[f.path];
      var cls = (sel ? 'sel ' : '') + (f.printed ? 'done' : '');
      var st = state.result[f.path];
      var stHtml = st === 'ok' ? '<span class="badge ok">已打印</span>'
        : st === 'dry' ? '<span class="badge warn">已试运行</span>'
          : st === 'err' ? '<span class="badge warn">失败</span>'
            : (f.printed ? '<span class="badge ok">打过</span>' : '');
      html += '<tr class="' + cls + '" data-path="' + esc(f.path) + '">'
        + '<td class="c-check"><input type="checkbox" ' + (sel ? 'checked' : '') + '></td>'
        + '<td class="c-idx idx">' + (i + 1) + '</td>'
        + '<td class="fname">' + esc(f.name) + '</td>'
        + '<td class="c-ext"><span class="' + badgeClass(f.ext) + '">' + extLabel(f.ext) + '</span></td>'
        + '<td class="c-size">' + fmtSize(f.size) + '</td>'
        + '<td class="c-time">' + esc(f.mtime) + '</td>'
        + '<td class="c-done">' + stHtml + '</td>'
        + '<td class="c-act"><button class="btn act-one">只打这个</button></td>'
        + '</tr>';
    });
    $('tbody').innerHTML = html;
    $('emptyHint').hidden = list.length > 0;

    var nSel = state.files.filter(function (f) { return state.selected[f.path]; }).length;
    $('count').textContent = '已选 ' + nSel + ' / ' + state.files.length;
    $('checkAll').checked = list.length > 0 && list.every(function (f) { return state.selected[f.path]; });
    updateSummary();
  }

  function updateSummary() {
    var picked = state.files.filter(function (f) { return state.selected[f.path]; });
    var n = picked.length;
    var pdf = picked.filter(function (f) { return f.ext === '.pdf'; }).length;
    var word = picked.filter(function (f) { return isWordExt(f.ext); }).length;

    $('sumCount').textContent = n + ' 份';
    var mix = [];
    if (pdf) mix.push('PDF ' + pdf);
    if (word) mix.push('Word ' + word);
    $('sumMix').textContent = mix.length ? mix.join('　') : '—';

    var copies = Math.max(1, parseInt($('copies').value, 10) || 1);
    $('btnPrint').textContent = copies > 1
      ? ('🖨️ 一键打印（' + n + ' 份 × ' + copies + '）')
      : ('🖨️ 一键打印（' + n + ' 份简历）');
    $('btnPrint').disabled = n === 0;
  }

  /* ---------------- 文件夹浏览器 ---------------- */

  function openPicker() {
    $('picker').hidden = false;
    var start = $('folder').value.trim();
    if (!start && state.info && state.info.lastFolder) start = state.info.lastFolder;
    pkLoad(start);
  }

  function pkLoad(path) {
    $('pkErr').textContent = '';
    $('pkList').innerHTML = '<div class="pk-empty">读取中…</div>';
    api('/api/list-dir', { path: path || '' }).then(function (d) {
      state.pkPath = d.path || '';
      $('pkPath').value = state.pkPath;
      renderCrumbs(state.pkPath);
      $('pkUp').disabled = !state.pkPath;

      var html = '';
      var shortcuts = asArray(d.shortcuts);
      if (shortcuts.length) {
        html += '<div class="pk-sec">快速定位</div><div class="pk-shortcuts">';
        shortcuts.forEach(function (s) {
          html += '<button class="btn pk-jump" data-jump="' + esc(s.path) + '">' + esc(s.name) + '</button>';
        });
        html += '</div>';
      }

      var drives = asArray(d.drives);
      if (drives.length) {
        html += '<div class="pk-sec">驱动器</div>';
        drives.forEach(function (x) {
          html += '<div class="pk-item pk-dir" data-path="' + esc(x.path) + '">'
            + '<span class="pk-ico">💽</span><span class="pk-name">' + esc(x.name) + '</span></div>';
        });
      }

      var dirs = asArray(d.dirs);
      if (dirs.length) {
        html += '<div class="pk-sec">' + (drives.length ? '子文件夹' : (state.pkPath ? '子文件夹' : '文件夹')) + '</div>';
        dirs.forEach(function (x) {
          html += '<div class="pk-item pk-dir" data-path="' + esc(x.path) + '">'
            + '<span class="pk-ico">📁</span><span class="pk-name">' + esc(x.name) + '</span></div>';
        });
      }

      if (!drives.length && !dirs.length) {
        html += '<div class="pk-empty">' + (shortcuts.length ? '这个文件夹里没有子文件夹，可以直接点「就用这个文件夹」'
          : '这个文件夹里没有子文件夹') + '</div>';
      }
      $('pkList').innerHTML = html;
      if (d.ok === false && d.error) { $('pkErr').textContent = d.error; }
    }).catch(function (e) {
      $('pkList').innerHTML = '';
      $('pkErr').textContent = '读取失败：' + e.message;
    });
  }

  function renderCrumbs(path) {
    var box = $('pkCrumbs');
    if (!path) { box.innerHTML = '<span class="crumb cur">此电脑</span>'; return; }
    var parts = path.split('\\').filter(function (x) { return x !== ''; });
    var html = '<span class="crumb" data-crumb="">此电脑</span>';
    var acc = '';
    parts.forEach(function (seg, i) {
      acc += (i === 0 ? seg : '\\' + seg);
      var isLast = (i === parts.length - 1);
      html += '<span class="crumb-sep">›</span><span class="crumb' + (isLast ? ' cur' : '')
        + '" data-crumb="' + esc(acc) + '">' + esc(seg) + '</span>';
    });
    box.innerHTML = html;
  }

  /* ---------------- 打印 ---------------- */

  function logLine(text, cls) {
    var box = $('log');
    var d = document.createElement('div');
    if (cls) d.className = cls;
    d.textContent = text;
    box.appendChild(d);
    box.scrollTop = box.scrollHeight;
  }

  function printList(paths) {
    if (state.printing) return;
    var printer = $('printer').value;
    var pInfo = printerByName(printer);
    var copies = Math.max(1, parseInt($('copies').value, 10) || 1);
    var dry = $('dryRun').checked;
    var enhance = $('enhance') ? $('enhance').value : 'auto';
    var total = paths.length;

    if (pInfo && pInfo.kind === 'virtual' && !dry) {
      var real = state.printers.filter(function (x) { return x.kind === 'real'; });
      var msg = '「' + printer + '」是虚拟打印机，任务会被直接丢弃，一张纸都不会出。';
      if (real.length) {
        if (!confirm(msg + '\n\n是否改用它打印：' + real[0].name + ' ？')) { return; }
        $('printer').value = real[0].name;
        updatePrinterInfo();
        printer = real[0].name;
      } else if (!confirm(msg + '\n\n仍要继续吗？')) {
        return;
      }
    }

    state.printing = true;
    state.stopFlag = false;
    state.result = {};
    $('overlay').hidden = false;
    $('log').innerHTML = '';
    $('btnStop').hidden = false;
    $('btnClose').hidden = true;
    $('ovTitle').textContent = dry ? '试运行（不会真的打印）' : '正在打印…';
    $('bar').style.width = '0%';
    logLine('目标打印机：' + (printer || '系统默认打印机'), 'l-info');
    logLine('每份份数：' + copies + '　　共 ' + total + ' 份简历', 'l-info');
    logLine('────────────────────────────', 'l-info');

    var i = 0, okCount = 0, failCount = 0, t0 = Date.now();

    function next() {
      if (state.stopFlag) { finish(true); return; }
      if (i >= total) { finish(false); return; }
      var path = paths[i];
      var name = path.split('\\').pop();
      $('ovStat').textContent = '正在处理 ' + (i + 1) + ' / ' + total + '：' + name;
      $('bar').style.width = Math.round((i / total) * 100) + '%';

      api('/api/print-one', { path: path, printer: printer, copies: copies, dryRun: dry, enhance: enhance })
        .then(function (r) {
          if (r.ok) {
            okCount++;
            state.result[path] = dry ? 'dry' : 'ok';
            logLine('✔ ' + name + '　→ ' + (r.method || '') + ' ' + (r.detail || ''), 'l-ok');
          } else {
            failCount++;
            state.result[path] = 'err';
            logLine('✘ ' + name + '　→ ' + (r.detail || r.error || '失败'), 'l-err');
          }
        })
        .catch(function (err) {
          failCount++;
          state.result[path] = 'err';
          logLine('✘ ' + name + '　→ ' + err.message, 'l-err');
        })
        .then(function () { i++; render(); next(); });
    }

    function finish(stopped) {
      state.printing = false;
      $('bar').style.width = '100%';
      var sec = ((Date.now() - t0) / 1000).toFixed(1);
      var virtualNote = (pInfo && pInfo.kind === 'virtual' && !dry)
        ? '　注意：目标是虚拟打印机，不会出纸。' : '';
      $('ovStat').textContent = (stopped ? '已停止。' : '完成。')
        + (dry ? '试运行 ' : '成功 ') + okCount + ' 份，失败 ' + failCount + ' 份，用时 ' + sec + ' 秒。' + virtualNote;
      $('ovTitle').textContent = failCount === 0
        ? (dry ? '试运行结束' : '全部打印完成 ✔')
        : '打印结束（有 ' + failCount + ' 份失败）';
      $('btnStop').hidden = true;
      $('btnClose').hidden = false;
      render();
      api('/api/end-batch', {}).catch(function () { }).then(function () { loadInfo(true); });
    }

    next();
  }

  function selectedPaths() {
    return state.files.filter(function (f) { return state.selected[f.path]; }).map(function (f) { return f.path; });
  }

  /* ---------------- 事件绑定 ---------------- */

  $('btnBrowse').onclick = openPicker;
  $('btnBrowse2').onclick = openPicker;
  $('btnScan').onclick = function () { scan(); };
  $('folder').onkeydown = function (e) { if (e.key === 'Enter') scan(); };
  $('recursive').onchange = function () { if (state.folder) scan(true); };
  $('btnRetest').onclick = function () { runSelfTest(false); };

  var btnFetch = $('btnFetchEngine');
  if (btnFetch) {
    btnFetch.onclick = function () {
      var self = this;
      if (!confirm('将从 SumatraPDF 官方站点下载 PDF 打印引擎（约 8 MB），\n下载后会校验 SHA-256 确保文件未被篡改。\n\n现在开始下载吗？')) return;
      self.disabled = true;
      self.textContent = '正在下载…';
      api('/api/fetch-engine', {}).then(function (d) {
        self.disabled = false;
        self.textContent = '下载 PDF 引擎';
        if (d.ok) {
          $('engineBanner').hidden = true;
          toast('PDF 引擎安装完成，可以打印 PDF 简历了');
          state.tested = false;
          loadInfo(true).then(function () { return runSelfTest(true); });
        } else {
          toast('下载失败：' + (d.detail || ''), true);
        }
      }).catch(function (e) {
        self.disabled = false;
        self.textContent = '下载 PDF 引擎';
        toast('下载失败：' + e.message, true);
      });
    };
  }

  // ---- 份数步进器 ----
  function bumpCopies(delta) {
    var v = Math.max(1, Math.min(99, (parseInt($('copies').value, 10) || 1) + delta));
    $('copies').value = v;
    updateSummary();
  }
  $('cpMinus').onclick = function () { bumpCopies(-1); };
  $('cpPlus').onclick = function () { bumpCopies(1); };
  $('copies').oninput = updateSummary;
  $('copies').onchange = function () {
    this.value = Math.max(1, Math.min(99, parseInt(this.value, 10) || 1));
    updateSummary();
  };

  // ---- 文件夹浏览器 ----
  $('pkCancel').onclick = function () { $('picker').hidden = true; };
  $('pkOk').onclick = function () {
    var chosen = $('pkPath').value.trim() || state.pkPath;
    if (!chosen) { toast('请先进入一个文件夹', true); return; }
    $('picker').hidden = true;
    $('folder').value = chosen;
    scan();
  };
  $('pkUp').onclick = function () {
    api('/api/list-dir', { path: state.pkPath }).then(function (d) { pkLoad(d.parent !== undefined ? d.parent : ''); });
  };
  $('pkGo').onclick = function () { pkLoad($('pkPath').value.trim()); };
  $('pkPath').onkeydown = function (e) { if (e.key === 'Enter') { e.preventDefault(); pkLoad(this.value.trim()); } };
  $('pkList').onclick = function (e) {
    var item = e.target.closest ? e.target.closest('.pk-dir') : null;
    if (item) { pkLoad(item.getAttribute('data-path')); return; }
    var jump = e.target.closest ? e.target.closest('.pk-jump') : null;
    if (jump) { pkLoad(jump.getAttribute('data-jump')); }
  };
  $('pkCrumbs').onclick = function (e) {
    var c = e.target.closest ? e.target.closest('.crumb') : null;
    if (c) pkLoad(c.getAttribute('data-crumb'));
  };
  $('picker').onclick = function (e) { if (e.target === this) this.hidden = true; };

  // ---- 主界面 ----
  $('printer').onchange = updatePrinterInfo;
  if ($('enhance')) { $('enhance').onchange = updateEnhanceInfo; }
  $('search').oninput = function () { state.filter = this.value.trim(); render(); };
  $('onlyUnprinted').onchange = function () { state.onlyUnprinted = this.checked; render(); };

  Array.prototype.forEach.call(document.querySelectorAll('[data-sel]'), function (b) {
    b.onclick = function () {
      var mode = this.getAttribute('data-sel');
      var list = visibleFiles();
      if (mode === 'all') list.forEach(function (f) { state.selected[f.path] = true; });
      if (mode === 'none') state.files.forEach(function (f) { state.selected[f.path] = false; });
      if (mode === 'invert') list.forEach(function (f) { state.selected[f.path] = !state.selected[f.path]; });
      if (mode === 'pdf') list.forEach(function (f) { state.selected[f.path] = (f.ext === '.pdf'); });
      if (mode === 'word') list.forEach(function (f) { state.selected[f.path] = isWordExt(f.ext); });
      render();
    };
  });

  $('checkAll').onchange = function () {
    var v = this.checked;
    visibleFiles().forEach(function (f) { state.selected[f.path] = v; });
    render();
  };

  $('tbody').onclick = function (e) {
    var tr = e.target.closest ? e.target.closest('tr') : null;
    if (!tr) return;
    var path = tr.getAttribute('data-path');
    if (e.target.classList.contains('act-one')) { printList([path]); return; }
    state.selected[path] = !state.selected[path];
    render();
  };

  $('btnPrint').onclick = function () {
    var paths = selectedPaths();
    if (paths.length === 0) { toast('还没有勾选简历', true); return; }
    printList(paths);
  };

  $('btnStop').onclick = function () {
    state.stopFlag = true;
    this.disabled = true;
    var self = this;
    setTimeout(function () { self.disabled = false; }, 1500);
  };
  $('btnClose').onclick = function () { $('overlay').hidden = true; };
  $('overlay').onclick = function (e) { if (e.target === this && !state.printing) this.hidden = true; };
  $('btnOpenRec').onclick = function () { api('/api/open', {}); };
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape') return;
    if (!$('picker').hidden) { $('picker').hidden = true; return; }
    if (!$('overlay').hidden && !state.printing) { $('overlay').hidden = true; }
  });

  $('btnTest').onclick = function () {
    var btn = this;
    var printer = $('printer').value;
    var pInfo = printerByName(printer);
    if (pInfo && pInfo.kind === 'virtual') {
      var real = state.printers.filter(function (x) { return x.kind === 'real'; });
      var msg = '「' + printer + '」是虚拟打印机，测试页也出不来。';
      if (real.length) {
        if (!confirm(msg + '\n\n是否改用 ' + real[0].name + ' 打测试页？')) return;
        $('printer').value = real[0].name;
        updatePrinterInfo();
        printer = real[0].name;
      }
    }
    btn.disabled = true;
    toast('正在发送测试页…');
    api('/api/testprint', { printer: printer }).then(function (d) {
      btn.disabled = false;
      toast(d.ok ? ('测试页已发送：' + (d.detail || '')) : ('测试页失败：' + (d.detail || '')), !d.ok);
    }).catch(function (e) { btn.disabled = false; toast(e.message, true); });
  };

  /* ---------------- 启动 ---------------- */
  loadInfo()
    .then(function () { return runSelfTest(true); })
    .catch(function (e) { toast('无法连接本地服务：' + e.message, true); });
})();
