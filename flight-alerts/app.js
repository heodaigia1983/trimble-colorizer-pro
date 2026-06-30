/* SGN-PQC Price Alert Tracker · Sun PhuQuoc Airways · Khứ hồi */

var STORE_PRICES   = 'sgn_pqc_prices';
var STORE_SETTINGS = 'sgn_pqc_settings';
var STORE_LOG      = 'sgn_pqc_alert_log';

var prices   = [];
var settings = { threshold: 0, days: [5, 6, 0, 1] };
var alertLog = [];
var sortAsc  = true;

var DAY_SHORT = ['CN','T2','T3','T4','T5','T6','T7'];
var DAY_FULL  = ['Chủ nhật','Thứ 2','Thứ 3','Thứ 4','Thứ 5','Thứ 6','Thứ 7'];

/* ── Persistence ── */
function loadAll() {
  try { prices = JSON.parse(localStorage.getItem(STORE_PRICES) || '[]'); } catch(e) { prices = []; }
  try { var s = JSON.parse(localStorage.getItem(STORE_SETTINGS)); if (s) settings = Object.assign(settings, s); } catch(e) {}
  try { alertLog = JSON.parse(localStorage.getItem(STORE_LOG) || '[]'); } catch(e) { alertLog = []; }
}
function saveAll() {
  try {
    localStorage.setItem(STORE_PRICES,   JSON.stringify(prices));
    localStorage.setItem(STORE_SETTINGS, JSON.stringify(settings));
    localStorage.setItem(STORE_LOG,      JSON.stringify(alertLog));
  } catch(e) {}
}

/* ── Helpers ── */
function fmtNum(n) {
  if (n == null || isNaN(n)) return '—';
  return Number(n).toLocaleString('vi-VN');
}
function fmtDate(dateStr) {
  if (!dateStr) return '—';
  var d = new Date(dateStr + 'T00:00:00');
  return DAY_SHORT[d.getDay()] + ' ' + pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1);
}
function fmtTs(ts) {
  var d = new Date(ts);
  return pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ' ' + pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1);
}
function pad2(n) { return String(n).padStart(2, '0'); }
function dayOf(dateStr) {
  if (!dateStr) return null;
  return new Date(dateStr + 'T00:00:00').getDay();
}
function esc(s) {
  return String(s).replace(/[&<>"]/g, function(c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}
function totalOf(p) { return (p.priceOut || 0) + (p.priceRet || 0); }
function isCheap(total) { return settings.threshold > 0 && total <= settings.threshold; }

/* ── Status label ── */
function statusLabel(total) {
  if (!settings.threshold) return '<span style="color:#9aa0a6">—</span>';
  var r = total / settings.threshold;
  if (r <= 1)    return '<span style="color:#1e8e3e;font-weight:700">🟢 Rẻ!</span>';
  if (r <= 1.15) return '<span style="color:#e37400">🟡 Khá rẻ</span>';
  return '<span style="color:#c5221f">🔴 Đắt</span>';
}

/* ── Render stats ── */
function renderStats() {
  var minEl   = document.getElementById('statMin');
  var avgEl   = document.getElementById('statAvg');
  var cntEl   = document.getElementById('statCount');
  if (!prices.length) {
    if (minEl) minEl.textContent = '—';
    if (avgEl) avgEl.textContent = '—';
    if (cntEl) cntEl.textContent = '0';
    return;
  }
  var totals = prices.map(totalOf);
  var min = Math.min.apply(null, totals);
  var avg = totals.reduce(function(a, b) { return a + b; }, 0) / totals.length;
  if (minEl) minEl.textContent = fmtNum(min);
  if (avgEl) avgEl.textContent = fmtNum(Math.round(avg));
  if (cntEl) cntEl.textContent = prices.length;
}

/* ── Render history table ── */
function renderHistory() {
  var empty = document.getElementById('emptyMsg');
  var table = document.getElementById('historyTable');
  var body  = document.getElementById('historyBody');
  if (!empty || !table || !body) return;

  if (!prices.length) {
    empty.style.display = '';
    table.style.display = 'none';
    return;
  }
  empty.style.display = 'none';
  table.style.display = '';

  var sorted = prices.slice().sort(function(a, b) {
    var ta = totalOf(a), tb = totalOf(b);
    return sortAsc ? ta - tb : tb - ta;
  });

  body.innerHTML = sorted.map(function(p) {
    var total = totalOf(p);
    var cheap = isCheap(total);
    return '<tr class="' + (cheap ? 'cheap' : '') + '">'
      + '<td>' + fmtDate(p.depDate) + '</td>'
      + '<td>' + fmtDate(p.retDate) + '</td>'
      + '<td class="r">' + fmtNum(p.priceOut) + '</td>'
      + '<td class="r">' + fmtNum(p.priceRet) + '</td>'
      + '<td class="r total-col">' + fmtNum(total) + '</td>'
      + '<td>' + statusLabel(total) + '</td>'
      + '<td style="max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">' + esc(p.note || '') + '</td>'
      + '<td style="white-space:nowrap;color:#9aa0a6">' + fmtTs(p.ts) + '</td>'
      + '<td><button onclick="deletePrice(' + p.id + ')" '
        + 'style="background:none;border:none;cursor:pointer;color:#d93025;font-size:12px;padding:0" '
        + 'title="Xoá">✕</button></td>'
      + '</tr>';
  }).join('');
}

/* ── Render alert log ── */
function renderLog() {
  var el = document.getElementById('alertLog');
  if (!el) return;
  if (!alertLog.length) {
    el.innerHTML = '<span style="color:#9aa0a6">Chưa có cảnh báo nào...</span>';
    return;
  }
  el.innerHTML = alertLog.slice().reverse().map(function(e) {
    return '<span style="color:#1e8e3e">✓</span> <span style="color:#9aa0a6">' + fmtTs(e.ts) + '</span> ' + esc(e.msg);
  }).join('\n');
  el.scrollTop = el.scrollHeight;
}

/* ── Render settings ── */
function renderSettings() {
  var inp = document.getElementById('thresholdInput');
  if (inp) inp.value = settings.threshold || '';
  document.querySelectorAll('.day-chip').forEach(function(chip) {
    chip.classList.toggle('active', settings.days.indexOf(parseInt(chip.dataset.day)) !== -1);
  });
}

/* ── Toast notification ── */
var _toastTimer = null;
function showToast(title, msg) {
  var t = document.getElementById('toast');
  var tTitle = document.getElementById('toastTitle');
  var tMsg   = document.getElementById('toastMsg');
  if (!t || !tTitle || !tMsg) return;
  tTitle.textContent = title;
  tMsg.textContent   = msg;
  t.classList.remove('hidden');
  if (_toastTimer) clearTimeout(_toastTimer);
  _toastTimer = setTimeout(closeToast, 7000);
}
function closeToast() {
  var t = document.getElementById('toast');
  if (t) t.classList.add('hidden');
}

/* ── Browser push notification ── */
function pushNotif(title, body) {
  if (typeof Notification !== 'undefined' && Notification.permission === 'granted') {
    try { new Notification(title, { body: body }); } catch(e) {}
  }
}

/* ── Check alert threshold when a price is added ── */
function checkAlert(p) {
  var total = totalOf(p);
  if (!isCheap(total)) return;

  var dep = fmtDate(p.depDate);
  var ret = fmtDate(p.retDate);
  var msg = dep + ' → ' + ret + ': ' + fmtNum(total) + 'k (ngưỡng ≤ ' + fmtNum(settings.threshold) + 'k)';

  showToast('🎉 Giá rẻ SGN ↔ PQC!', msg);
  pushNotif('✈ Giá rẻ SGN-PQC khứ hồi!', msg);

  alertLog.push({ ts: Date.now(), msg: msg });
  if (alertLog.length > 100) alertLog = alertLog.slice(-100);
  renderLog();
}

/* ── Update total preview ── */
function updatePreview() {
  var out   = parseFloat(document.getElementById('priceOut').value) || 0;
  var ret   = parseFloat(document.getElementById('priceRet').value) || 0;
  var total = out + ret;
  var el    = document.getElementById('totalPreview');
  if (!el) return;
  if (!out && !ret) { el.textContent = '—'; el.style.color = '#1a73e8'; return; }
  el.textContent  = fmtNum(total) + ' nghìn đ';
  el.style.color  = isCheap(total) ? '#1e8e3e' : '#1a73e8';
  el.style.fontWeight = isCheap(total) ? '800' : '700';
}

/* ── Day warning ── */
function updateDayWarning() {
  var depDate = document.getElementById('depDate').value;
  var warn = document.getElementById('dayWarning');
  if (!warn) return;
  var d = dayOf(depDate);
  var inList = d !== null && settings.days.indexOf(d) !== -1;
  warn.style.display = (depDate && !inList) ? '' : 'none';
}

/* ── Add price ── */
function addPrice() {
  var depDate  = document.getElementById('depDate').value;
  var retDate  = document.getElementById('retDate').value;
  var priceOut = parseFloat(document.getElementById('priceOut').value);
  var priceRet = parseFloat(document.getElementById('priceRet').value);
  var note     = (document.getElementById('priceNote').value || '').trim();

  if (!depDate)               { alert('Vui lòng chọn ngày đi.'); return; }
  if (!retDate)               { alert('Vui lòng chọn ngày về.'); return; }
  if (isNaN(priceOut) || priceOut <= 0) { alert('Vui lòng nhập giá chiều đi hợp lệ.'); return; }
  if (isNaN(priceRet) || priceRet <= 0) { alert('Vui lòng nhập giá chiều về hợp lệ.'); return; }
  if (retDate < depDate)      { alert('Ngày về phải sau ngày đi.'); return; }

  var p = {
    id:       Date.now(),
    depDate:  depDate,
    retDate:  retDate,
    priceOut: priceOut,
    priceRet: priceRet,
    note:     note,
    ts:       Date.now()
  };

  prices.unshift(p);
  saveAll();
  renderStats();
  renderHistory();
  checkAlert(p);

  document.getElementById('priceOut').value  = '';
  document.getElementById('priceRet').value  = '';
  document.getElementById('priceNote').value = '';
  updatePreview();
}

/* ── Delete price ── */
function deletePrice(id) {
  prices = prices.filter(function(p) { return p.id !== id; });
  saveAll();
  renderStats();
  renderHistory();
}

/* ── Save settings ── */
function saveSettings() {
  var val = parseFloat(document.getElementById('thresholdInput').value);
  settings.threshold = isNaN(val) ? 0 : val;

  var days = [];
  document.querySelectorAll('.day-chip.active').forEach(function(chip) {
    days.push(parseInt(chip.dataset.day));
  });
  settings.days = days;

  saveAll();
  renderHistory();
  updatePreview();
  updateDayWarning();

  var btn = document.getElementById('saveSettingsBtn');
  if (btn) {
    btn.textContent = '✓ Đã lưu!';
    btn.style.background = '#0f9d58';
    setTimeout(function() { btn.textContent = '💾 Lưu cài đặt'; btn.style.background = '#1a73e8'; }, 1800);
  }
}

/* ── Notification permission ── */
function requestNotif() {
  if (typeof Notification === 'undefined') {
    alert('Trình duyệt không hỗ trợ thông báo đẩy.'); return;
  }
  Notification.requestPermission().then(function(perm) {
    updateNotifBtn();
    if (perm === 'granted') {
      alertLog.push({ ts: Date.now(), msg: 'Đã bật thông báo trình duyệt.' });
      saveAll(); renderLog();
    }
  });
}
function updateNotifBtn() {
  var btn = document.getElementById('notifPermBtn');
  if (!btn) return;
  if (typeof Notification === 'undefined') {
    btn.textContent = '🔕 Trình duyệt không hỗ trợ'; btn.disabled = true; return;
  }
  if (Notification.permission === 'granted') {
    btn.textContent = '✅ Thông báo đã bật'; btn.style.color = '#1e8e3e';
  } else if (Notification.permission === 'denied') {
    btn.textContent = '🚫 Thông báo bị chặn'; btn.style.color = '#c5221f';
  }
}

/* ── Default dates (today + 3 days) ── */
function setDefaultDates() {
  var now  = new Date();
  var plus3 = new Date(now.getTime() + 3 * 86400000);
  var fmt = function(d) {
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  };
  var depEl = document.getElementById('depDate');
  var retEl = document.getElementById('retDate');
  if (depEl && !depEl.value) depEl.value = fmt(now);
  if (retEl && !retEl.value) retEl.value = fmt(plus3);
}

/* ── Init ── */
window.addEventListener('load', function() {
  loadAll();
  renderSettings();
  renderStats();
  renderHistory();
  renderLog();
  updateNotifBtn();
  setDefaultDates();
  updatePreview();

  /* Day chip toggle */
  document.querySelectorAll('.day-chip').forEach(function(chip) {
    chip.addEventListener('click', function() { this.classList.toggle('active'); });
  });

  /* Price input preview */
  document.getElementById('priceOut').addEventListener('input', updatePreview);
  document.getElementById('priceRet').addEventListener('input', updatePreview);
  document.getElementById('depDate').addEventListener('change', updateDayWarning);

  /* Buttons */
  document.getElementById('saveSettingsBtn').addEventListener('click', saveSettings);
  document.getElementById('notifPermBtn').addEventListener('click', requestNotif);
  document.getElementById('addPriceBtn').addEventListener('click', addPrice);

  document.getElementById('sortBtn').addEventListener('click', function() {
    sortAsc = !sortAsc;
    this.textContent = sortAsc ? '↑ Rẻ nhất trước' : '↓ Đắt nhất trước';
    renderHistory();
  });

  document.getElementById('clearAllBtn').addEventListener('click', function() {
    if (!prices.length) return;
    if (confirm('Xoá toàn bộ ' + prices.length + ' mức giá đã lưu?')) {
      prices = [];
      saveAll();
      renderStats();
      renderHistory();
    }
  });

  document.getElementById('clearLogBtn').addEventListener('click', function() {
    alertLog = [];
    saveAll();
    renderLog();
  });
});
