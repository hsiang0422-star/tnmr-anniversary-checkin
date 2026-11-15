(function () {
  'use strict';
  var CONF = window.CHECKIN_CONFIG || {};
  var API_URL = (CONF.API_URL || '').trim();
  var POLL_MS = CONF.POLL_MS || 4000;
  var MODE = document.body.dataset.mode; // 'checkin' | 'view'
  var PASS_KEY = 'tnmr-checkin-pass';
  var DEMO = !API_URL;

  var $ = function (id) { return document.getElementById(id); };
  var esc = function (s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  };
  function store(k, v) {
    try { if (v === undefined) return localStorage.getItem(k); if (v === null) localStorage.removeItem(k); else localStorage.setItem(k, v); }
    catch (e) { return null; }
  }

  /* ---------------- 資料存取 ---------------- */

  // 正式：Apps Script。POST 用 text/plain，避免跨網域預檢。
  var realApi = {
    list: function () {
      return fetch(API_URL + '?action=list&t=' + Date.now(), { cache: 'no-store' })
        .then(function (r) { return r.json(); }).then(unwrap);
    },
    post: function (body) {
      return fetch(API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(body)
      }).then(function (r) { return r.json(); }).then(unwrap);
    }
  };
  function unwrap(res) {
    if (!res || !res.ok) throw new Error((res && res.error) || '連線失敗，請再試一次。');
    return res;
  }

  // 示範：存在瀏覽器，規則與正式版相同。示範密碼為 0000。
  var DEMO_KEY = 'tnmr-checkin-demo';
  function taipeiHHMM(iso) {
    var d = new Date(iso);
    if (isNaN(d)) return '--:--';
    var t = new Date(d.getTime() + 8 * 3600 * 1000);
    var p = function (n) { return String(n).padStart(2, '0'); };
    return p(t.getUTCHours()) + ':' + p(t.getUTCMinutes());
  }
  function demoRead() { try { return JSON.parse(store(DEMO_KEY) || '[]'); } catch (e) { return []; } }
  function demoOut(arr) { return { ok: true, guests: arr.map(function (g) { return { id: g.id, name: g.name, title: g.title, time: taipeiHHMM(g.signedAt) }; }) }; }
  var demoApi = {
    list: function () { return Promise.resolve(demoOut(demoRead())); },
    post: function (b) {
      return new Promise(function (resolve) { setTimeout(resolve, 150); }).then(function () {
        if (b.pass !== '0000') throw new Error('管理密碼錯誤。（示範模式密碼：0000）');
        var arr = demoRead();
        var name = String(b.name || '').trim(), title = String(b.title || '').trim();
        var i = arr.findIndex(function (x) { return x.id === b.id; });
        if (b.action === 'check') return { ok: true };
        if ((b.action === 'add' || b.action === 'edit') && !name) throw new Error('請輸入來賓姓名。');
        if (b.action !== 'add' && i < 0) throw new Error('找不到這位來賓，可能已被刪除。');
        if (b.action === 'add') arr.push({ id: Math.random().toString(36).slice(2, 10), name: name, title: title, signedAt: new Date().toISOString() });
        if (b.action === 'edit') { arr[i].name = name; arr[i].title = title; }
        if (b.action === 'delete') arr.splice(i, 1);
        if (b.action === 'move') {
          var j = i + (b.dir < 0 ? -1 : 1);
          if (j >= 0 && j < arr.length) { var t = arr[i]; arr[i] = arr[j]; arr[j] = t; }
        }
        store(DEMO_KEY, JSON.stringify(arr));
        return demoOut(arr);
      });
    }
  };
  var api = DEMO ? demoApi : realApi;

  /* ---------------- 畫面 ---------------- */

  var guests = [];
  var busy = false;
  var lastJson = '';
  var pass = store(PASS_KEY) || '';
  var timer = null;

  function setStatus(msg, bad) {
    var el = $('status');
    if (!el) return;
    el.textContent = msg || '';
    el.className = 'status' + (bad ? ' bad' : '');
  }

  function listHTML(editable) {
    if (!guests.length) return '<div class="empty">尚無來賓簽到</div>';
    return guests.map(function (g, i) {
      return '<div class="guest">' +
        '<div class="num">' + (i + 1) + '</div>' +
        '<div class="info">' +
          '<div class="who"><span class="name">' + esc(g.name) + '</span>' +
            (g.title ? '<span class="title">' + esc(g.title) + '</span>' : '') + '</div>' +
          '<div class="time">簽到時間 ' + esc(g.time) + '</div>' +
        '</div>' +
        (editable ? '<div class="actions">' +
          '<button type="button" data-act="up" data-id="' + esc(g.id) + '"' + (i === 0 ? ' disabled' : '') + '>上移</button>' +
          '<button type="button" data-act="down" data-id="' + esc(g.id) + '"' + (i === guests.length - 1 ? ' disabled' : '') + '>下移</button>' +
          '<button type="button" data-act="edit" data-id="' + esc(g.id) + '">編輯</button>' +
          '<button type="button" class="danger" data-act="del" data-id="' + esc(g.id) + '">刪除</button>' +
        '</div>' : '') +
      '</div>';
    }).join('');
  }

  function render(force) {
    var json = JSON.stringify(guests);
    if (!force && json === lastJson) return;
    lastJson = json;
    var count = '目前共 ' + guests.length + ' 位來賓';
    $('count').textContent = MODE === 'view' ? count + '・每 4 秒自動更新' : count;
    $('list').innerHTML = listHTML(MODE === 'checkin');
  }

  function refresh() {
    var modal = $('modal');
    if (busy || (modal && !modal.hidden)) return;
    api.list().then(function (res) {
      if (busy) return;
      guests = res.guests;
      render();
    }).catch(function (err) {
      console.error(err);
      if (!guests.length) $('list').innerHTML = '<div class="empty">暫時無法讀取名單，稍後會自動重試。</div>';
    });
  }

  function startPolling() {
    refresh();
    if (!timer) timer = setInterval(refresh, POLL_MS);
  }

  /* ---------------- 簽到處 ---------------- */

  function mutate(body, okMsg) {
    busy = true;
    body.pass = pass;
    return api.post(body).then(function (res) {
      guests = res.guests;
      render(true);
      if (okMsg) setStatus(okMsg);
      return true;
    }).catch(function (err) {
      var msg = err && err.message || '儲存失敗，請再試一次。';
      setStatus(msg, true);
      if (/密碼/.test(msg)) logout(msg);
      return false;
    }).then(function (ok) { busy = false; return ok; });
  }

  function openModal(o) {
    $('modal-title').textContent = o.title;
    $('modal-body').innerHTML = o.body;
    var ok = $('modal-ok');
    ok.textContent = o.okText;
    ok.className = o.okClass;
    ok.disabled = false;
    $('modal').hidden = false;
    ok.onclick = function () {
      ok.disabled = true;
      o.onOk().then(function (done) {
        ok.disabled = false;
        if (done !== false) closeModal();
      });
    };
    $('modal-cancel').onclick = closeModal;
    setTimeout(function () {
      var f = $('modal-body').querySelector('input');
      if (f) f.focus();
    }, 0);
  }
  function closeModal() { $('modal').hidden = true; }

  function openEdit(g) {
    openModal({
      title: '編輯來賓', okText: '儲存', okClass: 'primary',
      body: '<label>姓名<input id="edit-name" value="' + esc(g.name) + '" maxlength="40"></label>' +
            '<label>職稱<input id="edit-title" value="' + esc(g.title) + '" maxlength="60"></label>' +
            '<div id="modal-err" class="err"></div>',
      onOk: function () {
        var name = $('edit-name').value.trim();
        var title = $('edit-title').value.trim();
        if (!name) { $('modal-err').textContent = '請輸入來賓姓名。'; return Promise.resolve(false); }
        return mutate({ action: 'edit', id: g.id, name: name, title: title }, '已更新「' + name + '」').then(function (ok) {
          if (!ok) $('modal-err').textContent = $('status').textContent;
          return ok;
        });
      }
    });
  }

  function openDelete(g) {
    openModal({
      title: '確認刪除', okText: '刪除', okClass: 'danger',
      body: '<p>確定要刪除來賓「' + esc(g.name) + '」' + (g.title ? '（' + esc(g.title) + '）' : '') + '嗎？</p>' +
            '<p style="font-size:14px;color:#5f6b7a">刪除後，後面來賓的序號會自動往前補。</p>',
      onOk: function () { return mutate({ action: 'delete', id: g.id }, '已刪除「' + g.name + '」'); }
    });
  }

  function showMain() {
    $('sec-login').hidden = true;
    $('sec-main').hidden = false;
    $('logout').hidden = false;
    startPolling();
    $('add-name').focus();
  }

  function logout(msg) {
    pass = '';
    store(PASS_KEY, null);
    closeModal();
    $('sec-main').hidden = true;
    $('logout').hidden = true;
    $('sec-login').hidden = false;
    $('login-status').textContent = typeof msg === 'string' ? msg : '';
    $('login-status').className = 'status' + (typeof msg === 'string' ? ' bad' : '');
    $('login-pass').value = '';
    $('login-pass').focus();
  }

  function initCheckin() {
    $('login-form').addEventListener('submit', function (e) {
      e.preventDefault();
      var p = $('login-pass').value.trim();
      if (!p) return;
      $('login-status').className = 'status';
      $('login-status').textContent = '確認中…';
      api.post({ action: 'check', pass: p }).then(function () {
        pass = p;
        store(PASS_KEY, p);
        showMain();
      }).catch(function (err) {
        $('login-status').className = 'status bad';
        $('login-status').textContent = err.message;
      });
    });
    $('logout').addEventListener('click', function () { logout(); });

    $('add-form').addEventListener('submit', function (e) {
      e.preventDefault();
      var name = $('add-name').value.trim();
      var title = $('add-title').value.trim();
      if (!name) { setStatus('請輸入來賓姓名。', true); return; }
      $('add-btn').disabled = true;
      setStatus('登錄中…');
      mutate({ action: 'add', name: name, title: title }, '已登錄「' + name + '」').then(function (ok) {
        $('add-btn').disabled = false;
        if (ok) {
          $('add-name').value = '';
          $('add-title').value = '';
          $('add-name').focus();
        }
      });
    });

    $('list').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-act]');
      if (!b || busy) return;
      var id = b.dataset.id, act = b.dataset.act;
      var g = guests.filter(function (x) { return x.id === id; })[0];
      if (!g) return;
      if (act === 'up' || act === 'down') mutate({ action: 'move', id: id, dir: act === 'up' ? -1 : 1 }, '已調整順序');
      else if (act === 'edit') openEdit(g);
      else if (act === 'del') openDelete(g);
    });

    if (pass) showMain(); else logout();
  }

  if (DEMO) $('demo').hidden = false;
  if (MODE === 'checkin') initCheckin();
  else startPolling();
})();
