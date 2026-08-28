
// ============================================================
// CONFIG - leave BOTH blank when Apps Script serves this page
// ============================================================
var API_URL   = '';   // e.g. 'https://script.google.com/macros/s/AKfy.../exec'
var API_TOKEN = '';   // must match the webAppToken row in Settings

// ============================================================
// API adapter - direct call when Apps Script serves us,
// fetch + token when we are hosted somewhere else.
// ============================================================
var NATIVE = (typeof google !== 'undefined' && google.script && google.script.run);

function api(fn, args){
  args = args || [];
  return new Promise(function(resolve, reject){
    if (NATIVE){
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(function(e){ reject(new Error(e && e.message ? e.message : String(e))); })
        [fn].apply(null, args);
      return;
    }
    if (!API_URL){
      reject(new Error('API_URL is empty. Either open this page from the Apps Script /exec URL, or fill in API_URL and API_TOKEN at the top of this file.'));
      return;
    }
    fetch(API_URL, {
      method: 'POST',
      // text/plain keeps the browser from sending a CORS preflight, which
      // Apps Script cannot answer.
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ token: API_TOKEN, fn: fn, args: args })
    })
    .then(function(r){ return r.json(); })
    .then(function(j){ j.ok ? resolve(j.data) : reject(new Error(j.error || 'Server error')); })
    .catch(function(e){ reject(new Error('Network: ' + e.message)); });
  });
}

// ============================================================
// helpers
// ============================================================
var $    = function(s){ return document.querySelector(s); };
var view = function(){ return $('#view'); };

function esc(s){
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function toast(msg, bad){
  var t = document.createElement('div');
  t.id = 'toast'; t.textContent = msg;
  if (bad) t.className = 'bad';
  var old = $('#toast'); if (old) old.remove();
  document.body.appendChild(t);
  setTimeout(function(){ if (t.parentNode) t.remove(); }, bad ? 4200 : 2300);
}
function busy(){ view().innerHTML = '<div class="loading"><i class="fa-solid fa-circle-notch spin"></i></div>'; }
function fail(e){
  view().innerHTML = '<div class="empty"><i class="fa-solid fa-triangle-exclamation"></i>' +
                     esc(e.message) + '</div>';
}
function nothing(icon, txt){
  return '<div class="empty"><i class="fa-solid fa-' + icon + '"></i>' + esc(txt) + '</div>';
}

// ============================================================
// state
// ============================================================
var TAB    = 'leads';
var LEADS  = null;
var SEQS   = null;
var FOUND  = null;   // Add-tab search results
var LOOKUP = '';     // Add-tab email searched

// ============================================================
// header
// ============================================================
function refreshHeader(){
  api('uiGetStats').then(function(s){
    $('#s1').textContent = s.sendsToday + '/' + s.maxSends;
    $('#s2').textContent = s.counts.Active;
    $('#s3').textContent = s.counts.Replied;
    var p = $('#pill');
    if (s.paused)            { p.textContent = 'Paused';  p.className = 'pill pause'; }
    else if (s.chainRunning) { p.textContent = 'Running'; p.className = 'pill run';   }
    else                     { p.textContent = 'Idle';    p.className = 'pill on';    }
  }).catch(function(e){ $('#pill').textContent = 'Error'; console.error(e); });
}

// ============================================================
// TAB: leads
// ============================================================
function renderLeads(){
  if (!LEADS) { busy(); api('uiGetLeads').then(function(d){ LEADS = d; renderLeads(); }).catch(fail); return; }
  if (!LEADS.length) { view().innerHTML = nothing('users', 'No leads yet. Use the Add tab.'); return; }

  var order = { Active:0, Error:1, OOO:2, Bounced:3, Replied:4, Paused:5, Done:6 };
  var sorted = LEADS.slice().sort(function(a,b){
    var d = (order[a.status] === undefined ? 9 : order[a.status]) -
            (order[b.status] === undefined ? 9 : order[b.status]);
    return d !== 0 ? d : String(a.nextSendDate).localeCompare(String(b.nextSendDate));
  });

  view().innerHTML = sorted.map(function(l, i){
    return '<div class="card tap" data-lead="' + i + '">' +
      '<div class="row">' +
        '<div class="grow">' +
          '<div class="nm ell">' + esc(l.leadName || l.leadEmail) + '</div>' +
          '<div class="sub ell">' + esc(l.leadName ? l.leadEmail : (l.sequenceName || 'no sequence')) + '</div>' +
        '</div>' +
        '<span class="st st-' + esc(l.status || 'Done') + '">' + esc(l.status) + '</span>' +
      '</div>' +
      '<div class="meta">' +
        '<span><i class="fa-solid fa-list-ol"></i>Step ' + l.step +
              (l.totalSteps ? '/' + l.totalSteps : '') + '</span>' +
        (l.nextSendDate ? '<span><i class="fa-regular fa-calendar"></i>' + esc(l.nextSendDate) + '</span>' : '') +
        (l.sequenceName && l.leadName ? '<span><i class="fa-solid fa-layer-group"></i>' + esc(l.sequenceName) + '</span>' : '') +
      '</div>' +
    '</div>';
  }).join('');

  view().querySelectorAll('[data-lead]').forEach(function(el){
    el.onclick = function(){ openLead(sorted[+el.dataset.lead]); };
  });
}

function openLead(l){
  var STATUSES = ['Active','Paused','Replied','OOO','Bounced','Done','Error'];
  var html =
    '<div class="sheet" id="ov"><div class="sheet-in">' +
      '<div class="sheet-hd">' +
        '<h3>' + esc(l.leadName || l.leadEmail) + '</h3>' +
        '<button class="x" id="ovx"><i class="fa-solid fa-xmark"></i></button>' +
      '</div>' +
      '<div class="kv"><span>Email</span><div>' + esc(l.leadEmail) + '</div></div>' +
      '<div class="kv"><span>Sequence</span><div>' + esc(l.sequenceName || '—') + '</div></div>' +
      '<div class="kv"><span>Step</span><div>' + l.step + (l.totalSteps ? ' of ' + l.totalSteps : '') + '</div></div>' +
      '<div class="kv"><span>Next send</span><div>' + esc(l.nextSendDate || '—') + '</div></div>' +
      '<div class="kv"><span>Last sent</span><div>' + esc(l.lastSentDate || 'never') + '</div></div>' +
      (l.snippet ? '<div class="kv"><span>Their reply</span><div>' + esc(l.snippet) + '</div></div>' : '') +
      (l.notes   ? '<div class="kv"><span>Notes</span><div>' + esc(l.notes) + '</div></div>' : '') +
      '<h2 class="sec">Change status</h2>' +
      '<div class="btnrow">' +
        STATUSES.map(function(s){
          return '<button class="btn ghost sm" data-st="' + s + '"' +
                 (s === l.status ? ' disabled' : '') + '>' + s + '</button>';
        }).join('') +
      '</div>' +
      '<p class="muted" style="margin:13px 0 0;font-size:12px">' +
        'Setting a lead back to Active also makes it due today.</p>' +
    '</div></div>';

  var wrap = document.createElement('div');
  wrap.innerHTML = html;
  document.body.appendChild(wrap.firstChild);

  var close = function(){ var o = $('#ov'); if (o) o.remove(); };
  $('#ovx').onclick = close;
  $('#ov').onclick  = function(e){ if (e.target.id === 'ov') close(); };

  $('#ov').querySelectorAll('[data-st]').forEach(function(b){
    b.onclick = function(){
      var st = b.dataset.st;
      $('#ov').querySelectorAll('[data-st]').forEach(function(x){ x.disabled = true; });
      b.innerHTML = '<i class="fa-solid fa-circle-notch spin"></i>';
      api('uiSetLeadStatus', [l.threadId, st]).then(function(d){
        LEADS = d; close(); renderLeads(); refreshHeader();
        toast('Set to ' + st);
      }).catch(function(e){ close(); toast(e.message, true); });
    };
  });
}

// ============================================================
// TAB: add lead
// ============================================================
function renderAdd(){
  var seqOpts = (SEQS || []).map(function(s){
    return '<option value="' + esc(s.name) + '">' + esc(s.name) + ' (' + s.steps.length + ')</option>';
  }).join('');

  view().innerHTML =
    '<div class="field">' +
      '<label>Lead email</label>' +
      '<input id="q" type="email" inputmode="email" autocapitalize="off" autocorrect="off" ' +
             'placeholder="name@company.com" value="' + esc(LOOKUP) + '">' +
    '</div>' +
    '<button class="btn" id="go"><i class="fa-solid fa-magnifying-glass"></i>Find threads</button>' +
    '<div id="res"></div>' +
    '<datalist id="seqlist">' + seqOpts + '</datalist>';

  $('#go').onclick = doSearch;
  $('#q').onkeydown = function(e){ if (e.key === 'Enter') doSearch(); };
  if (FOUND) paintResults();
}

function doSearch(){
  var q = $('#q').value.trim();
  if (!q) { toast('Enter an email address', true); return; }
  LOOKUP = q; FOUND = null;
  $('#res').innerHTML = '<div class="loading"><i class="fa-solid fa-circle-notch spin"></i></div>';
  $('#go').disabled = true;
  api('uiSearchThreads', [q]).then(function(d){
    FOUND = d; $('#go').disabled = false; paintResults();
  }).catch(function(e){
    $('#go').disabled = false;
    $('#res').innerHTML = nothing('triangle-exclamation', e.message);
  });
}

function paintResults(){
  if (!FOUND.length){ $('#res').innerHTML = nothing('inbox', 'No Gmail threads with ' + LOOKUP); return; }
  $('#res').innerHTML =
    '<h2 class="sec">' + FOUND.length + ' thread' + (FOUND.length > 1 ? 's' : '') + '</h2>' +
    FOUND.map(function(t, i){
      return '<div class="card' + (t.alreadyIn ? '' : ' tap') + '" data-th="' + i + '">' +
        '<div class="row">' +
          '<div class="grow"><div class="nm ell">' + esc(t.subject || '(no subject)') + '</div></div>' +
          (t.alreadyIn ? '<span class="st st-' + esc(t.alreadyIn) + '">' + esc(t.alreadyIn) + '</span>'
                       : '<i class="fa-solid fa-chevron-right" style="color:var(--faint);font-size:12px"></i>') +
        '</div>' +
        '<div class="sub" style="white-space:normal">' + esc(t.body.substring(0, 150)) + '…</div>' +
        '<div class="meta">' +
          '<span><i class="fa-regular fa-clock"></i>' + esc(t.date) + '</span>' +
          '<span><i class="fa-regular fa-envelope"></i>' + t.msgCount + ' msg</span>' +
        '</div>' +
      '</div>';
    }).join('');

  $('#res').querySelectorAll('[data-th]').forEach(function(el){
    var t = FOUND[+el.dataset.th];
    if (t.alreadyIn) return;
    el.onclick = function(){ openEnroll(t); };
  });
}

function openEnroll(t){
  var seqOpts = (SEQS || []).map(function(s){
    return '<option value="' + esc(s.name) + '">' + esc(s.name) + ' — ' + s.steps.length + ' steps</option>';
  }).join('');

  var html =
    '<div class="sheet" id="ov"><div class="sheet-in">' +
      '<div class="sheet-hd">' +
        '<h3>Enroll lead</h3>' +
        '<button class="x" id="ovx"><i class="fa-solid fa-xmark"></i></button>' +
      '</div>' +
      '<div class="kv"><span>Thread</span><div>' + esc(t.subject || '(no subject)') + '</div></div>' +
      '<div class="kv" style="margin-bottom:16px"><span>Email</span><div>' + esc(LOOKUP) + '</div></div>' +
      '<div class="field">' +
        '<label>Lead name</label>' +
        '<input id="en" placeholder="John Smith">' +
        '<p class="muted" style="margin:6px 0 0;font-size:11.5px">' +
          'Used by {{firstName}}. Leave blank only if your messages do not use it.</p>' +
      '</div>' +
      '<div class="field">' +
        '<label>Sequence</label>' +
        (seqOpts ? '<select id="es"><option value="">— pick one —</option>' + seqOpts + '</select>'
                 : '<input id="es" placeholder="sequence name">') +
      '</div>' +
      '<button class="btn" id="ok"><i class="fa-solid fa-check"></i>Enroll</button>' +
    '</div></div>';

  var wrap = document.createElement('div');
  wrap.innerHTML = html;
  document.body.appendChild(wrap.firstChild);

  var close = function(){ var o = $('#ov'); if (o) o.remove(); };
  $('#ovx').onclick = close;
  $('#ov').onclick  = function(e){ if (e.target.id === 'ov') close(); };

  $('#ok').onclick = function(){
    var nm = $('#en').value.trim();
    var sq = $('#es').value.trim();
    if (!sq){ toast('Pick a sequence', true); return; }
    $('#ok').disabled = true;
    $('#ok').innerHTML = '<i class="fa-solid fa-circle-notch spin"></i>';
    api('uiEnrollLead', [t.threadId, LOOKUP, nm, sq, 0]).then(function(d){
      LEADS = d; close();
      t.alreadyIn = 'Active'; paintResults();
      refreshHeader();
      toast('Enrolled — first send tomorrow');
    }).catch(function(e){
      $('#ok').disabled = false;
      $('#ok').innerHTML = '<i class="fa-solid fa-check"></i>Enroll';
      toast(e.message, true);
    });
  };
}

// ============================================================
// TAB: settings
// ============================================================
function renderSettings(){
  busy();
  api('uiGetSettings').then(function(list){
    var pause = list.filter(function(s){ return s.name === 'pauseAllFollowUps'; })[0];
    var isOn  = pause && pause.value.toUpperCase() === 'TRUE';

    view().innerHTML =
      '<div class="switch' + (isOn ? ' armed' : '') + '">' +
        '<div class="grow">' +
          '<b>' + (isOn ? 'Follow-ups paused' : 'Follow-ups running') + '</b>' +
          '<div class="muted" style="font-size:12.5px">' +
            (isOn ? 'Nothing will be sent until you switch this off'
                  : 'Emergency stop for every lead') + '</div>' +
        '</div>' +
        '<button class="knob' + (isOn ? ' on' : '') + '" id="pz"></button>' +
      '</div>' +
      '<button class="btn ghost" id="run"><i class="fa-solid fa-play"></i>Run chain now</button>' +
      '<h2 class="sec">All settings</h2>' +
      list.filter(function(s){ return s.name !== 'pauseAllFollowUps'; }).map(function(s){
        return '<div class="field">' +
          '<label>' + esc(s.name) + '</label>' +
          '<input data-set="' + esc(s.name) + '" value="' + esc(s.value) + '"' +
                 (s.name === 'webAppToken' ? ' disabled' : '') + '>' +
        '</div>';
      }).join('') +
      '<p class="muted" style="font-size:12px;margin-top:-4px">Changes save when you tap away from a field.</p>';

    $('#pz').onclick = function(){
      var next = isOn ? 'FALSE' : 'TRUE';
      $('#pz').disabled = true;
      api('uiSaveSetting', ['pauseAllFollowUps', next]).then(function(){
        renderSettings(); refreshHeader();
        toast(next === 'TRUE' ? 'Everything paused' : 'Follow-ups resumed');
      }).catch(function(e){ $('#pz').disabled = false; toast(e.message, true); });
    };

    $('#run').onclick = function(){
      $('#run').disabled = true;
      $('#run').innerHTML = '<i class="fa-solid fa-circle-notch spin"></i>Starting';
      api('uiRunChainNow').then(function(m){
        toast(m); refreshHeader();
        $('#run').disabled = false;
        $('#run').innerHTML = '<i class="fa-solid fa-play"></i>Run chain now';
      }).catch(function(e){
        $('#run').disabled = false;
        $('#run').innerHTML = '<i class="fa-solid fa-play"></i>Run chain now';
        toast(e.message, true);
      });
    };

    view().querySelectorAll('[data-set]').forEach(function(inp){
      var original = inp.value;
      inp.onblur = function(){
        if (inp.value === original || inp.disabled) return;
        var want = inp.value;
        api('uiSaveSetting', [inp.dataset.set, want]).then(function(){
          original = want; refreshHeader(); toast('Saved ' + inp.dataset.set);
        }).catch(function(e){ inp.value = original; toast(e.message, true); });
      };
    });
  }).catch(fail);
}

// ============================================================
// TAB: log
// ============================================================
function renderLog(){
  busy();
  api('uiGetLog', [50]).then(function(rows){
    if (!rows.length){ view().innerHTML = nothing('receipt', 'Nothing sent yet.'); return; }
    var tone = function(r){
      if (r === 'Sent') return 'st-Replied';
      if (r === 'Failed' || r === 'Sent-WriteError') return 'st-Error';
      return 'st-Done';
    };
    view().innerHTML = rows.map(function(r){
      return '<div class="card">' +
        '<div class="row">' +
          '<div class="grow"><div class="nm ell" style="font-size:13.5px">' + esc(r.email) + '</div></div>' +
          '<span class="st ' + tone(r.result) + '">' + esc(r.result.replace('Skipped - ','')) + '</span>' +
        '</div>' +
        (r.preview ? '<div class="sub ell">' + esc(r.preview) + '</div>' : '') +
        '<div class="meta">' +
          '<span><i class="fa-regular fa-clock"></i>' + esc(r.time) + '</span>' +
          (r.sequence ? '<span><i class="fa-solid fa-layer-group"></i>' + esc(r.sequence) +
                        ' · ' + esc(r.step) + '</span>' : '') +
        '</div>' +
      '</div>';
    }).join('');
  }).catch(fail);
}

// ============================================================
// TAB: sequences (read only)
// ============================================================
function renderSeq(){
  if (!SEQS){ busy(); loadSeqs().then(renderSeq).catch(fail); return; }
  if (!SEQS.length){ view().innerHTML = nothing('layer-group', 'No sequences in the Sequences sheet.'); return; }

  view().innerHTML = SEQS.map(function(s){
    return '<h2 class="sec">' + esc(s.name) + ' · ' + s.steps.length + ' steps</h2>' +
      s.steps.map(function(st){
        return '<div class="card">' +
          '<div class="row" style="margin-bottom:9px">' +
            '<div class="grow"><b style="font-size:13px">Step ' + st.n + '</b></div>' +
            '<span class="st st-Done">wait ' + st.awaitDays + 'd</span>' +
          '</div>' +
          '<pre class="msg">' + esc(st.message) + '</pre>' +
        '</div>';
      }).join('');
  }).join('') +
  '<p class="muted" style="margin-top:16px;font-size:12.5px">' +
    'Read-only here — editing long message text is far easier in the Sheet.</p>';
}

function loadSeqs(){
  return api('uiGetSequences').then(function(d){ SEQS = d; return d; });
}

// ============================================================
// routing
// ============================================================
var TABS = { leads:renderLeads, add:renderAdd, settings:renderSettings, log:renderLog, seq:renderSeq };

function go(tab){
  TAB = tab;
  document.querySelectorAll('nav button').forEach(function(b){
    b.classList.toggle('on', b.dataset.tab === tab);
  });
  window.scrollTo(0, 0);
  TABS[tab]();
}

document.querySelectorAll('nav button').forEach(function(b){
  b.onclick = function(){ go(b.dataset.tab); };
});

// boot: sequences are needed by the Add tab's picker, so fetch them up front
loadSeqs().catch(function(){ SEQS = []; });
refreshHeader();
go('leads');
