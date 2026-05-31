/**
 * thongke.js – Thống kê GDQP (BLD + BGD)
 * Được tải SAU inline script của index.html, ghi đè các hàm cũ.
 *
 * Luồng lọc 1: Trường → Năm → Học kỳ → Tự chọn tất cả lớp → Xuất
 * Luồng lọc 2: Trường → Năm → Chọn lớp (nhiều lớp, lọc theo chế độ) → Xuất
 */

/* ─────────────────────────────────────────────────────────────
   Hằng số
───────────────────────────────────────────────────────────── */
var BLD_HP_MAP = {
  /* Dạng ngắn (190081...) */
  '190081':1, '190082':2, '190083':3, '190084':4,
  /* Dạng dài (1900081...) như đặc tả */
  '1900081':1, '1900082':2, '1900083':3, '190004':4
};
var BLD_SUBJECTS = (function(){
  var s = {}; Object.keys(BLD_HP_MAP).forEach(function(k){ s[k]=true; }); return s;
})();
var BGD_SUBJECTS_LIST = ['190036','190008','195001'];
var BGD_SUBJECTS = (function(){
  var s = {}; BGD_SUBJECTS_LIST.forEach(function(k){ s[k]=true; }); return s;
})();

var TK_SS_STATE = 'gdqp_tk_ui_state';

/* ─────────────────────────────────────────────────────────────
   Đảm bảo biến toàn cục tồn tại
───────────────────────────────────────────────────────────── */
function ensureTKGlobals(){
  if(typeof tkMode === 'undefined')     { window.tkMode     = 'bld'; }
  if(typeof tkByLopBLD === 'undefined') { window.tkByLopBLD = {}; }
  if(typeof tkByLopBGD === 'undefined') { window.tkByLopBGD = {}; }
  if(typeof tkByLop === 'undefined')    { window.tkByLop    = {}; }
  if(typeof tkAllClasses === 'undefined'){ window.tkAllClasses = []; }
  if(typeof tkSelected === 'undefined') { window.tkSelected  = new Set(); }
  if(typeof tkRawRows === 'undefined')  { window.tkRawRows   = []; }
  if(typeof tkCols === 'undefined')     { window.tkCols      = null; }
}

/* ─────────────────────────────────────────────────────────────
   Chuẩn hóa
───────────────────────────────────────────────────────────── */
function normalizeSchoolText(s){
  return String(s||'')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase().replace(/đ/g,'d').replace(/\s+/g,'');
}

function normalizeSemesterValue(v){
  var s = String(v||'').trim();
  if(!s) return '';
  var m = s.match(/\d+/);
  return m ? m[0] : s;
}

function getSemesterFromRow(row){
  if(!row || !tkCols) return '';
  var key = tkCols.semester || tkCols.hocKy;
  if(!key || row[key] == null) return '';
  return normalizeSemesterValue(row[key]);
}

function normalizeSubjectCode(v){
  return String(v||'').trim();
}

function normalizeTKResultCode(raw){
  var t = String(raw||'').trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,'');
  if(t==='CT'||t==='CAMTHI')                return 'CT';
  if(t==='VT'||t==='VANGTHI')               return 'VT';
  if(t==='DC'||t==='DINHCHI'||t==='ĐC')     return 'DC';
  return '';
}

function parseTKScoreValue(rawValue){
  var raw = (rawValue===null||rawValue===undefined) ? '' : String(rawValue).trim();
  if(!raw) return {raw:'', score:null};
  // Kiểm tra mã đặc biệt trước
  var code = normalizeTKResultCode(raw);
  if(code) return {raw:code, score:null};
  var n = parseFloat(raw.replace(',','.'));
  return {raw: raw, score: Number.isFinite(n) ? n : null};
}

function getTKResultPriority(parsed){
  var code = normalizeTKResultCode(parsed && parsed.raw);
  if(code==='CT') return 100;
  if(code==='DC') return 90;
  if(code==='VT') return 80;
  if(parsed && Number.isFinite(parsed.score)) return 10 + Number(parsed.score||0);
  return 0;
}

function pct(n, total){
  return total ? ((n*100)/total).toFixed(2) : '0.00';
}

function sanitizeFileNamePart(text){
  return String(text || '')
    .trim()
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/* ─────────────────────────────────────────────────────────────
   Lấy điểm từ hàng (ưu tiên cột T1_ĐTK L1)
───────────────────────────────────────────────────────────── */
function getTKScoreFromRow(row){
  if(!row || !tkCols) return {raw:'', score:null};
  var candidates = [];
  if(tkCols.t1)    candidates.push(tkCols.t1);
  if(tkCols.score) candidates.push(tkCols.score);
  for(var i=0; i<candidates.length; i++){
    var key = candidates[i];
    if(row[key]!==null && row[key]!==undefined){
      var parsed = parseTKScoreValue(row[key]);
      if(parsed.raw !== '') return parsed;
    }
  }
  return {raw:'', score:null};
}

/* ─────────────────────────────────────────────────────────────
   Xây dữ liệu BLD
   – Mỗi SV cần đủ 4 mã môn BLD_HP_MAP
   – Điểm lấy từ T1_ĐTK L1
   – TB = (hp1*3 + hp2*2 + hp3*1 + hp4*2) / 8
───────────────────────────────────────────────────────────── */
function buildBLDDataByClass(rows){
  var byClass = {};
  rows.forEach(function(r){
    if(!tkCols) return;
    var maMH = tkCols.maMH && r[tkCols.maMH]!=null
      ? normalizeSubjectCode(r[tkCols.maMH]) : '';
    var hp = BLD_HP_MAP[maMH];
    if(!hp) return;

    var maSV = tkCols.maSV && r[tkCols.maSV]!=null
      ? String(r[tkCols.maSV]).trim() : '';
    if(!maSV) return;

    var maLop = tkCols.maLop && r[tkCols.maLop]!=null
      ? String(r[tkCols.maLop]).trim() : 'Không rõ';

    if(!byClass[maLop]) byClass[maLop] = {students:{}};
    if(!byClass[maLop].students[maSV]){
      byClass[maLop].students[maSV] = {
        maSV: maSV, maLop: maLop,
        hoLot: tkCols.hoLot ? (r[tkCols.hoLot]||'') : '',
        ten:   tkCols.ten   ? (r[tkCols.ten]  ||'') : '',
        hp1:null, hp2:null, hp3:null, hp4:null,
        hp1r:'',  hp2r:'',  hp3r:'',  hp4r:''
      };
    }

    var sv = byClass[maLop].students[maSV];
    var parsed = getTKScoreFromRow(r);
    var old    = {raw: sv['hp'+hp+'r'], score: sv['hp'+hp]};
    if(parsed.raw !== '' && getTKResultPriority(parsed) >= getTKResultPriority(old)){
      sv['hp'+hp]  = parsed.score;
      sv['hp'+hp+'r'] = parsed.raw;
    }
  });

  Object.keys(byClass).forEach(function(lop){
    byClass[lop].students = Object.values(byClass[lop].students);
  });
  return byClass;
}

/* ─────────────────────────────────────────────────────────────
   Xây dữ liệu BGD
   – Mỗi hàng (SV × MaMH) là 1 bản ghi riêng
   – Điểm lấy thẳng từ T1_ĐTK L1
───────────────────────────────────────────────────────────── */
function buildBGDDataByClass(rows){
  var byClass = {};
  rows.forEach(function(r){
    if(!tkCols) return;
    var maMH = tkCols.maMH && r[tkCols.maMH]!=null
      ? normalizeSubjectCode(r[tkCols.maMH]) : '';
    if(!BGD_SUBJECTS[maMH]) return;

    var maSV = tkCols.maSV && r[tkCols.maSV]!=null
      ? String(r[tkCols.maSV]).trim() : '';
    if(!maSV) return;

    var maLop = tkCols.maLop && r[tkCols.maLop]!=null
      ? String(r[tkCols.maLop]).trim() : 'Không rõ';
    var key = maSV + '|' + maMH;

    if(!byClass[maLop]) byClass[maLop] = {students:{}};
    var parsed = getTKScoreFromRow(r);
    var next = {
      maSV: maSV, maLop: maLop,
      hoLot: tkCols.hoLot ? (r[tkCols.hoLot]||'') : '',
      ten:   tkCols.ten   ? (r[tkCols.ten]  ||'') : '',
      maMH:  maMH,
      score: parsed.score,
      raw:   parsed.raw
    };
    var old = byClass[maLop].students[key];
    if(!old || getTKResultPriority(next) >= getTKResultPriority(old)){
      byClass[maLop].students[key] = next;
    }
  });

  Object.keys(byClass).forEach(function(lop){
    byClass[lop].students = Object.values(byClass[lop].students);
  });
  return byClass;
}

/* ─────────────────────────────────────────────────────────────
   Phân loại BLD
───────────────────────────────────────────────────────────── */
function isCodeBLD(s, h){ return s[h+'r'] !== '' && s[h] === null; }

function calcBLDTB(s){
  if(isCodeBLD(s,'hp1')||isCodeBLD(s,'hp2')||isCodeBLD(s,'hp3')||isCodeBLD(s,'hp4')) return null;
  if(s.hp1==null||s.hp2==null||s.hp3==null||s.hp4==null) return null;
  return Math.round((s.hp1*3 + s.hp2*2 + s.hp3 + s.hp4*2) / 8 * 10) / 10;
}

function hasAnyBLDScoreBelow5(s){
  var scores = [Number(s.hp1), Number(s.hp2), Number(s.hp3), Number(s.hp4)];
  for(var i = 0; i < scores.length; i++){
    if(Number.isFinite(scores[i]) && scores[i] < 5) return true;
  }
  return false;
}

function isHpEmptyOrZero(score, raw) {
  if (score === null || score === undefined) return true;
  if (score === 0) return true;
  var r = String(raw || '').trim();
  if (r === '' || r === '0') return true;
  return false;
}

function categoryBLD(s){
  var codes = [s.hp1r, s.hp2r, s.hp3r, s.hp4r].map(normalizeTKResultCode);
  
  var allEmptyOrZero = isHpEmptyOrZero(s.hp1, s.hp1r) &&
                       isHpEmptyOrZero(s.hp2, s.hp2r) &&
                       isHpEmptyOrZero(s.hp3, s.hp3r) &&
                       isHpEmptyOrZero(s.hp4, s.hp4r);
  if(allEmptyOrZero) return 'chua_hoc';
  
  // Hỏng (VT hoặc điểm dưới 5) có mức ưu tiên cao nhất trong các lỗi
  if(codes.indexOf('VT') >= 0) return 'hong';
  if(hasAnyBLDScoreBelow5(s)) return 'hong';
  
  // Đình chỉ (DC)
  if(codes.indexOf('DC') >= 0) return 'dc';
  
  // Cấm thi (CT)
  if(codes.indexOf('CT') >= 0) return 'ct';
  
  var tb = calcBLDTB(s);
  if(tb === null) return 'hong';
  
  if(tb >= 9)  return 'sx';
  if(tb >= 8)  return 'g';
  if(tb >= 7)  return 'kha';
  if(tb >= 5)  return 'tb';
  return 'hong';
}

/* ─────────────────────────────────────────────────────────────
   Phân loại BGD
───────────────────────────────────────────────────────────── */
function categoryBGD(s){
  var code = normalizeTKResultCode(s.raw);
  if(code==='CT') return 'ct';
  if(code==='DC') return 'dc';
  if(code==='VT') return 'hong';
  
  var rawStr = String(s.raw || '').trim();
  var isZeroOrEmpty = (s.score === null || s.score === undefined || s.score === 0 || rawStr === '' || rawStr === '0');
  if (isZeroOrEmpty) {
    return 'chua_hoc';
  }
  
  var score = Number(s.score);
  if(!Number.isFinite(score)) return 'hong';
  if(score >= 9)  return 'sx';
  if(score >= 8)  return 'g';
  if(score >= 7)  return 'kha';
  if(score >= 5)  return 'tb';
  return 'hong';
}

/* ─────────────────────────────────────────────────────────────
   Thống kê theo lớp / theo lựa chọn
───────────────────────────────────────────────────────────── */
function classStats(k){
  var out = {sx:0,g:0,kha:0,tb:0,ct:0,vt:0,dc:0,hong:0,chua_hoc:0,total:0};
  var grp = tkByLop[k];
  if(!grp || !Array.isArray(grp.students)) return out;
  grp.students.forEach(function(s){
    var c = (tkMode === 'bld') ? categoryBLD(s) : categoryBGD(s);
    out[c]++; out.total++;
  });
  return out;
}

function statsFromSelection(){
  var out = {sx:0,g:0,kha:0,tb:0,ct:0,vt:0,dc:0,hong:0,chua_hoc:0,total:0};
  var arr = Array.isArray(tkSelected) ? tkSelected : [...tkSelected];
  arr.forEach(function(k){
    var c = classStats(k);
    Object.keys(out).forEach(function(key){ out[key] += (c[key]||0); });
  });
  return out;
}

/* ─────────────────────────────────────────────────────────────
   Chế độ lọc (học kỳ vs lớp)
───────────────────────────────────────────────────────────── */
function getTKFilterMode(){
  var el = document.getElementById('tkFilterType');
  return (el && el.value === 'class') ? 'class' : 'semester';
}

/* ─────────────────────────────────────────────────────────────
   Danh sách lớp được lọc theo ô tìm kiếm
───────────────────────────────────────────────────────────── */
function normStr(s){
  return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase().replace(/đ/g,'d').replace(/\s+/g,'');
}

function getFilteredClasses(){
  var searchEl = document.getElementById('tkSearch');
  var q = normStr(searchEl ? (searchEl.value||'') : '');
  return q ? tkAllClasses.filter(function(k){ return normStr(k).indexOf(q) >= 0; }) : tkAllClasses;
}

/* ─────────────────────────────────────────────────────────────
   Đồng bộ dataset theo chế độ (BLD / BGD)
───────────────────────────────────────────────────────────── */
function syncTKDatasetByMode(){
  tkByLop      = (tkMode === 'bld') ? tkByLopBLD : tkByLopBGD;
  tkAllClasses = Object.keys(tkByLop).sort();
  tkSelected   = new Set([...tkSelected].filter(function(k){ return tkAllClasses.indexOf(k) >= 0; }));

  var filterMode = getTKFilterMode();
  if(filterMode === 'semester'){
    tkSelected = new Set(tkAllClasses);
  }

  renderTKClassList();
  updateTKSummary();
  updateTKMetaText();
}

/* ─────────────────────────────────────────────────────────────
   ÁP DỤNG BỘ LỌC – Hàm chính
   Luồng 1: Trường → Năm → Học kỳ  → auto-chọn tất cả lớp
   Luồng 2: Trường → Năm → Lớp     → người dùng chọn lớp
───────────────────────────────────────────────────────────── */
function applyTKFilters(){
  ensureTKGlobals();

  if(!tkCols || !Array.isArray(tkRawRows) || !tkRawRows.length){
    tkByLop = {}; tkByLopBLD = {}; tkByLopBGD = {};
    tkAllClasses = []; tkSelected = new Set();
    renderTKClassList(); updateTKSummary(); updateTKMetaText();
    return;
  }

  var filterMode  = getTKFilterMode();
  var schoolEl    = document.getElementById('tkSchoolSelect');
  var yearEl      = document.getElementById('tkYearSelect');
  var semesterEl  = document.getElementById('tkSemesterSelect');

  var selSchool   = String((schoolEl   ? schoolEl.value   : '')||'').trim();
  var selYear     = String((yearEl     ? yearEl.value     : '')||'').trim();
  var selSemester = (filterMode === 'semester')
    ? String((semesterEl ? semesterEl.value : '')||'').trim()
    : '';

  /* Bước 1: lọc hàng theo Trường + Năm + (Học kỳ nếu luồng 1) */
  var filteredRows = tkRawRows.filter(function(r){
    if(selSchool){
      var school = (typeof getSchoolFromRow === 'function') ? getSchoolFromRow(r) : '';
      if(normalizeSchoolText(school) !== normalizeSchoolText(selSchool)) return false;
    }
    if(selYear){
      if(!tkCols.year) return false;
      var y = (r[tkCols.year] != null) ? String(r[tkCols.year]).trim() : '';
      if(y !== selYear) return false;
    }
    if(selSemester){
      var hk = getSemesterFromRow(r);
      if(hk !== selSemester) return false;
    }
    return true;
  });

  /* Bước 2: xây dữ liệu cả 2 chế độ từ hàng đã lọc */
  tkByLopBLD = buildBLDDataByClass(filteredRows);
  tkByLopBGD = buildBGDDataByClass(filteredRows);

  /* Bước 3: chọn dataset theo chế độ hiện tại */
  tkByLop      = (tkMode === 'bld') ? tkByLopBLD : tkByLopBGD;
  tkAllClasses = Object.keys(tkByLop).sort();

  /* Bước 4: xử lý lựa chọn lớp */
  if(filterMode === 'semester'){
    /* Luồng 1: tự chọn tất cả */
    tkSelected = new Set(tkAllClasses);
  } else {
    /* Luồng 2: giữ những lớp đã chọn còn hợp lệ */
    tkSelected = new Set([...tkSelected].filter(function(k){
      return tkAllClasses.indexOf(k) >= 0;
    }));
  }

  renderTKClassList();
  updateTKSummary();
  updateTKMetaText();
  saveTKState();
}

/* ─────────────────────────────────────────────────────────────
   Giao diện: danh sách lớp
───────────────────────────────────────────────────────────── */
function renderTKClassList(){
  var listEl = document.getElementById('tkClassList');
  if(!listEl) return;
  listEl.innerHTML = '';

  var filterMode = getTKFilterMode();

  /* Khi lọc theo học kỳ: không hiển thị checkbox lớp */
  if(filterMode !== 'class'){
    listEl.innerHTML = '<div style="font-size:13px;color:#55617e;padding:10px 6px;text-align:center">'
      + 'Đang dùng luồng <b>Học kỳ</b>.<br>'
      + 'Tất cả lớp phù hợp đã được tự động chọn.<br>'
      + '<span style="color:#1a7a44">✓ ' + tkSelected.size + ' lớp</span></div>';
    updateTKMetaText();
    return;
  }

  var filtered = getFilteredClasses();
  if(!filtered.length){
    var modeLabel = (tkMode === 'bld') ? 'BLD' : 'BGD';
    listEl.innerHTML = '<div style="font-size:13px;color:#55617e;padding:10px 6px;text-align:center">'
        + 'Không có ' + getTKPrimaryPlural() + ' nào có dữ liệu <b>' + modeLabel + '</b><br>'
      + 'theo Trường + Năm đã chọn.<br>'
      + '<span style="font-size:12px">(Thử đổi chế độ BLD/BGD hoặc nhấn <b>Tiến hành thống kê</b>)</span>'
      + '</div>';
    updateTKMetaText();
    return;
  }

  var frag = document.createDocumentFragment();
  filtered.forEach(function(k){
    var row = document.createElement('label');
    row.className = 'tk-item';
    row.style.cssText = 'display:flex;align-items:center;gap:7px;padding:5px 4px;border-radius:5px;cursor:pointer';
    row.onmouseover = function(){ row.style.background='#f0f4ff'; };
    row.onmouseout  = function(){ row.style.background=''; };

    var cb = document.createElement('input');
    cb.type    = 'checkbox';
    cb.checked = tkSelected.has(k);
    cb.addEventListener('change', function(){
      if(cb.checked) tkSelected.add(k); else tkSelected.delete(k);
      updateTKSummary();
      updateTKMetaText();
      saveTKState();
    });

    var span = document.createElement('span');
    span.textContent = k;
    span.style.cssText = 'font-size:13px;color:#1d3058';

    row.appendChild(cb);
    row.appendChild(span);
    frag.appendChild(row);
  });
  listEl.appendChild(frag);
  updateTKMetaText();
}

/* ─────────────────────────────────────────────────────────────
   Giao diện: thẻ lớp
───────────────────────────────────────────────────────────── */
function renderClassCards(){
  var el = document.getElementById('tkClassCards');
  if(!el) return;
  var keys = [...tkSelected].sort();
  if(!keys.length){
    el.innerHTML = '<div class="tk-class-card"><h6>Chưa có ' + getTKPrimaryPlural() + ' nào được chọn</h6>'
      + '<div class="r"><span>Tỷ lệ đạt</span><span>0%</span></div></div>';
    return;
  }
  el.innerHTML = keys.map(function(k){
    var s    = classStats(k);
    var pass = s.sx+s.g+s.kha+s.tb;
    var fail = s.ct+s.vt+s.dc+s.hong+s.chua_hoc;
    return '<div class="tk-class-card">'
      + '<h6 style="font-size:13px;margin-bottom:6px;color:#152d5f;word-break:break-all">' + k + '</h6>'
      + '<div class="r"><span>Tổng SV</span><span><b>' + s.total + '</b></span></div>'
      + '<div class="r" style="color:#1a7a44"><span>Đạt</span><span>' + pass + ' (' + pct(pass,s.total) + '%)</span></div>'
      + '<div class="r" style="color:#a83200"><span>Không đạt</span><span>' + fail + ' (' + pct(fail,s.total) + '%)</span></div>'
      + '<div class="r"><span>SX / G / Khá / TB</span><span style="font-size:12px">'
        + s.sx + '&nbsp;/&nbsp;' + s.g + '&nbsp;/&nbsp;' + s.kha + '&nbsp;/&nbsp;' + s.tb + '</span></div>'
      + '<div class="r"><span>CT / VT / ĐC / Hỏng / Chưa học</span><span style="font-size:12px">'
        + s.ct + '&nbsp;/&nbsp;' + s.vt + '&nbsp;/&nbsp;' + s.dc + '&nbsp;/&nbsp;' + s.hong + '&nbsp;/&nbsp;' + s.chua_hoc + '</span></div>'
      + '</div>';
  }).join('');
}

/* ─────────────────────────────────────────────────────────────
   Giao diện: bảng thống kê theo lớp
───────────────────────────────────────────────────────────── */
function renderClassStatsTable(){
  var tbody = document.getElementById('tkClassTableBody');
  if(!tbody) return;
  var keys = [...tkSelected].sort();
  if(!keys.length){
    tbody.innerHTML = '<tr><td class="l" colspan="12">Chưa chọn ' + getTKPrimaryPlural() + ' nào</td></tr>';
    return;
  }
  tbody.innerHTML = keys.map(function(k){
    var s    = classStats(k);
    var pass = s.sx+s.g+s.kha+s.tb;
    return '<tr>'
      + '<td class="l">' + k + '</td>'
      + '<td><b>' + s.total + '</b></td>'
      + '<td style="color:#0a6b2e">' + s.sx + '</td>'
      + '<td style="color:#1d7a44">' + s.g  + '</td>'
      + '<td style="color:#2a7d9b">' + s.kha + '</td>'
      + '<td style="color:#7b6000">' + s.tb  + '</td>'
      + '<td style="color:#c04000">' + s.ct  + '</td>'
      + '<td style="color:#b04200">' + s.vt  + '</td>'
      + '<td style="color:#8a0000">' + s.dc  + '</td>'
      + '<td style="color:#6b0000">' + s.hong + '</td>'
      + '<td style="color:#4a5568">' + s.chua_hoc + '</td>'
      + '<td><b>' + pct(pass,s.total) + '%</b></td>'
      + '</tr>';
  }).join('');
}

/* ─────────────────────────────────────────────────────────────
   Cập nhật KPI tổng hợp
───────────────────────────────────────────────────────────── */
function updateTKSummary(){
  var s    = statsFromSelection();
  var pass = s.sx+s.g+s.kha+s.tb;
  var fail = s.ct+s.vt+s.dc+s.hong+s.chua_hoc;

  function set(id, val){ var el=document.getElementById(id); if(el) el.textContent=val; }

  set('kpiTotal', s.total);
  set('kpiClass', tkSelected.size);
  set('kpiPass',  pct(pass,s.total)+'%');
  set('kpiFail',  pct(fail,s.total)+'%');

  set('tkSX',   s.sx   + ' (TL: ' + pct(s.sx,s.total)   + '%)');
  set('tkG',    s.g    + ' (TL: ' + pct(s.g,s.total)    + '%)');
  set('tkKha',  s.kha  + ' (TL: ' + pct(s.kha,s.total)  + '%)');
  set('tkTB',   s.tb   + ' (TL: ' + pct(s.tb,s.total)   + '%)');
  set('tkCT',   s.ct   + ' (TL: ' + pct(s.ct,s.total)   + '%)');
  set('tkVT',   s.vt   + ' (TL: ' + pct(s.vt,s.total)   + '%)');
  set('tkDC',   s.dc   + ' (TL: ' + pct(s.dc,s.total)   + '%)');
  set('tkHong', s.hong + ' (TL: ' + pct(s.hong,s.total) + '%)');
  set('tkChuaHoc', s.chua_hoc + ' (TL: ' + pct(s.chua_hoc,s.total) + '%)');

  renderClassCards();
  renderClassStatsTable();
}

/* ─────────────────────────────────────────────────────────────
   Hiển thị / ẩn bộ lọc học kỳ vs lớp
───────────────────────────────────────────────────────────── */
function updateTKFilterInputVisibility(){
  var mode      = getTKFilterMode();
  var semWrap   = document.getElementById('tkSemesterWrap');
  var classWrap = document.getElementById('tkClassFilterWrap');
  if(semWrap)   semWrap.style.display   = (mode === 'semester') ? '' : 'none';
  if(classWrap) classWrap.style.display = (mode === 'class')    ? '' : 'none';
  var modeSelect = document.getElementById('tkModeSelect');
  if(modeSelect) modeSelect.value = (tkMode === 'bgd') ? 'bgd' : 'bld';
}

function updateTKModeLabels(){
  var headEl  = document.getElementById('tkPrimaryColLabel');
  if(headEl)  headEl.textContent = getTKPrimaryLabel();
  var searchEl = document.getElementById('tkSearch');
  if(searchEl) searchEl.placeholder = (tkMode === 'bld') ? 'Tìm kiếm...' : 'Tìm mã lớp...';
  var exportBtn = document.getElementById('tkExportBanBtn');
  if(exportBtn) exportBtn.textContent = (tkMode === 'bld') ? 'Xuất DS CT/VT/DC/Hỏng' : 'Xuất DS VT + Hỏng';
  updateTKFilterInputVisibility();
}

/* ─────────────────────────────────────────────────────────────
   Meta text
───────────────────────────────────────────────────────────── */
function updateTKMetaText(){
  var modeText    = (tkMode === 'bld') ? 'BLD' : 'BGD';
  var filterMode  = getTKFilterMode();
  var schoolEl    = document.getElementById('tkSchoolSelect');
  var yearEl      = document.getElementById('tkYearSelect');
  var semEl       = document.getElementById('tkSemesterSelect');
  var selSchool   = (schoolEl && schoolEl.value)  ? schoolEl.value  : 'Tất cả trường';
  var selYear     = (yearEl   && yearEl.value)    ? yearEl.value    : 'Tất cả năm học';
  var selSem      = (semEl    && semEl.value)     ? 'HK' + semEl.value : 'Tất cả HK';
  var flowText    = (filterMode === 'semester')
    ? ('Luồng học kỳ: ' + selSem + ' – ' + tkSelected.size + ' lớp auto chọn')
    : ('Luồng lớp: ' + tkSelected.size + '/' + tkAllClasses.length + ' lớp đã chọn');
  var meta = document.getElementById('tkMeta');
  if(meta) meta.textContent = '[' + modeText + '] ' + flowText + ' | ' + selSchool + ' | ' + selYear;
}

/* ─────────────────────────────────────────────────────────────
   Chuyển chế độ BLD / BGD
───────────────────────────────────────────────────────────── */
function switchTKMode(mode){
  tkMode = (mode === 'bgd') ? 'bgd' : 'bld';
  updateTKModeLabels();
  syncTKDatasetByMode();
  saveTKState();
}

/* ─────────────────────────────────────────────────────────────
   Chọn / bỏ chọn lớp hàng loạt
───────────────────────────────────────────────────────────── */
function selectAllFilteredClasses(){
  if(getTKFilterMode() !== 'class') return;
  getFilteredClasses().forEach(function(k){ tkSelected.add(k); });
  renderTKClassList();
  updateTKSummary();
  saveTKState();
}

function clearSelectedClasses(){
  if(getTKFilterMode() !== 'class') return;
  tkSelected = new Set();
  renderTKClassList();
  updateTKSummary();
  saveTKState();
}

/* ─────────────────────────────────────────────────────────────
   Nút "Tiến hành thống kê"
───────────────────────────────────────────────────────────── */
function runTKStatistics(){
  if(!Array.isArray(tkRawRows) || !tkRawRows.length || !tkCols){
    if(typeof setStatus === 'function')
      setStatus('Chưa có dữ liệu. Vui lòng tải file đầy đủ trước.', 'warn');
    return;
  }

  var runBtn = document.getElementById('tkRunBtn');
  if(runBtn) runBtn.disabled = true;
  if(typeof setStatus === 'function') setStatus('Đang thống kê...', 'warn');

  setTimeout(function(){
    try {
      applyTKFilters();

      if(tkAllClasses.length === 0){
        if(typeof setStatus === 'function'){
          var modeLabel = (tkMode === 'bld') ? 'BLD' : 'BGD';
          setStatus('Không có dữ liệu ' + modeLabel + ' theo điều kiện lọc. '
            + 'Thử đổi chế độ BLD/BGD hoặc kiểm tra bộ lọc.', 'warn');
        }
        return;
      }

      if(getTKFilterMode() === 'class' && tkSelected.size === 0){
        if(typeof setStatus === 'function')
          setStatus('Vui lòng chọn ít nhất 1 lớp trước khi thống kê.', 'warn');
        return;
      }

      var s = statsFromSelection();
      if(typeof setStatus === 'function'){
        var modeLabel = (tkMode === 'bld') ? 'BLD' : 'BGD';
        setStatus('Thống kê ' + modeLabel + ' xong: '
          + tkSelected.size + ' lớp, ' + s.total + ' sinh viên', 'ok');
      }
    } catch(err) {
      if(typeof setStatus === 'function')
        setStatus('Lỗi thống kê: ' + (err && err.message ? err.message : String(err)), 'warn');
    } finally {
      if(runBtn) runBtn.disabled = false;
    }
  }, 0);
}

/* ─────────────────────────────────────────────────────────────
   Lấy label cột chính (BLD = Mã môn, BGD = Lớp)
───────────────────────────────────────────────────────────── */
function getTKPrimaryLabel(){
  return 'Lớp';
}

function getTKPrimaryPlural(){
  return 'lớp';
}

/* ─────────────────────────────────────────────────────────────
   Lấy lý do cấm thi từ 4 học phần BLD
───────────────────────────────────────────────────────────── */
function getBGDBanReason(s){
  var reasons = [];
  if(String(s.hp1r || '').trim().toUpperCase() === 'CT') reasons.push('190081');
  if(String(s.hp2r || '').trim().toUpperCase() === 'CT') reasons.push('190082');
  if(String(s.hp3r || '').trim().toUpperCase() === 'CT') reasons.push('190083');
  if(String(s.hp4r || '').trim().toUpperCase() === 'CT') reasons.push('190084');
  return reasons.join(', ');
}

function getBLDFailedSubjectsText(s, filterWanted){
  var map = [
    {code:'190081', hp:'hp1'},
    {code:'190082', hp:'hp2'},
    {code:'190083', hp:'hp3'},
    {code:'190084', hp:'hp4'}
  ];
  var failed = [];
  map.forEach(function(item){
    var rawCode = normalizeTKResultCode(s[item.hp + 'r']);
    
    var subCate = '';
    if(rawCode === 'CT') {
      subCate = 'ct';
    } else if(rawCode === 'DC') {
      subCate = 'dc';
    } else if(isHpEmptyOrZero(s[item.hp], s[item.hp + 'r'])) {
      subCate = 'chua_hoc';
    } else if(rawCode === 'VT') {
      subCate = 'hong';
    } else {
      var score = Number(s[item.hp]);
      if(Number.isFinite(score) && score < 5) {
        subCate = 'hong';
      }
    }
    
    if(filterWanted && subCate !== filterWanted) {
      return;
    }
    
    if(rawCode === 'CT' || rawCode === 'VT' || rawCode === 'DC'){
      failed.push(item.code + '(' + rawCode + ')');
    } else {
      var score = Number(s[item.hp]);
      if(Number.isFinite(score) && score < 5){
        failed.push(item.code + '(' + score + ')');
      } else if (rawCode === '0' || s[item.hp] === 0) {
        failed.push(item.code + '(0)');
      } else if (s[item.hp] === null || s[item.hp] === undefined) {
        failed.push(item.code + '(Chưa điểm)');
      }
    }
  });
  return failed.join(', ');
}

function getBGDFailedSubjectsText(s){
  var maMH = String(s.maMH || '').trim();
  var code = normalizeTKResultCode(s.raw);
  if(code === 'CT' || code === 'VT' || code === 'DC'){
    return maMH ? (maMH + '(' + code + ')') : code;
  }
  var score = Number(s.score);
  if(Number.isFinite(score) && score < 5){
    return maMH ? (maMH + '(' + score + ')') : String(score);
  }
  return maMH;
}

function getStudentFailedSubjectsText(s, filterWanted){
  return (tkMode === 'bld') ? getBLDFailedSubjectsText(s, filterWanted) : getBGDFailedSubjectsText(s);
}

/* ─────────────────────────────────────────────────────────────
   Thu thập sinh viên cấm thi
───────────────────────────────────────────────────────────── */
function collectBannedStudents(){
  var rows = [];
  [...tkSelected].sort().forEach(function(key){
    var group = tkByLop[key];
    if(!group || !Array.isArray(group.students)) return;
    group.students.forEach(function(s){
      var cate = tkMode === 'bld' ? categoryBLD(s) : categoryBGD(s);
      if(cate !== 'ct') return;
      var maSV = s.maSV || '';
      var hoLot = s.hoLot || '';
      var ten = s.ten || '';
      var maLop = s.maLop || (tkMode === 'bgd' ? key : '');
      var maMH = tkMode === 'bld' ? (s.maMH || key) : getBGDBanReason(s);
      var ghiChu = tkMode === 'bld' ? (s.raw || 'CT') : 'CT ở học phần nêu trên';
      rows.push([key, maLop, maSV, hoLot, ten, maMH, ghiChu]);
    });
  });
  return rows;
}

/* ─────────────────────────────────────────────────────────────
   Thu thập sinh viên theo mã kết qủa (ct, vt, dc, hong)
───────────────────────────────────────────────────────────── */
function collectStudentsByResultCode(targetCode){
  var rows = [];
  var wanted = String(targetCode || '').trim().toLowerCase();
  if(!wanted) return rows;

  if(tkMode === 'bgd'){
    [...tkSelected].sort().forEach(function(key){
      var group = tkByLop[key];
      if(!group || !Array.isArray(group.students)) return;
      group.students.forEach(function(s){
        if(categoryBGD(s) !== wanted) return;
        var maSV = s.maSV || '';
        if(!maSV) return;
        var maLop = s.maLop || key || '';
        var hoLot = s.hoLot || '';
        var ten = s.ten || '';
        var maMH = getStudentFailedSubjectsText(s, wanted);
        var ghiChu = String(s.raw || '').trim() || wanted.toUpperCase();
        rows.push([key, maLop, maSV, hoLot, ten, maMH, ghiChu]);
      });
    });
  } else {
    var map = [
      {code:'190081', hp:'hp1'},
      {code:'190082', hp:'hp2'},
      {code:'190083', hp:'hp3'},
      {code:'190084', hp:'hp4'}
    ];
    [...tkSelected].sort().forEach(function(key){
      var group = tkByLop[key];
      if(!group || !Array.isArray(group.students)) return;
      group.students.forEach(function(s){
        map.forEach(function(item){
          var rawCode = normalizeTKResultCode(s[item.hp + 'r']);
          var subCate = '';
          if(rawCode === 'CT') {
            subCate = 'ct';
          } else if(rawCode === 'DC') {
            subCate = 'dc';
          } else if(isHpEmptyOrZero(s[item.hp], s[item.hp + 'r'])) {
            subCate = 'chua_hoc';
          } else if(rawCode === 'VT') {
            subCate = 'hong';
          } else {
            var score = Number(s[item.hp]);
            if(Number.isFinite(score) && score < 5) {
              subCate = 'hong';
            }
          }

          if(subCate !== wanted) return;

          var maSV = s.maSV || '';
          var hoLot = s.hoLot || '';
          var ten = s.ten || '';
          var maLop = s.maLop || '';
          var maMH = item.code + '(' + (rawCode || (s[item.hp] !== null ? s[item.hp] : 'Chưa điểm')) + ')';
          var rawText = String(s[item.hp + 'r'] || '').trim();
          var ghiChu = rawText || wanted.toUpperCase();
          rows.push([key, maLop, maSV, hoLot, ten, maMH, ghiChu]);
        });
      });
    });
  }

  return rows;
}

/* ─────────────────────────────────────────────────────────────
   Thêm sheet danh sách kết qủa vào workbook
───────────────────────────────────────────────────────────── */
function appendResultListSheet(wb, rows, sheetName, title){
  var primaryLabel = getTKPrimaryLabel();
  var aoa = [
    [title],
    ['Chế độ', tkMode === 'bld' ? 'BGD' : 'BLD'],
    ['Tổng số', rows.length],
    [],
    [primaryLabel, 'Mã lớp', 'MSSV', 'Họ lót', 'Tên', 'Học phần chưa đạt', 'Ghi chú'],
  ];
  
  rows.forEach(function(row){
    aoa.push(row);
  });
  
  var ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{wch:18},{wch:18},{wch:16},{wch:24},{wch:16},{wch:22},{wch:30}];
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

/* ─────────────────────────────────────────────────────────────
   Xuất danh sách CT/VT/DC
───────────────────────────────────────────────────────────── */
function exportTKBanList(){
  if(typeof XLSX === 'undefined'){
    if(typeof setStatus === 'function') setStatus('Thư viện XLSX chưa tải.', 'warn');
    return;
  }
  
  var selectedKeys = [...tkSelected];
  if(!selectedKeys.length){
    if(typeof setStatus === 'function') setStatus('Chưa chọn lớp/mã môn để xuất danh sách CT/VT/DC/Hỏng', 'warn');
    return;
  }

  var bannedRows = collectStudentsByResultCode('ct');
  var absentRows = collectStudentsByResultCode('vt'); // will be empty
  var suspendedRows = collectStudentsByResultCode('dc');
  var failedRows = collectStudentsByResultCode('hong');
  var chuaHocRows = collectStudentsByResultCode('chua_hoc');
  var totalRows = bannedRows.length + absentRows.length + suspendedRows.length + failedRows.length + chuaHocRows.length;

  if(!totalRows){
    if(typeof setStatus === 'function') setStatus('Không có sinh viên CT/VT/DC/Hỏng/Chưa học trong lựa chọn hiện tại', 'warn');
    return;
  }

  var wb = XLSX.utils.book_new();
  var primaryLabel = getTKPrimaryLabel();
  var modeLabel = tkMode === 'bld' ? 'BGD' : 'BLD';
  var selectedClassText = selectedKeys.length === 1
    ? selectedKeys[0]
    : selectedKeys.join(', ');
  var mergedRows = [];

  bannedRows.forEach(function(row){ mergedRows.push(row.concat(['CT'])); });
  absentRows.forEach(function(row){ mergedRows.push(row.concat(['VT'])); });
  suspendedRows.forEach(function(row){ mergedRows.push(row.concat(['DC'])); });
  failedRows.forEach(function(row){
    var rawStatus = String(row[6] || '').trim().toUpperCase();
    var label = (rawStatus === 'VT') ? 'VT' : 'Hỏng';
    mergedRows.push(row.concat([label]));
  });
  chuaHocRows.forEach(function(row){ mergedRows.push(row.concat(['Chưa học'])); });
  
  var aoa = [
    ['DANH SÁCH TỔNG HỢP CT/ĐC/HỎNG/CHƯA HỌC'],
    ['Chế độ', modeLabel],
    ['Số ' + (tkMode === 'bld' ? 'mã môn' : 'lớp') + ' đã chọn', selectedKeys.length],
    ['Mã lớp đã chọn', selectedClassText],
    ['Tổng sinh viên', totalRows],
    ['Cấm thi', bannedRows.length],
    ['Đình chỉ', suspendedRows.length],
    ['Hỏng (Điểm <5 & VT)', failedRows.length],
    ['Chưa học', chuaHocRows.length],
    [],
    [primaryLabel, 'Mã lớp', 'MSSV', 'Họ lót', 'Tên', 'Học phần chưa đạt', 'Ghi chú', 'Nhóm']
  ];

  mergedRows.forEach(function(row){
    aoa.push(row);
  });

  var ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{wch:18},{wch:18},{wch:16},{wch:24},{wch:16},{wch:30},{wch:20},{wch:10}];
  XLSX.utils.book_append_sheet(wb, ws, 'TongHop');

  if(tkMode === 'bld'){
    appendResultListSheet(wb, bannedRows, 'DS_CamThi', 'DANH SÁCH CẤM THI');
    appendResultListSheet(wb, suspendedRows, 'DS_DinhChi', 'DANH SÁCH ĐÌNH CHỈ THI');
    appendResultListSheet(wb, chuaHocRows, 'DS_ChuaHoc', 'DANH SÁCH CHƯA HỌC');
    appendResultListSheet(
      wb,
      (function(){
        var vtHongMap = {};

        function addVtHongRow(row, statusText){
          var mapKey = [row[0], row[1], row[2]].join('|');
          if(!vtHongMap[mapKey]){
            vtHongMap[mapKey] = {
              primary: row[0] || '',
              maLop: row[1] || '',
              maSV: row[2] || '',
              hoLot: row[3] || '',
              ten: row[4] || '',
              subjects: [],
              statuses: []
            };
          }
          var item = vtHongMap[mapKey];
          var subjectText = String(row[5] || '').trim();
          if(subjectText && item.subjects.indexOf(subjectText) < 0) item.subjects.push(subjectText);
          if(statusText && item.statuses.indexOf(statusText) < 0) item.statuses.push(statusText);
        }

        absentRows.forEach(function(row){ addVtHongRow(row, 'VT'); });
        failedRows.forEach(function(row){
          var rawStatus = String(row[6] || '').trim().toUpperCase();
          var statusText = (rawStatus === 'VT') ? 'VT' : 'Hỏng';
          addVtHongRow(row, statusText);
        });

        return Object.keys(vtHongMap).map(function(key){
          var item = vtHongMap[key];
          return [
            item.primary,
            item.maLop,
            item.maSV,
            item.hoLot,
            item.ten,
            item.subjects.join(', '),
            item.statuses.join(' / ')
          ];
        });
      })(),
      'DS_VT_Hong',
      'DANH SÁCH VẮNG THI + HỎNG'
    );
  } else {
    appendResultListSheet(wb, bannedRows, 'DS_CamThi', 'DANH SÁCH CẤM THI');
    appendResultListSheet(wb, suspendedRows, 'DS_DinhChi', 'DANH SÁCH ĐÌNH CHỈ THI');
    appendResultListSheet(wb, chuaHocRows, 'DS_ChuaHoc', 'DANH SÁCH CHƯA HỌC');
  }

  var d = new Date();
  var fn = 'DanhSach_CT_VT_DC_' +
    (tkMode === 'bld' ? 'BGD_' : 'BLD_') +
    d.getFullYear() +
    String(d.getMonth()+1).padStart(2,'0') +
    String(d.getDate()).padStart(2,'0') +
    '.xlsx';

  XLSX.writeFile(wb, fn, {bookSST:false, type:'binary'});
  if(typeof setStatus === 'function') setStatus('Đã xuất danh sách BGD với VT + Hỏng chung 1 sheet theo lựa chọn hiện tại', 'ok');
}

/* ─────────────────────────────────────────────────────────────
   Xuất Excel
───────────────────────────────────────────────────────────── */
function exportTK(){
  if(typeof XLSX === 'undefined'){
    if(typeof setStatus === 'function') setStatus('Thư viện XLSX chưa tải.', 'warn');
    return;
  }
  var s = statsFromSelection();
  if(!s.total){
    if(typeof setStatus === 'function') setStatus('Chưa có dữ liệu thống kê để xuất.', 'warn');
    return;
  }

  var pass = s.sx+s.g+s.kha+s.tb;
  var fail = s.ct+s.vt+s.dc+s.hong+s.chua_hoc;
  var modeLabel = (tkMode === 'bld') ? 'BGD' : 'BLD';

  /* Lấy thông tin lọc để ghi vào header */
  var schoolEl   = document.getElementById('tkSchoolSelect');
  var yearEl     = document.getElementById('tkYearSelect');
  var semEl      = document.getElementById('tkSemesterSelect');
  var filterMode = getTKFilterMode();
  var headerInfo = [
    (schoolEl && schoolEl.value)  ? ('Trường: ' + schoolEl.value)   : 'Tất cả trường',
    (yearEl   && yearEl.value)    ? ('Năm học: ' + yearEl.value)     : 'Tất cả năm học',
    (filterMode === 'semester' && semEl && semEl.value)
      ? ('Học kỳ: ' + semEl.value)
      : ('Số lớp: ' + tkSelected.size)
  ].join(' | ');

  var wb = XLSX.utils.book_new();

  /* Sheet tổng hợp */
  var aoa = [
    ['THỐNG KÊ KẾT QUẢ GDQP – ' + modeLabel],
    [headerInfo],
    [],
    ['Tổng sinh viên', s.total],
    ['Tỷ lệ đạt',     pct(pass,s.total)+'%'],
    ['Tỷ lệ không đạt', pct(fail,s.total)+'%'],
    [],
    ['ĐẠT'],
    ['Xuất sắc (9–10)', s.sx,   'TL', pct(s.sx,s.total)+'%'],
    ['Giỏi (8–8.9)',    s.g,    'TL', pct(s.g,s.total)+'%'],
    ['Khá (7–7.9)',     s.kha,  'TL', pct(s.kha,s.total)+'%'],
    ['Trung bình (5–6.9)', s.tb,'TL', pct(s.tb,s.total)+'%'],
    [],
    ['KHÔNG ĐẠT'],
    ['Cấm thi',  s.ct,   'TL', pct(s.ct,s.total)+'%'],
    ['Vắng thi', s.vt,   'TL', pct(s.vt,s.total)+'%'],
    ['Đình chỉ', s.dc,   'TL', pct(s.dc,s.total)+'%'],
    ['Hỏng',     s.hong, 'TL', pct(s.hong,s.total)+'%'],
    ['Chưa học', s.chua_hoc, 'TL', pct(s.chua_hoc,s.total)+'%'],
    [],
    ['Lớp','Tổng','SX','G','Khá','TB','CT','VT','ĐC','Hỏng','Chưa học','TL đạt (%)']
  ];

  [...tkSelected].sort().forEach(function(k){
    var c = classStats(k);
    var p = c.sx+c.g+c.kha+c.tb;
    aoa.push([k, c.total, c.sx, c.g, c.kha, c.tb, c.ct, c.vt, c.dc, c.hong, c.chua_hoc, pct(p,c.total)+'%']);
  });

  var ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{wch:30},{wch:10},{wch:10},{wch:6},{wch:10},{wch:10},{wch:10},{wch:10},{wch:10},{wch:10},{wch:10},{wch:14}];
  XLSX.utils.book_append_sheet(wb, ws, 'ThongKe_' + modeLabel);

  /* Sheet chi tiết từng lớp */
  [...tkSelected].sort().forEach(function(k){
    var c = classStats(k);
    var pClass = c.sx+c.g+c.kha+c.tb;
    var aoaC = [
      ['THỐNG KÊ LỚP – ' + k],
      ['Chế độ', modeLabel],
      ['Tổng SV', c.total],
      ['Đạt', pClass, pct(pClass,c.total)+'%'],
      ['Không đạt', c.total-pClass, pct(c.total-pClass,c.total)+'%'],
      [],
      ['ĐẠT'],
      ['Xuất sắc', c.sx,   pct(c.sx,c.total)+'%'],
      ['Giỏi',     c.g,    pct(c.g,c.total)+'%'],
      ['Khá',      c.kha,  pct(c.kha,c.total)+'%'],
      ['Trung bình',c.tb,  pct(c.tb,c.total)+'%'],
      [],
      ['KHÔNG ĐẠT'],
      ['Cấm thi',  c.ct,   pct(c.ct,c.total)+'%'],
      ['Vắng thi', c.vt,   pct(c.vt,c.total)+'%'],
      ['Đình chỉ', c.dc,   pct(c.dc,c.total)+'%'],
      ['Hỏng',     c.hong, pct(c.hong,c.total)+'%'],
      ['Chưa học',  c.chua_hoc, pct(c.chua_hoc,c.total)+'%']
    ];
    var wsC = XLSX.utils.aoa_to_sheet(aoaC);
    wsC['!cols'] = [{wch:24},{wch:10},{wch:12}];
    var sheetName = ('Lop_'+k).replace(/[\\\/?*[\]:]/g,'_').substring(0,31);
    XLSX.utils.book_append_sheet(wb, wsC, sheetName);
  });

  /* Sheet dữ liệu gốc theo lớp đã chọn (giống dạng file đầy đủ) */
  if(Array.isArray(tkRawRows) && tkRawRows.length && tkCols && tkCols.maLop){
    var selectedSet = new Set([...tkSelected]);
    var selectedRows = tkRawRows.filter(function(r){
      if(!r) return false;
      var maLop = r[tkCols.maLop] != null ? String(r[tkCols.maLop]).trim() : '';
      return selectedSet.has(maLop);
    });

    if(selectedRows.length){
      var header = Object.keys(selectedRows[0]);
      var aoaRaw = [header];
      selectedRows.forEach(function(r){
        aoaRaw.push(header.map(function(h){
          return r[h] == null ? '' : r[h];
        }));
      });
      var wsRaw = XLSX.utils.aoa_to_sheet(aoaRaw);
      wsRaw['!cols'] = header.map(function(){ return {wch:16}; });
      XLSX.utils.book_append_sheet(wb, wsRaw, 'DuLieuDaChon');
    }
  }

  var schoolElForFile = document.getElementById('tkSchoolSelect');
  var selectedSchool = (schoolElForFile && schoolElForFile.value)
    ? schoolElForFile.value
    : 'Tất cả trường';
  var schoolSuffix = sanitizeFileNamePart(selectedSchool || 'Tất cả trường');

  var fn;
  if(tkMode === 'bld'){
    var bldSchool = sanitizeFileNamePart(selectedSchool || 'Tất cả trường');
    fn = 'thongke BLD-GDQP- ' + bldSchool + ' - ' + schoolSuffix + '.xlsx';
  } else {
    var selectedClasses = [...tkSelected].sort();
    var classPart = selectedClasses.length === 1
      ? ('lớp ' + selectedClasses[0])
      : 'lớp tổng hợp';
    fn = 'thongke BGD-GDQP- Đại học Trà Vinh - '
      + sanitizeFileNamePart(classPart)
      + ' - '
      + schoolSuffix
      + '.xlsx';
  }

  XLSX.writeFile(wb, fn, {bookSST:false, type:'binary'});
  if(typeof setStatus === 'function') setStatus('Đã xuất file thống kê ' + modeLabel, 'ok');
}

/* ─────────────────────────────────────────────────────────────
   Lưu / khôi phục trạng thái giao diện
───────────────────────────────────────────────────────────── */
function saveTKState(){
  try {
    var schoolEl   = document.getElementById('tkSchoolSelect');
    var yearEl     = document.getElementById('tkYearSelect');
    var ftEl       = document.getElementById('tkFilterType');
    var semEl      = document.getElementById('tkSemesterSelect');
    var modeSelEl  = document.getElementById('tkModeSelect');
    var state = {
      mode:       tkMode,
      school:     schoolEl  ? schoolEl.value  : '',
      year:       yearEl    ? yearEl.value    : '',
      filterType: ftEl      ? ftEl.value      : 'semester',
      semester:   semEl     ? semEl.value     : '',
      selected:   [...tkSelected]
    };
    sessionStorage.setItem(TK_SS_STATE, JSON.stringify(state));
  } catch(e){}
}

function hydrateTKFromSession(){
  /* 1. Khôi phục trạng thái UI */
  try {
    var raw = sessionStorage.getItem(TK_SS_STATE);
    if(raw){
      var state = JSON.parse(raw);
      if(state.mode) tkMode = state.mode;
      var schoolEl  = document.getElementById('tkSchoolSelect');
      var yearEl    = document.getElementById('tkYearSelect');
      var ftEl      = document.getElementById('tkFilterType');
      var semEl     = document.getElementById('tkSemesterSelect');
      var modeEl    = document.getElementById('tkModeSelect');
      if(schoolEl && state.school)     schoolEl.value  = state.school;
      if(yearEl   && state.year)       yearEl.value    = state.year;
      if(ftEl     && state.filterType) ftEl.value      = state.filterType;
      if(semEl    && state.semester)   semEl.value     = state.semester;
      if(modeEl)                       modeEl.value    = tkMode;
      if(Array.isArray(state.selected)) tkSelected = new Set(state.selected);
    }
  } catch(e){}

  updateTKModeLabels();

  /* 2. Khôi phục dữ liệu từ session nếu chưa có */
  if(!Array.isArray(tkRawRows) || !tkRawRows.length){
    var rawRows = sessionStorage.getItem('gdqp_full_rows')
               || sessionStorage.getItem('gdqp_main_rows');
    if(rawRows){
      try {
        var rows = JSON.parse(rawRows);
        if(Array.isArray(rows) && rows.length){
          tkRawRows = rows;
          if(typeof detectCols === 'function')            tkCols = detectCols(rows[0]);
          if(typeof collectTKFilterOptions === 'function') collectTKFilterOptions();
          if(typeof renderSchoolOptions === 'function')    renderSchoolOptions();
          if(typeof renderYearOptions === 'function')      renderYearOptions();
        }
      } catch(e){}
    }
  }

  /* 3. Áp dụng lọc nếu có dữ liệu */
  if(Array.isArray(tkRawRows) && tkRawRows.length && tkCols){
    applyTKFilters();
  } else {
    renderTKClassList();
    updateTKSummary();
    updateTKMetaText();
  }
}

/* ─────────────────────────────────────────────────────────────
   Khởi tạo
───────────────────────────────────────────────────────────── */
ensureTKGlobals();
updateTKModeLabels();
if(typeof hydrateTKFromSession === 'function') hydrateTKFromSession();