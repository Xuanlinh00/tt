/* STATE */
let byLop = {}, selLop = new Set(), allKeys = [];
let activeBD = null, activeQD = null;
let outerMode = 'bd';

const HP_MAP = {'190081':1,'190082':2,'190083':3,'190084':4};
const SS_KEY_DATA = 'gdqp_byLop';
const SS_KEY_FILE = 'gdqp_fileName';
const SS_KEY_BIRTH_FILE = 'gdqp_birth_fileName';
let birthFileName = '';

const EXPORT_TEMPLATE = {
  school: 'ĐẠI HỌC TRÀ VINH',
  centerTop: 'TRUNG TÂM GIÁO DỤC QUỐC PHÒNG',
  centerBottom: 'VÀ AN NINH ĐẠI HỌC TRÀ VINH',
  monhoc: 'Giáo dục quốc phòng và an ninh',
  monhocDetail: '(Trình độ: ĐH, CĐSP; Thời lượng: 165 tiết)',
  soQD: '67',
  ngayQD: '21 tháng 02 năm 2025',
  nguoiKy: 'Trương Minh Hải',
  chucVu: 'KT. GIÁM ĐỐC - PHÓ GIÁM ĐỐC'
};

function buildExportDate(){
  return 'Vĩnh Long, ngày      tháng      năm';
}
function getExportSettings(){
  return {
    ...EXPORT_TEMPLATE,
    xdate: buildExportDate()
  };
}

/* UPLOAD */
const zone=document.getElementById('zone'), fi=document.getElementById('fi'), fiNoiSinh=document.getElementById('fiNoiSinh');
zone.addEventListener('dragover',e=>{e.preventDefault();zone.classList.add('over')});
zone.addEventListener('dragleave',()=>zone.classList.remove('over'));
zone.addEventListener('drop',e=>{e.preventDefault();zone.classList.remove('over');if(e.dataTransfer.files[0])loadFile(e.dataTransfer.files[0])});
zone.addEventListener('click',e=>{if(e.target.tagName!=='SPAN')fi.click()});
fi.addEventListener('change',e=>{if(e.target.files[0])loadFile(e.target.files[0])});
fiNoiSinh.addEventListener('change',e=>{if(e.target.files[0])loadNoiSinhFile(e.target.files[0])});

function showMsg(t,tp){const el=document.getElementById('msg');el.style.display='block';el.className='st '+tp;el.innerHTML=t}

function shareFileToParent(type, file){
  if(!file || !window.parent || window.parent===window) return;
  window.parent.postMessage({ type, file, fileName: file.name || '' }, '*');
}

function loadFile(file, fromShared=false){
  showMsg(` Đang đọc <b>${file.name}</b>...`,'info');
  if(!fromShared) shareFileToParent('gdqp-share-main', file);
  const r=new FileReader();
  r.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'binary'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:null});
      if(!rows.length){showMsg(' File trống!','err');return}

      if(isNoiSinhOnlyFile(rows)){
        if(!Object.keys(byLop).length){
          showMsg(' Đây là file Nơi sinh. Vui lòng tải file điểm trước, sau đó mới ghép file này.','err');
          return;
        }
        const merged = mergeNoiSinh(rows);
        if(merged===0){
          showMsg(' Không ghép được Nơi sinh nào. Kiểm tra cột MaSV/MSSV và Nơi sinh trong file bổ sung.','err');
          return;
        }
        birthFileName=file.name;
        sessionStorage.setItem(SS_KEY_BIRTH_FILE, birthFileName);
        try{ sessionStorage.setItem(SS_KEY_DATA, JSON.stringify(byLop)); }catch(e){}
        renderTabs();
        renderQDTabs();
        refreshLoadedBanner();
        showMsg(` Đã ghép Nơi sinh cho <b>${merged}</b> sinh viên từ file <b>${file.name}</b>.`,'ok');
        return;
      }

      buildData(rows, file.name);
      showMsg(` Đọc xong <b>${rows.length.toLocaleString()}</b> dòng  tìm thấy <b>${allKeys.length}</b> lớp`,'ok');
    }catch(err){showMsg(' Lỗi: '+err.message,'err')}
  };
  r.readAsBinaryString(file);
}

function isNoiSinhOnlyFile(rows){
  if(!rows.length) return false;
  const col=detectCols(rows[0]);
  const hasMaSV = !!col.maSV;
  const hasNoiSinh = !!col.noiSinh;
  const hasScore = !!col.score;
  return hasMaSV && hasNoiSinh && !hasScore;
}

function triggerNoiSinhUpload(){
  if(!Object.keys(byLop).length){
    alert('Vui lòng tải file điểm trước khi thêm file Nơi sinh.');
    return;
  }
  fiNoiSinh.click();
}

function setExternalUploadMode(enabled){
  const zoneEl = document.getElementById('zone');
  const actionsEl = document.getElementById('fileBannerActions');
  if(zoneEl) zoneEl.style.display = enabled ? 'none' : '';
  if(actionsEl) actionsEl.style.display = enabled ? 'none' : 'flex';
}

function loadNoiSinhFile(file, fromShared=false){
  showMsg(` Đang đọc file Nơi sinh <b>${file.name}</b>...`,'info');
  if(!fromShared) shareFileToParent('gdqp-share-birth', file);
  const r=new FileReader();
  r.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'binary'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:null});
      if(!rows.length){showMsg(' File Nơi sinh trống!','err');return}
      if(!isNoiSinhOnlyFile(rows)){
        showMsg(' File này chưa đúng dạng bổ sung Nơi sinh (cần có MaSV/MSSV và Nơi sinh, không phải file điểm).','err');
        return;
      }
      const merged = mergeNoiSinh(rows);
      if(merged===0){
        showMsg(' Không ghép được Nơi sinh nào. Kiểm tra cột MaSV/MSSV và Nơi sinh trong file bổ sung.','err');
        return;
      }
      birthFileName=file.name;
      sessionStorage.setItem(SS_KEY_BIRTH_FILE, birthFileName);
      try{ sessionStorage.setItem(SS_KEY_DATA, JSON.stringify(byLop)); }catch(e){}
      renderTabs();
      renderQDTabs();
      refreshLoadedBanner();
      showMsg(` Đã ghép Nơi sinh cho <b>${merged}</b> sinh viên từ file <b>${file.name}</b>.`,'ok');
    }catch(err){showMsg(' Lỗi file Nơi sinh: '+err.message,'err')}
  };
  r.readAsBinaryString(file);
}

function detectCols(sample){
  const keys=Object.keys(sample);
  const nk = s => String(s||'')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .toLowerCase()
    .replace(/đ/g,'d')
    .replace(/\s+/g,'');
  const find=(...ns)=>keys.find(k=>ns.some(n=>nk(k)===nk(n)))||null;
  return{
    maSV:    find('MaSV','MSSV','masv'),
    hoLot:   find('HoLotSV','HoLot','holot'),
    ten:     find('TenSV','Ten'),
    ngay:    find('NgaySinhC','NgaySinh'),
    noiSinh: keys.find(k=>nk(k).includes('noisinh'))||null,
    maLop:   find('MaLop','malop'),
    maMH:    find('MaMH','mamh'),
    score:   keys.find(k=>{const v=nk(k);return v==='t1'||v.includes('dtkl1')||v.includes('diemthi1')||v.includes('diemtongketlan1');})||null,
  };
}

function buildData(rows, fileName){
  const col=detectCols(rows[0]);
  byLop={};
  birthFileName='';
  sessionStorage.removeItem(SS_KEY_BIRTH_FILE);
  rows.forEach(r=>{
    const maMH=r[col.maMH]!=null?String(r[col.maMH]).trim():null;
    const hp=maMH?HP_MAP[maMH]:null;
    if(!hp)return;
    const maSV=r[col.maSV]!=null?String(r[col.maSV]).trim():'';
    const maLop=r[col.maLop]!=null?String(r[col.maLop]).trim():'Không rõ';
    if(!byLop[maLop])byLop[maLop]={maLop,svMap:{}};
    if(!byLop[maLop].svMap[maSV])byLop[maLop].svMap[maSV]={
      maSV,maLop,
      hoLot:   r[col.hoLot]   ||'',
      ten:     r[col.ten]     ||'',
      ngay:    r[col.ngay]    ||'',
      noiSinh: col.noiSinh&&r[col.noiSinh]?String(r[col.noiSinh]):'',
      hp1:null,hp2:null,hp3:null,hp4:null,
      hp1r:'',hp2r:'',hp3r:'',hp4r:''
    };
    const sv=byLop[maLop].svMap[maSV];
    const noiSinhNow = col.noiSinh&&r[col.noiSinh]!=null ? String(r[col.noiSinh]).trim() : '';
    if(noiSinhNow && !String(sv.noiSinh||'').trim()) sv.noiSinh = noiSinhNow;
    const rawVal=r[col.score];
    const rawStr=(rawVal!==null&&rawVal!==undefined&&String(rawVal).trim()!=='')
      ? String(rawVal).trim() : '';
    const parsed = rawStr!==''?parseFloat(rawStr):NaN;
    const numVal = (!isNaN(parsed) && rawStr!=='')?parsed:null;
    const oldRaw = sv['hp'+hp+'r'] || '';
    if(rawStr!=='' || oldRaw===''){
      sv['hp'+hp]    = numVal;
      sv['hp'+hp+'r']= rawStr;
    }
  });
  Object.values(byLop).forEach(g=>{
    g.students=Object.values(g.svMap).sort((a,b)=>a.maSV.localeCompare(b.maSV));
    delete g.svMap;
  });
  allKeys=Object.keys(byLop).sort();
  if(!allKeys.length){showMsg(' Không tìm thấy dữ liệu MaMH 190081-190084!','err');return}

  try{ sessionStorage.setItem(SS_KEY_DATA, JSON.stringify(byLop)); }catch(e){}
  if(fileName) sessionStorage.setItem(SS_KEY_FILE, fileName);

  selLop=new Set();
  activeBD=null; activeQD=null;

  showUI(fileName||sessionStorage.getItem(SS_KEY_FILE)||'file.xlsx');
}

function restoreFromSession(){
  const raw=sessionStorage.getItem(SS_KEY_DATA);
  if(!raw) return false;
  try{
    byLop=JSON.parse(raw);
    birthFileName=sessionStorage.getItem(SS_KEY_BIRTH_FILE)||'';
    allKeys=Object.keys(byLop).sort();
    if(!allKeys.length) return false;
    selLop=new Set();
    activeBD=null; activeQD=null;
    const fname=sessionStorage.getItem(SS_KEY_FILE)||'file.xlsx';
    showUI(fname);
    return true;
  }catch(e){ return false; }
}

function showUI(fileName){
  const svTotal = Object.values(byLop).reduce((s,g)=>s+g.students.length,0);
  document.getElementById('zone').style.display='none';
  document.getElementById('fileBanner').style.display='flex';
  document.getElementById('loadedFileName').textContent=fileName;
  document.getElementById('loadedInfo').textContent=`${allKeys.length} lớp  ${svTotal.toLocaleString()} sinh viên`;
  document.getElementById('loadedBirthFile').textContent=birthFileName||'Chưa có';
  document.getElementById('mainPanel').style.display='block';
  initSearch();
  renderList();
  renderTags();
  renderTabs();
  renderQDTabs();
}

function refreshLoadedBanner(){
  const svTotal = Object.values(byLop).reduce((s,g)=>s+g.students.length,0);
  document.getElementById('loadedInfo').textContent=`${allKeys.length} lớp  ${svTotal.toLocaleString()} sinh viên`;
  document.getElementById('loadedBirthFile').textContent=birthFileName||'Chưa có';
}

function mergeNoiSinh(rows){
  if(!rows.length) return 0;
  const col=detectCols(rows[0]);
  if(!col.maSV || !col.noiSinh) return 0;

  const byClassAndSV = new Map();
  const bySV = new Map();
  Object.values(byLop).forEach(g=>{
    g.students.forEach(sv=>{
      const key = `${sv.maLop}|${sv.maSV}`;
      byClassAndSV.set(key, sv);
      if(!bySV.has(sv.maSV)) bySV.set(sv.maSV, []);
      bySV.get(sv.maSV).push(sv);
    });
  });

  let merged=0;
  rows.forEach(r=>{
    const maSV = r[col.maSV]!=null?String(r[col.maSV]).trim():'';
    const noiSinhRaw = col.noiSinh&&r[col.noiSinh]!=null?String(r[col.noiSinh]).trim():'';
    const maLop = col.maLop&&r[col.maLop]!=null?String(r[col.maLop]).trim():'';
    if(!maSV || !noiSinhRaw) return;

    let touched=false;
    if(maLop){
      const sv = byClassAndSV.get(`${maLop}|${maSV}`);
      if(sv){ sv.noiSinh=noiSinhRaw; touched=true; }
    }else{
      const list = bySV.get(maSV) || [];
      list.forEach(sv=>{ sv.noiSinh=noiSinhRaw; touched=true; });
    }
    if(touched) merged++;
  });
  return merged;
}

/* SEARCH */
let _searchHandler=null;
function initSearch(){
  const inp=document.getElementById('searchInput');
  inp.value='';
  if(_searchHandler) inp.removeEventListener('input',_searchHandler);
  _searchHandler=()=>renderList(inp.value);
  inp.addEventListener('input',_searchHandler);
}

function norm(s){return s.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/đ/g,'d').replace(/\s+/g,'')}

function renderList(q=''){
  const list=document.getElementById('lopList');
  const countEl=document.getElementById('listCount');
  list.innerHTML='';
  const tokens=norm(q.trim()).split(/\s+/).filter(Boolean);
  const filtered=allKeys.filter(k=>!tokens.length||tokens.every(t=>norm(k).includes(t)));

  countEl.textContent=tokens.length
    ?` "${q.trim()}": ${filtered.length} / ${allKeys.length} lớp`
    :`Tổng ${allKeys.length} lớp  nhấn để chọn / bỏ chọn`;

  if(!filtered.length){
    list.innerHTML=`<div style="padding:16px;text-align:center;color:var(--muted);font-style:italic;font-size:12px">Không tìm thấy lớp nào khớp với "<b>${q}</b>"<br><span style="font-size:11px">Xoá ô tìm kiếm để xem tất cả</span></div>`;
    return;
  }
  const frag=document.createDocumentFragment();
  filtered.forEach(k=>{
    const cnt=byLop[k].students.length;
    const sel=selLop.has(k);
    const el=document.createElement('div');
    el.className='lop-item'+(sel?' sel':'');
    el.innerHTML=`<div><b>${k}</b><span class="cnt" style="margin-left:8px">${cnt} SV</span></div><span class="chk"></span>`;
    el.addEventListener('click',()=>{
      if(selLop.has(k)){selLop.delete(k);el.classList.remove('sel')}
      else{selLop.add(k);el.classList.add('sel')}
      saveSelLop();
      renderTags();renderTabs();renderQDTabs();
    });
    frag.appendChild(el);
  });
  list.appendChild(frag);
}

function saveSelLop(){
  try{ sessionStorage.setItem('gdqp_selLop', JSON.stringify([...selLop])); }catch(e){}
}
function loadSelLop(){
  try{
    const raw=sessionStorage.getItem('gdqp_selLop');
    if(raw) selLop=new Set(JSON.parse(raw).filter(k=>byLop[k]));
  }catch(e){}
}

/* TAGS */
function renderTags(){
  const wrap=document.getElementById('selTags');
  wrap.innerHTML='';
  document.getElementById('selCount').textContent=selLop.size;
  if(!selLop.size){wrap.innerHTML='<div class="sel-none">Chưa chọn lớp nào  nhấn vào lớp bên trái để thêm</div>';return}
  [...selLop].sort().forEach(k=>{
    const tag=document.createElement('div');
    tag.className='sel-tag';
    tag.innerHTML=`${k}<span class="rm" title="Bỏ chọn" onclick="removeLop('${k}')"></span>`;
    wrap.appendChild(tag);
  });
}
function removeLop(k){
  selLop.delete(k);
  saveSelLop();
  renderList(document.getElementById('searchInput').value);
  renderTags();renderTabs();renderQDTabs();
}
function selectAll(){
  selLop=new Set(allKeys);
  saveSelLop();
  renderList(document.getElementById('searchInput').value);
  renderTags();renderTabs();renderQDTabs();
}
function clearAll(){
  selLop=new Set();
  saveSelLop();
  renderList(document.getElementById('searchInput').value);
  renderTags();renderTabs();renderQDTabs();
}

function switchOuter(mode){
  outerMode=mode;
  document.getElementById('otab-bd').classList.toggle('on', mode==='bd');
  document.getElementById('otab-qd').classList.toggle('on', mode==='qd');
  document.getElementById('panel-bd').style.display = mode==='bd'?'block':'none';
  document.getElementById('panel-qd').style.display = mode==='qd'?'block':'none';
}

function renderTabs(){
  const tabrow=document.getElementById('tabrow');tabrow.innerHTML='';
  const keys=[...selLop].sort();
  if(!keys.length){
    document.getElementById('tabbody').innerHTML='<div class="empty" style="padding:44px">Chọn lớp bên trên để xem trước bảng điểm</div>';
    return;
  }
  if(!keys.includes(activeBD)) activeBD=keys[0];
  keys.forEach(k=>{
    const btn=document.createElement('button');
    btn.className='tabbtn'+(k===activeBD?' on':'');
    btn.textContent=k; btn.title=k;
    btn.onclick=()=>{activeBD=k;renderTabs()};
    tabrow.appendChild(btn);
  });
  drawBDTable(activeBD);
}

function isCode(s, hp){ return s[hp+'r']!=='' && s[hp]===null; }
function isCTCode(s, hp){ return String(s[hp+'r']||'').trim().toUpperCase()==='CT'; }
function allHPAreCT(s){
  return isCTCode(s,'hp1') && isCTCode(s,'hp2') && isCTCode(s,'hp3') && isCTCode(s,'hp4');
}
function fmtHP(s, hp){
  const r=s[hp+'r'];
  if(r==='') return '';
  if(s[hp]!==null) return s[hp].toFixed(1);
  return r;
}

function calcTB(s){
  if(isCode(s,'hp1')||isCode(s,'hp2')||isCode(s,'hp3')||isCode(s,'hp4'))return null;
  const{hp1,hp2,hp3,hp4}=s;
  if(hp1==null||hp2==null||hp3==null||hp4==null)return null;
  return Math.round((hp1*3+hp2*2+hp3*1+hp4*2)/8*10)/10;
}

function xepLoai(s,tb){
  if(isCode(s,'hp1')||isCode(s,'hp2')||isCode(s,'hp3')||isCode(s,'hp4'))return 'Không đạt';
  if(s.hp1==null||s.hp2==null||s.hp3==null||s.hp4==null)return 'Không đạt';
  if(s.hp1<5||s.hp2<5||s.hp3<5||s.hp4<5)return 'Không đạt';
  if(tb===null||isNaN(tb))return 'Không đạt';
  if(tb>=8.95)return 'Xuất sắc';
  if(tb>=7.95)return 'Giỏi';
  if(tb>=6.95)return 'Khá';
  if(tb>=5.95)return 'Trung bình';
  return 'Không đạt';
}
function ketQua(xl){return xl==='Không đạt'?'Hỏng':'Đạt'}

function drawBDTable(key){
  if(!key||!byLop[key])return;
  const cfg = getExportSettings();
  const{maLop,students}=byLop[key];
  let dat=0,hong=0;
  const tbody=students.map((s,i)=>{
    const tb=calcTB(s); const xl=xepLoai(s,tb); const kq=ketQua(xl);
    if(kq==='Đạt')dat++;else hong++;

    const tbStr = allHPAreCT(s) ? 'CT' : (tb!==null ? tb.toFixed(1) : '');

    const hpCell=(h)=>{
      const v=fmtHP(s,h);
      const isC=(s[h]===null&&s[h+'r']!=='');
      return `<td style="${isC?'color:#b85c00;font-weight:bold':''}${s[h]!==null&&s[h]<5?'color:var(--fail);font-weight:bold':''}">${v}</td>`;
    };
    return`<tr>
      <td>${i+1}</td><td>${s.maSV}</td>
      <td class="L">${s.hoLot}</td><td>${s.ten}</td><td>${s.ngay}</td>
      ${hpCell('hp1')}${hpCell('hp2')}${hpCell('hp3')}${hpCell('hp4')}
      <td class="bold">${tbStr}</td>
      <td>${xl}</td><td class="${kq==='Đạt'?'pass':'fail'}">${kq}</td>
      <td></td><td class="L">${maLop}</td>
    </tr>`;
  }).join('');
  const rate=students.length?Math.round(dat/students.length*100):0;
  document.getElementById('tabbody').innerHTML=`
    <div class="phdr">
      <h4>DANH SÁCH GHI ĐIỂM MÔN GIÁO DỤC QUỐC PHÒNG VÀ AN NINH</h4>
      <p>Nhóm/lớp: <b>${maLop}</b> &ensp;|&ensp; Môn học: <b>${cfg.monhoc} ${cfg.monhocDetail||''}</b> &ensp;|&ensp; <b>${students.length} sinh viên</b></p>
    </div>
    <div class="tblwrap">
      <table class="tbl">
        <thead>
          <tr>
            <th rowspan="2">TT</th><th rowspan="2">MSSV</th>
            <th rowspan="2">HỌ VÀ</th><th rowspan="2">TÊN</th>
            <th rowspan="2">NGÀY SINH</th>
            <th colspan="4">ĐIỂM THI</th>
            <th rowspan="2">ĐIỂM TB</th><th rowspan="2">XẾP LOẠI</th>
            <th rowspan="2">KẾT QUẢ</th><th rowspan="2">GHI CHÚ</th><th rowspan="2">Mã lớp</th>
          </tr>
          <tr>
            <th class="sub">HP I<br>(3)</th><th class="sub">HP II<br>(2)</th>
            <th class="sub">HP III<br>(1)</th><th class="sub">HP IV<br>(2)</th>
          </tr>
        </thead>
        <tbody>${tbody}</tbody>
      </table>
    </div>
    <div class="srow">
      <div class="si"><span class="sl">Tổng:</span><span class="sv">${students.length} SV</span></div>
      <div class="si"><span class="sl">Đạt:</span><span class="sv" style="color:var(--pass)">${dat}</span></div>
      <div class="si"><span class="sl">Hỏng:</span><span class="sv" style="color:var(--fail)">${hong}</span></div>
      <div class="si"><span class="sl">Tỷ lệ đạt:</span><span class="sv">${rate}%</span></div>
    </div>`;
}

let qdChiDat = true;

function renderQDTabs(){
  const tabrow=document.getElementById('qdtabrow'); tabrow.innerHTML='';
  const keys=[...selLop].sort();
  if(!keys.length){
    document.getElementById('qdtabbody').innerHTML='<div class="empty" style="padding:44px">Chọn lớp bên trên để xem trước quyết định</div>';
    return;
  }
  if(!keys.includes(activeQD)) activeQD=keys[0];
  tabrow.innerHTML=`<div class="empty" style="padding:10px 14px">Đang gộp ${keys.length} lớp vào một danh sách quyết định chung</div>`;
  drawQDTable();
}

function getQDCombinedList(keys){
  const merged=[];
  keys.forEach(key=>{
    const group=byLop[key];
    if(!group) return;
    group.students.forEach(s=>merged.push({...s, maLop:key}));
  });
  merged.sort((a,b)=>a.maLop.localeCompare(b.maLop)||a.maSV.localeCompare(b.maSV));
  if(!qdChiDat) return merged;
  return merged.filter(s=>{const tb=calcTB(s);return ketQua(xepLoai(s,tb))==='Đạt'});
}

function drawQDTable(){
  const keys=[...selLop].sort();
  if(!keys.length)return;
  const cfg = getExportSettings();
  const nguoiKy= cfg.nguoiKy;
  const chucVu = cfg.chucVu;
  const school = cfg.school;
  const centerTop = cfg.centerTop;
  const centerBottom = cfg.centerBottom;
  const xdate  = cfg.xdate;

  const list = getQDCombinedList(keys);

  const tbody=list.map((s,i)=>{
    const tb=calcTB(s); const xl=xepLoai(s,tb);
    return`<tr>
      <td>${i+1}</td><td>${s.maSV}</td>
      <td class="L">${s.hoLot}</td><td>${s.ten}</td><td>${s.ngay}</td>
      <td class="L">${s.noiSinh||''}</td>
      <td class="bold">${allHPAreCT(s)?'CT':(tb!==null?tb.toFixed(1):'')}</td>
      <td>${xl}</td><td></td><td class="L">${s.maLop}</td>
    </tr>`;
  }).join('');

  const total = list.length;
  const chiDatLabel = qdChiDat
    ? `<span style="color:var(--pass);font-size:11px">(chỉ hiện ${total} SV đạt)</span>`
    : `<span style="color:var(--muted);font-size:11px">(tất cả ${total} SV)</span>`;

  document.getElementById('qdtabbody').innerHTML=`
    <div class="qd-preview">
      <div class="qd-filter-bar">
        <label>
          <input type="checkbox" ${qdChiDat?'checked':''} onchange="qdChiDat=this.checked;drawQDTable()">
          Chỉ hiển thị sinh viên <b>Đạt</b>
        </label>
        <span class="qd-stat">${chiDatLabel}</span>
      </div>

      <div class="qd-header-row">
        <div class="qd-header-col left">
          <div style="font-size:13px;font-weight:bold">${school.replace('TRƯỜNG ','')}</div>
          <div style="font-size:12px;font-weight:bold">${centerTop}</div>
          <div style="font-size:12px;font-weight:bold">${centerBottom}</div>
        </div>
        <div class="qd-header-col right">
          <div style="font-size:13px;font-weight:bold">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
          <div style="font-size:12px;font-weight:bold;text-decoration:underline">Độc lập - Tự do - Hạnh phúc</div>
          <hr class="qd-hairline">
        </div>
      </div>
      <div class="qd-date">${xdate}</div>

      <div class="qd-title">Danh sách cấp chứng chỉ giáo dục quốc phòng và an ninh</div>
      <div class="qd-sub">(Ban hành kèm theo Quyết định số      /QĐ-GDQP ngày      tháng      năm      )</div>

      <div class="tblwrap">
        <table class="tbl">
          <thead>
            <tr>
              <th>TT</th><th>Mã SV</th><th>Họ Và</th><th>Tên</th>
              <th>Ngày Sinh</th><th>Nơi Sinh</th><th>Điểm TB</th>
              <th>Xếp Loại</th><th>Ghi Chú</th><th>Mã Lớp</th>
            </tr>
          </thead>
          <tbody>${tbody||'<tr><td colspan="10" style="padding:20px;color:var(--muted);font-style:italic">Không có sinh viên</td></tr>'}</tbody>
        </table>
      </div>

      <div style="margin-top:12px;font-size:13px;font-style:italic">
        Trên danh sách có <b>${total}</b> sinh viên
      </div>

      <div style="display:flex;justify-content:flex-end;margin-top:24px;font-size:13px;text-align:center">
        <div>
          <div style="font-weight:bold">${chucVu}</div>
          <div style="margin-top:40px;font-weight:bold">${nguoiKy}</div>
        </div>
      </div>
    </div>`;
}

function hpVal(s, h){
  const r = s[h+'r'];
  if(r === '') return '';
  if(s[h] !== null) return s[h];
  return r;
}

function centerSheetCells(ws, startRow, endRow, cols){
  for(let r=startRow; r<=endRow; r++){
    cols.forEach(c=>{
      const addr = XLSX.utils.encode_cell({r, c});
      if(!ws[addr]) return;
      if(!ws[addr].s) ws[addr].s={};
      if(!ws[addr].s.alignment) ws[addr].s.alignment={};
      ws[addr].s.alignment.horizontal='center';
      ws[addr].s.alignment.vertical='center';
    });
  }
}

function applyCellStyle(ws, row, col, style){
  const addr = XLSX.utils.encode_cell({r:row, c:col});
  ws[addr] = ws[addr] || {};
  ws[addr].s = { ...(ws[addr].s || {}), ...style };
}

function applyRangeStyle(ws, rowStart, rowEnd, colStart, colEnd, style){
  for(let r=rowStart; r<=rowEnd; r++){
    for(let c=colStart; c<=colEnd; c++){
      applyCellStyle(ws, r, c, style);
    }
  }
}

function setRowHeights(ws, rows){
  ws['!rows'] = ws['!rows'] || [];
  rows.forEach(({row, hpt})=>{
    ws['!rows'][row] = { hpt };
  });
}

function exportExcel(scope='all'){
  let keys = scope==='current' ? (activeBD?[activeBD]:[]) : [...selLop].sort();
  if(scope==='current' && !keys.length && selLop.size) keys=[[...selLop].sort()[0]];
  if(!keys.length){alert('Chưa chọn lớp nào!');return}
  const wb=XLSX.utils.book_new();
  const cfg = getExportSettings();
  const school=cfg.school, centerTop=cfg.centerTop, centerBottom=cfg.centerBottom, monhoc=cfg.monhoc, monhocDetail=cfg.monhocDetail||'', xdate=cfg.xdate;
  const M=(r1,c1,r2,c2)=>({s:{r:r1,c:c1},e:{r:r2,c:c2}});

  keys.forEach(key=>{
    const{maLop,students}=byLop[key];
    const aoa=[];
    aoa.push([school,'','','','','CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM','','','','','','','','']);
    aoa.push([centerTop,'','','','','Độc lập - Tự do - Hạnh phúc','','','','','','','','']);
    aoa.push([centerBottom,'','','','',xdate,'','','','','','','','']);
    aoa.push([]);
    aoa.push(['DANH SÁCH GHI ĐIỂM MÔN GIÁO DỤC QUỐC PHÒNG VÀ AN NINH','','','','','','','','','','','','','']);
    aoa.push(['Trường: '+school,'','','','','','','','','','','','','']);
    aoa.push(['Nhóm/ lớp: '+maLop,'','','','','','','','','','','','','']);
    aoa.push(['Môn học: '+monhoc,'','','','','','',monhocDetail,'','','','','','']);
    aoa.push([]);
    aoa.push(['TT','MSSV','HỌ VÀ','TÊN','NGÀY SINH','ĐIỂM THI','','','','ĐIỂM TB','XẾP LOẠI','KẾT QUẢ','GHI CHÚ','Mã lớp']);
    aoa.push(['','','','','','HP I (3)','HP II (2)','HP III (1)','HP IV (2)','','','','','']);

    let dat=0, hong=0;
    students.forEach((s,i)=>{
      const tb  = calcTB(s);
      const xl  = xepLoai(s, tb);
      const kq  = ketQua(xl);
      if(kq==='Đạt') dat++; else hong++;

      const tbVal = allHPAreCT(s) ? 'CT' : (tb !== null ? tb : '');

      aoa.push([
        i+1, s.maSV, s.hoLot, s.ten, s.ngay,
        hpVal(s,'hp1'), hpVal(s,'hp2'), hpVal(s,'hp3'), hpVal(s,'hp4'),
        tbVal,
        xl,
        kq,
        '',
        maLop
      ]);
    });

    aoa.push([]);
    const ni = aoa.length;
    aoa.push(['Ghi chú: "CT" - Cấm thi.','','','','','','','','','','','','','']);
    aoa.push(['','Tổng số SV trên danh sách:','','', students.length,'sinh viên/học sinh','','','','','','','','']);
    aoa.push(['','Số sinh viên đạt:','','',  dat,            'sinh viên/học sinh','','','','','','','','']);
    aoa.push(['','Số sinh viên hỏng:','','', hong,            'sinh viên/học sinh','','','','','','','','']);
    aoa.push([]);
    aoa.push(['Cán bộ ghi điểm','','','','Phòng ĐT, QLSV, ĐTBD','','','','','KT. GIÁM ĐỐC','','','','']);
    aoa.push(['','','','','','','','','','PHÓ GIÁM ĐỐC','','','','']);
    aoa.push([]);aoa.push([]);aoa.push([]);
    aoa.push(['Trương Tấn Tài','','','','Trần Trọng Khiêm','','','','','Trương Minh Hải','','','','']);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const sigTitleRow = ni + 5;
    const sigSubTitleRow = ni + 6;
    const sigNameRow = ni + 10;
    ws['!merges']=[
      M(0,0,0,4),M(0,5,0,13), M(1,0,1,4),M(1,5,1,13), M(2,0,2,4),M(2,5,2,13),
      M(4,0,4,13),M(5,0,5,13),M(6,0,6,13),M(7,0,7,6),M(7,7,7,13),
      M(9,0,10,0),M(9,1,10,1),M(9,2,10,2),M(9,3,10,3),M(9,4,10,4),
      M(9,5,9,8),M(9,9,10,9),M(9,10,10,10),M(9,11,10,11),M(9,12,10,12),M(9,13,10,13),
      M(ni,0,ni,13),
      M(ni+1,1,ni+1,3),M(ni+1,5,ni+1,7),
      M(ni+2,1,ni+2,3),M(ni+2,5,ni+2,7),
      M(ni+3,1,ni+3,3),M(ni+3,5,ni+3,7),
      M(sigTitleRow,0,sigTitleRow,3),M(sigTitleRow,4,sigTitleRow,8),M(sigTitleRow,9,sigTitleRow,13),
      M(sigSubTitleRow,0,sigSubTitleRow,3),M(sigSubTitleRow,4,sigSubTitleRow,8),M(sigSubTitleRow,9,sigSubTitleRow,13),
      M(sigNameRow,0,sigNameRow,3),M(sigNameRow,4,sigNameRow,8),M(sigNameRow,9,sigNameRow,13),
    ];
    setRowHeights(ws, [
      {row:0, hpt:18}, {row:1, hpt:18}, {row:2, hpt:18},
      {row:4, hpt:22}, {row:5, hpt:18}, {row:6, hpt:18}, {row:7, hpt:18},
      {row:9, hpt:24}, {row:10, hpt:24}, {row:ni, hpt:16},
      {row:sigTitleRow, hpt:17}, {row:sigSubTitleRow, hpt:17}, {row:sigNameRow, hpt:17}
    ]);

    const borderStyle = {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' }
    };

    applyRangeStyle(ws, 0, 0, 0, 4, {
      font: { name: 'Times New Roman', sz: 12 },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    });
    applyRangeStyle(ws, 1, 2, 0, 4, {
      font: { name: 'Times New Roman', sz: 12, bold: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    });
    applyRangeStyle(ws, 0, 0, 5, 13, {
      font: { name: 'Times New Roman', sz: 12, bold: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    });
    applyRangeStyle(ws, 1, 1, 5, 13, {
      font: { name: 'Times New Roman', sz: 12, bold: true, underline: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    });
    applyRangeStyle(ws, 2, 2, 5, 13, {
      font: { name: 'Times New Roman', sz: 11, italic: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    });
    applyRangeStyle(ws, 4, 4, 0, 13, {
      font: { name: 'Times New Roman', sz: 12, bold: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    });
    applyRangeStyle(ws, 5, 6, 0, 13, {
      font: { name: 'Times New Roman', sz: 11, bold: true },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
    });
    applyRangeStyle(ws, 7, 7, 0, 6, {
      font: { name: 'Times New Roman', sz: 11, bold: true },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
    });
    applyRangeStyle(ws, 7, 7, 7, 13, {
      font: { name: 'Times New Roman', sz: 11, italic: true },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
    });
    applyRangeStyle(ws, 9, 9, 0, 13, {
      font: { name: 'Times New Roman', sz: 10, bold: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: borderStyle
    });
    applyRangeStyle(ws, 10, 10, 0, 13, {
      font: { name: 'Times New Roman', sz: 9 },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: borderStyle
    });

    ws['!cols']=[{wch:3.5},{wch:8.5},{wch:14.5},{wch:6.2},{wch:10.5},{wch:5.2},{wch:5.2},{wch:5.2},{wch:5.2},{wch:6.1},{wch:8.6},{wch:5.3},{wch:4.8},{wch:9.6}];
    ws['!pageSetup'] = { paperSize: 9, orientation: 'landscape', fitToWidth: 1, fitToHeight: 0 };
    ws['!margins'] = { left: 0.25, right: 0.25, top: 0.3, bottom: 0.3, header: 0.2, footer: 0.2 };
    for(let r=11; r<=10 + students.length; r++){
      applyRangeStyle(ws, r, r, 0, 13, {
        font: { name: 'Times New Roman', sz: 10 },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: borderStyle
      });
      applyRangeStyle(ws, r, r, 2, 3, {
        font: { name: 'Times New Roman', sz: 10 },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: false },
        border: borderStyle
      });
      applyRangeStyle(ws, r, r, 1, 1, {
        font: { name: 'Times New Roman', sz: 10 },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
        border: borderStyle
      });
      applyRangeStyle(ws, r, r, 4, 4, {
        font: { name: 'Times New Roman', sz: 10 },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
        border: borderStyle
      });
      applyRangeStyle(ws, r, r, 13, 13, {
        font: { name: 'Times New Roman', sz: 10 },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
        border: borderStyle
      });
      applyRangeStyle(ws, r, r, 10, 10, {
        font: { name: 'Times New Roman', sz: 10 },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
        border: borderStyle
      });
    }
    applyRangeStyle(ws, ni, ni, 0, 13, {
      font: { name: 'Times New Roman', sz: 11, italic: true },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
    });
    applyRangeStyle(ws, ni + 1, ni + 3, 0, 13, {
      font: { name: 'Times New Roman', sz: 11 },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
    });
    applyRangeStyle(ws, ni + 1, ni + 3, 4, 4, {
      font: { name: 'Times New Roman', sz: 11, bold: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: false }
    });
    applyRangeStyle(ws, ni + 1, ni + 3, 5, 5, {
      font: { name: 'Times New Roman', sz: 11, italic: true },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: false }
    });
    applyRangeStyle(ws, sigTitleRow, sigNameRow, 0, 13, {
      font: { name: 'Times New Roman', sz: 11 },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: false }
    });
    applyRangeStyle(ws, sigNameRow, sigNameRow, 0, 13, {
      font: { name: 'Times New Roman', sz: 11, bold: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: false }
    });
    applyRangeStyle(ws, sigTitleRow, sigSubTitleRow, 0, 13, {
      font: { name: 'Times New Roman', sz: 12, bold: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: false }
    });
    centerSheetCells(ws, 11, 11 + students.length - 1, [5,6,7,8,9]);
    centerSheetCells(ws, sigTitleRow, sigNameRow, [0,1,2,3,4,5,6,7,8,9,10,11,12,13]);
    XLSX.utils.book_append_sheet(wb, ws, key.replace(/[\\\/\?\*\[\]:]/g,'_').substring(0,31));
  });

  const t=new Date();
  const d = t.getFullYear()+String(t.getMonth()+1).padStart(2,'0')+String(t.getDate()).padStart(2,'0');
  const fileName = scope==='current'
    ? `BangDiem_GDQP_${keys[0]}_${d}.xlsx`
    : `BangDiem_GDQP_All_${d}.xlsx`;
  XLSX.writeFile(wb, fileName, {bookSST:false, type:'binary', cellStyles:true});
}

function exportQuyetDinh(scope='all'){
  const keys = [...selLop].sort();
  if(!keys.length){alert('Chưa chọn lớp nào!');return}
  const wb=XLSX.utils.book_new();
  const cfg = getExportSettings();
  const school=cfg.school, centerTop=cfg.centerTop, centerBottom=cfg.centerBottom, xdate=cfg.xdate;
  const nguoiKy=cfg.nguoiKy, chucVu=cfg.chucVu;
  const M=(r1,c1,r2,c2)=>({s:{r:r1,c:c1},e:{r:r2,c:c2}});

  const list=getQDCombinedList(keys);
  const aoa=[];
  aoa.push(['ĐẠI HỌC TRÀ VINH','','','','','CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM','','','','']);
  aoa.push([centerTop,'','','','','Độc lập - Tự do - Hạnh phúc','','','','']);
  aoa.push([centerBottom,'','','','',xdate,'','','','']);
  aoa.push([]);
  aoa.push(['DANH SÁCH CẤP CHỨNG CHỈ GIÁO DỤC QUỐC PHÒNG VÀ AN NINH','','','','','','','','','']);
  aoa.push(['(Ban hành kèm theo Quyết định số      /QĐ-GDQP ngày      tháng      năm      )','','','','','','','','','']);
  aoa.push([]);
  aoa.push(['TT','Mã SV','Họ Và','Tên','Ngày Sinh','Nơi Sinh','Điểm TB','Xếp Loại','Ghi Chú','Mã Lớp']);

  list.forEach((s,i)=>{
    const tb=calcTB(s); const xl=xepLoai(s,tb);
    aoa.push([i+1, s.maSV, s.hoLot, s.ten, s.ngay,
      s.noiSinh||'',
      allHPAreCT(s) ? 'CT' : (tb!==null ? tb : ''),
      xl,
      '',
      s.maLop
    ]);
  });

  const fi2=aoa.length; aoa.push([]);
  aoa.push(['Trên danh sách có '+list.length+' sinh viên','','','','','','','','','']);
  aoa.push([]);
  const signerTitleText = String(chucVu || '').replace(' - ', '\n');
  aoa.push(['','','','','','',signerTitleText,'','','']);
  aoa.push([]);aoa.push([]);aoa.push([]);
  aoa.push(['','','','','','',nguoiKy,'','','']);

  const ws=XLSX.utils.aoa_to_sheet(aoa);
  const signerTitleRow = fi2 + 3;
  const signerNameRow = fi2 + 7;
  ws['!merges']=[
    M(0,0,0,4),M(0,5,0,9), M(1,0,1,4),M(1,5,1,9), M(2,0,2,4),M(2,5,2,9),
    M(4,0,4,9),M(5,0,5,9),
    M(fi2+1,0,fi2+1,9),
    M(signerTitleRow,6,signerTitleRow,9),
    M(signerNameRow,6,signerNameRow,9),
  ];
  setRowHeights(ws, [
    {row:0, hpt:20}, {row:1, hpt:20}, {row:2, hpt:20},
    {row:4, hpt:24}, {row:5, hpt:20}, {row:7, hpt:30},
    {row:fi2+1, hpt:18}, {row:signerTitleRow, hpt:32}, {row:signerNameRow, hpt:18}
  ]);

  const borderStyle = {
    top: { style: 'thin' },
    bottom: { style: 'thin' },
    left: { style: 'thin' },
    right: { style: 'thin' }
  };

  applyRangeStyle(ws, 0, 0, 0, 4, {
    font: { name: 'Times New Roman', sz: 12 },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
  });
  applyRangeStyle(ws, 1, 2, 0, 4, {
    font: { name: 'Times New Roman', sz: 12, bold: true },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
  });
  applyRangeStyle(ws, 0, 2, 5, 9, {
    font: { name: 'Times New Roman', sz: 12, bold: true },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
  });
  applyRangeStyle(ws, 1, 1, 5, 9, {
    font: { name: 'Times New Roman', sz: 12, bold: true, underline: true },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
  });
  applyRangeStyle(ws, 2, 2, 5, 9, {
    font: { name: 'Times New Roman', sz: 11, italic: true },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
  });
  applyRangeStyle(ws, 4, 4, 0, 9, {
    font: { name: 'Times New Roman', sz: 13, bold: true },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
  });
  applyRangeStyle(ws, 5, 5, 0, 9, {
    font: { name: 'Times New Roman', sz: 11, italic: true },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
  });
  applyRangeStyle(ws, 7, 7, 0, 9, {
    font: { name: 'Times New Roman', sz: 10, bold: true },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
    border: borderStyle
  });

  ws['!cols']=[{wch:3.5},{wch:9},{wch:14.5},{wch:6.2},{wch:10.5},{wch:11},{wch:6},{wch:8},{wch:5},{wch:9.5}];
  ws['!pageSetup'] = { paperSize: 9, orientation: 'landscape', fitToWidth: 1, fitToHeight: 0 };
  ws['!margins'] = { left: 0.25, right: 0.25, top: 0.3, bottom: 0.3, header: 0.2, footer: 0.2 };
  for(let r=8; r<=8 + list.length - 1; r++){
    applyRangeStyle(ws, r, r, 0, 9, {
      font: { name: 'Times New Roman', sz: 10 },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
      border: borderStyle
    });
    applyRangeStyle(ws, r, r, 2, 3, {
      font: { name: 'Times New Roman', sz: 10 },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: false },
      border: borderStyle
    });
    applyRangeStyle(ws, r, r, 4, 4, {
      font: { name: 'Times New Roman', sz: 10 },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
      border: borderStyle
    });
    applyRangeStyle(ws, r, r, 6, 6, {
      font: { name: 'Times New Roman', sz: 10 },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
      border: borderStyle
    });
  }
  applyRangeStyle(ws, fi2+1, fi2+1, 0, 9, {
    font: { name: 'Times New Roman', sz: 11, italic: true },
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
  });
  applyRangeStyle(ws, signerTitleRow, signerNameRow, 0, 9, {
    font: { name: 'Times New Roman', sz: 11 },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
  });
  applyRangeStyle(ws, signerTitleRow, signerTitleRow, 0, 9, {
    font: { name: 'Times New Roman', sz: 11, bold: true },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
  });
  applyRangeStyle(ws, signerNameRow, signerNameRow, 0, 9, {
    font: { name: 'Times New Roman', sz: 11, bold: true },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
  });
  centerSheetCells(ws, signerTitleRow, signerNameRow, [6]);
  XLSX.utils.book_append_sheet(wb,ws, 'TongHop_QD');

  const t=new Date();
  const d = t.getFullYear()+String(t.getMonth()+1).padStart(2,'0')+String(t.getDate()).padStart(2,'0');
  const fileName = `QuyetDinh_GDQP_TongHop_${d}.xlsx`;
  XLSX.writeFile(wb, fileName, {bookSST:false, type:'binary', cellStyles:true});
}

function clearData(){
  sessionStorage.removeItem(SS_KEY_DATA);
  sessionStorage.removeItem(SS_KEY_FILE);
  sessionStorage.removeItem(SS_KEY_BIRTH_FILE);
  sessionStorage.removeItem('gdqp_selLop');
  byLop={}; selLop=new Set(); allKeys=[]; activeBD=null; activeQD=null; birthFileName='';
  document.getElementById('zone').style.display='block';
  document.getElementById('fileBanner').style.display='none';
  document.getElementById('mainPanel').style.display='none';
  document.getElementById('msg').style.display='none';
  fi.value='';
  fiNoiSinh.value='';
}

window.addEventListener('DOMContentLoaded', ()=>{
  const params = new URLSearchParams(window.location.search);
  if(params.get('embed')==='1') setExternalUploadMode(true);

  const hasData = restoreFromSession();
  if(hasData){
    loadSelLop();
    renderList();
    renderTags();
    renderTabs();
    renderQDTabs();
  }
});

window.addEventListener('message', (event)=>{
  const data = event.data || {};
  if(data.type==='gdqp-load-main' && data.file){
    loadFile(data.file, true);
  }
  if(data.type==='gdqp-load-birth' && data.file){
    loadNoiSinhFile(data.file, true);
  }
  if(data.type==='gdqp-toggle-external-upload'){
    setExternalUploadMode(!!data.enabled);
  }
});
