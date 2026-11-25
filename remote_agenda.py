"""Browser agenda editor and read-only schedule for a paired meeting room."""

AGENDA_HTML = """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="referrer" content="no-referrer">
<title>会议议程 · ppt time</title>
<link rel="stylesheet" href="/agenda.css"><script defer src="/agenda.js"></script></head>
<body><main><header><a href="/" id="desk">返回会场工作台</a><span id="venue"></span></header>
<div class="heading"><h1>会议议程</h1><button hidden type="button" id="add">添加议程</button></div>
<p id="feedback" role="status" aria-live="polite">正在连接会场…</p>
<div id="workspace" hidden>
<form id="filters" aria-label="议程筛选">
<label>会议<select id="meeting"><option value="">全部会议</option></select></label>
<label>会场<select id="room"><option value="">全部会场</option></select></label>
<label>日期<select id="date"><option value="">全部日期</option></select></label>
<label>时段<select id="period"><option value="">全天</option><option value="am">上午</option><option value="pm">下午及晚上</option></select></label>
<label>讲者<input id="speaker" type="search" placeholder="搜索讲者"></label>
<label>单位<input id="organization" type="search" placeholder="搜索单位"></label>
<label>文稿<select id="uploaded"><option value="">全部文稿</option><option value="received">本场已收件</option><option value="waiting">本场待收件</option><option value="yes">已关联文件名</option><option value="no">未关联文件名</option></select></label>
<label>状态<select id="status"><option value="">全部状态</option><option>待开始</option><option>进行中</option><option>已完成</option></select></label>
</form>
<div class="viewbar"><span id="summary" aria-live="polite"></span><label>视图
<select id="view"><option value="list">列表</option><option value="cards">卡片</option><option value="timeline">时间线</option></select></label>
<button type="button" id="reset">清除筛选</button></div>
<div id="rows"></div></div>
<dialog id="editor" aria-labelledby="form-title"><form id="edit-form">
<h2 id="form-title">编辑议程</h2>
<div class="fields">
<label class="wide">议题<input name="title" required maxlength="250"></label>
<label>会议<input name="meeting" maxlength="250" list="meeting-names"></label>
<label>会场<input name="room" maxlength="250" list="room-names"></label>
<label>日期<input name="date" type="date"></label>
<label>开始时间<input name="start" maxlength="250" placeholder="09:30"></label>
<label>讲者<input name="speaker" maxlength="250"></label>
<label>单位<input name="organization" maxlength="250"></label>
<label>计时秒数<input name="seconds" type="number" min="0" max="86400" step="1" inputmode="numeric"></label>
<label>状态<select name="status"><option>待开始</option><option>进行中</option><option>已完成</option></select></label>
<label class="wide">关联的文稿文件名<input name="file" maxlength="250" list="file-names" placeholder="例：开幕报告.pptx"></label>
</div><p id="edit-message" aria-live="polite"></p>
<footer><button type="button" id="cancel">取消</button><button type="submit" id="save">保存</button></footer>
</form></dialog>
<datalist id="meeting-names"></datalist><datalist id="room-names"></datalist><datalist id="file-names"></datalist>
</main></body></html>"""


AGENDA_CSS = """
:root {color-scheme:light; --ink:#18312f; --muted:#526662; --line:#ccd9d4; --paper:#fff; --green:#12634f; font-family:system-ui,"Microsoft YaHei",sans-serif}
* {box-sizing:border-box}
body {margin:0; background:#f3f6f4; color:var(--ink)}
main {max-width:1080px; margin:auto; padding:clamp(16px,3vw,36px)}
[hidden] {display:none!important}
header,.heading,.viewbar,footer,.row-actions {display:flex; align-items:center; gap:.6rem; flex-wrap:wrap}
header {justify-content:space-between; color:var(--muted); font-size:.9rem} a {color:var(--green)}
.heading {justify-content:space-between; margin:1.3rem 0 .8rem}
h1,h2,h3,p {margin:0} h1 {font-size:clamp(1.75rem,4vw,2.4rem)} h2 {font-size:1.3rem} h3 {font-size:1rem; overflow-wrap:anywhere}
button,select,input {font:inherit; border:1px solid #9eafa7; border-radius:.4rem; background:white; color:var(--ink); padding:.45rem .65rem; min-height:38px}
button {cursor:pointer} button:hover {border-color:var(--green)} button:disabled {opacity:.5;cursor:default}
#add,#save {background:var(--green);color:white;border-color:var(--green)}
#feedback:not(:empty), #edit-message:not(:empty) {padding:.65rem 0;color:var(--muted)}
#feedback.error,#edit-message.error {color:#9a321d}
#filters,.fields {display:grid;grid-template-columns:repeat(auto-fit,minmax(134px,1fr));gap:.65rem}
#filters {background:var(--paper);border:1px solid var(--line);border-radius:.7rem;padding:.9rem}
label {display:flex; flex-direction:column; gap:.25rem;font-size:.85rem;color:var(--muted)}
label input,label select {width:100%; font-size:.95rem}
.viewbar {margin:1.1rem 0}.viewbar label {display:flex;flex-direction:row;align-items:center}.viewbar select {width:auto}
#summary {flex:1;min-width:min(230px,100%);font-size:.95rem}
#rows {display:grid;gap:.5rem}
.session {padding:.85rem 1rem;display:grid;grid-template-columns:minmax(6rem,16%) minmax(0,1fr) auto;gap:.6rem;align-items:start;background:var(--paper);border:1px solid var(--line);border-radius:.6rem}
.time,.detail,.attachment {color:var(--muted);font-size:.88rem;overflow-wrap:anywhere}
.detail,.attachment {margin-top:.3rem}.time {font-variant-numeric:tabular-nums}
.row-actions button {font-size:.85rem;min-height:34px;padding:.25rem .5rem}
#rows.cards {grid-template-columns:repeat(auto-fill,minmax(min(100%,255px),1fr))}
.cards .session {display:flex;flex-direction:column}.cards .row-actions {margin-top:auto}
.timeline .session {border-inline-start:4px solid var(--green)} .day {margin:.8rem 0 .25rem;font-size:1.05rem}
dialog {width:min(96vw,570px);max-height:92vh;overflow:auto;border:1px solid var(--line);border-radius:.8rem;color:var(--ink);padding:1.25rem}
dialog::backdrop {background:#14242273} .fields {grid-template-columns:1fr 1fr;margin:1rem 0}.wide {grid-column:1/-1}
footer {justify-content:flex-end;margin-top:1rem}
@media(max-width:650px){.session {grid-template-columns:1fr}.row-actions {gap:.4rem}}
"""


AGENDA_JS = r"""
'use strict';
const el = id => document.getElementById(id);
const filterIds = ['meeting','room','date','period','speaker','organization','uploaded','status'];
const textFields = ['title','meeting','room','date','start','speaker','organization','status','file'];
const states = ['待开始','进行中','已完成'];
let token = location.hash.slice(1);
if (!token) { try { token = sessionStorage.getItem('meeting-link') || ''; } catch {} }
if (location.hash) history.replaceState(null, '', location.pathname);
if (token) el('desk').href = '/#' + encodeURIComponent(token);
let role = '', rows = [], revision = '', busy = false, polling = false, editIndex = null, editEpoch = 0;
let receivedFiles = new Set();
const form = el('edit-form'), editor = el('editor');

function message(text, error=false, field='feedback') {
  const target = el(field);
  target.textContent = text;
  target.classList.toggle('error', error);
}
function element(tag, text, style) {
  const node = document.createElement(tag);
  if (text) node.textContent = text;
  if (style) node.className = style;
  return node;
}
function orderedUnique(values) {
  return [...new Set(values.filter(Boolean))].sort((a,b)=>a.localeCompare(b,'zh-CN',{numeric:true}));
}
function setChoices(id, values, blank) {
  const input = el(id), old = input.value, options = [];
  if (blank) options.push(new Option(blank,''));
  for (const value of orderedUnique(values)) options.push(new Option(value,value));
  input.replaceChildren(...options);
  if (blank) input.value = values.includes(old) ? old : '';
}
function cleanRow(value) {
  const item = {uid:value.uid || ''};
  for (const key of textFields) item[key] = typeof value[key] === 'string' ? value[key] : '';
  item.status = item.status || states[0];
  item.seconds = Number.isInteger(value.seconds) ? value.seconds : 0;
  return item;
}
function decodeState(data) {
  if (data.role !== 'monitor' && data.role !== 'control') throw Error('请使用会场的监看或控制链接。');
  const state = data.state || {};
  if (!Array.isArray(state.agenda)) throw Error('会场未提供议程数据。');
  return {role:data.role, state, rows:state.agenda.map(cleanRow)};
}
async function api(path, options={}) {
  const response = await fetch(path,{...options, cache:'no-store',
    headers:{Authorization:'Bearer '+token,...options.headers}});
  const data = await response.json();
  if (!response.ok) throw Error(data.error || '连接失败，请重新打开会场配对链接。');
  return data;
}
function applyState(data) {
  role = data.role;
  rows = data.rows;
  revision = String(data.state.agenda_revision || '');
  el('venue').textContent = [data.state.room, role==='control'?'控制端':'只读监看'].filter(Boolean).join(' · ');
  el('workspace').hidden = false;
  el('add').hidden = role!=='control';
  const meetingNames = [...(data.state.meetings || []), ...rows.map(row=>row.meeting)];
  const roomNames = rows.map(row=>row.room);
  const files = (data.state.files || []).filter(name=>typeof name==='string');
  receivedFiles = new Set(files.map(name=>name.toLocaleLowerCase()));
  setChoices('meeting',meetingNames,'全部会议'); setChoices('room',roomNames,'全部会场');
  setChoices('date',rows.map(row=>row.date),'全部日期');
  setChoices('meeting-names',meetingNames); setChoices('room-names',roomNames);
  setChoices('file-names',[...files,...rows.map(row=>row.file)]);
  render();
}
function periodOf(start) {
  const match = /^(\d{1,2})\s*[:：]\s*([0-5]\d)/.exec(start.trim());
  if (!match || +match[1] >= 24) return '';
  return +match[1] < 12 ? 'am' : 'pm';
}
function matches(row, filters) {
  for (const key of ['meeting','room','date','status']) if (filters[key] && row[key]!==filters[key]) return false;
  for (const key of ['speaker','organization']) {
    if (!row[key].toLocaleLowerCase().includes(filters[key].trim().toLocaleLowerCase())) return false;
  }
  if (filters.period && periodOf(row.start)!==filters.period) return false;
  const received = receivedFiles.has(row.file.toLocaleLowerCase());
  return !filters.uploaded || ({received, waiting:!received, yes:!!row.file, no:!row.file})[filters.uploaded];
}
function clockRank(start) {
  const match = /^(\d{1,2})\s*[:：]\s*([0-5]\d)/.exec(start.trim());
  return match && +match[1] < 24 ? +match[1]*60 + +match[2] : 1440;
}
function button(label, handler, disabled=false) {
  const node = element('button',label);
  node.type='button'; node.disabled=disabled; node.addEventListener('click',handler);
  return node;
}
function render() {
  const filters = Object.fromEntries(filterIds.map(key=>[key,el(key).value]));
  const visible = rows.map((row,index)=>({row,index})).filter(item=>matches(item.row,filters));
  const mode = el('view').value;
  if (mode==='timeline') visible.sort((a,b)=>(a.row.date || '9999').localeCompare(b.row.date || '9999')
    || clockRank(a.row.start)-clockRank(b.row.start) || a.index-b.index);
  const counts = states.map(status=>status+' '+visible.filter(item=>item.row.status===status).length);
  const received = visible.filter(item=>receivedFiles.has(item.row.file.toLocaleLowerCase())).length;
  el('summary').textContent = [visible.length+' 项议程', ...counts, '本场已收件 '+received].join(' · ');
  const box=el('rows'); box.className=mode;
  const contents = [];
  let previousDate=null;
  for (const {row,index} of visible) {
    if (mode==='timeline' && row.date!==previousDate) {
      previousDate=row.date; contents.push(element('h2',row.date || '日期待定','day'));
    }
    const article = element('article','','session');
    const stamp = [mode==='timeline'?'':row.date,row.start].filter(Boolean).join(' ');
    article.append(element('span',stamp || '时间待定','time'));
    const body = element('div');
    body.append(element('h3',row.title));
    body.append(element('p',[row.meeting,row.room,row.speaker,row.organization,row.status,
      row.seconds?row.seconds+' 秒':''].filter(Boolean).join(' · '),'detail'));
    if (row.file) body.append(element('p',row.file+' · '+(receivedFiles.has(row.file.toLocaleLowerCase())?'已在本场文件看板':'尚未在本场文件看板'),'attachment'));
    article.append(body);
    if (role==='control') {
      const actions=element('div','','row-actions');
      actions.append(button('编辑',()=>openEditor(index),busy));
      if (mode!=='timeline') {
        actions.append(button('上移',()=>move(index,-1),busy || index===0));
        actions.append(button('下移',()=>move(index,1),busy || index===rows.length-1));
      }
      actions.append(button('删除',()=>remove(index),busy));
      article.append(actions);
    }
    contents.push(article);
  }
  if (!contents.length) contents.push(element('p',rows.length?'没有符合筛选条件的议程。':'当前会场尚无议程。'));
  box.replaceChildren(...contents);
  el('add').disabled=busy;
}
async function refresh() {
  if (polling || busy || editor.open) return;
  polling=true;
  const epoch = editEpoch;
  try {
    const data = decodeState(await api('/api/state'));
    if (!busy && !editor.open && epoch===editEpoch) {
      const first = !role;
      applyState(data);
      if (first || el('feedback').classList.contains('error')) message('');
      try { sessionStorage.setItem('meeting-link',token); } catch {}
    }
  } catch (err) { if (!busy && epoch===editEpoch) message(err.message,true); }
  finally {polling=false;}
}
function makeUid() {
  return Array.from(crypto.getRandomValues(new Uint8Array(16)), v=>v.toString(16).padStart(2,'0')).join('');
}
function openEditor(index) {
  if (role!=='control' || busy) return;
  editIndex=index;
  const row = index===null ? cleanRow({meeting:el('meeting').value,room:el('room').value,date:el('date').value}) : rows[index];
  for (const key of [...textFields,'seconds']) form.elements.namedItem(key).value=row[key];
  el('form-title').textContent=index===null?'添加议程':'编辑议程';
  message('',false,'edit-message');
  editor.showModal();
  form.elements.namedItem('title').focus();
}
async function saveRows(next) {
  if (role!=='control' || busy) return false;
  editEpoch++;
  busy=true; el('save').disabled=true; el('cancel').disabled=true; render();
  const base=JSON.stringify(rows), target=JSON.stringify(next);
  const channel = editor.open ? 'edit-message' : 'feedback';
  try {
    const current=decodeState(await api('/api/state'));
    if (current.role!=='control') throw Error('仅控制端可以编辑议程。');
    if (JSON.stringify(current.rows)!==base) {
      if (editor.open) editor.close();
      applyState(current);
      throw Error('会场的议程已更新，请核对后重新操作。');
    }
    const headers={'Content-Type':'application/json'};
    if (revision) headers['If-Match']=revision;
    const compact = next.map(row=>Object.fromEntries(Object.entries(row).filter(([,value])=>value!=='')));
    await api('/api/agenda',{method:'POST',headers,body:JSON.stringify(compact)});
    if (editor.open) editor.close();
    for (let n=0; n<9; n++) {
      await new Promise(resolve=>setTimeout(resolve, 250));
      const latest=decodeState(await api('/api/state'));
      if (JSON.stringify(latest.rows)===target) {
        applyState(latest); message('议程已保存。'); return true;
      }
    }
    applyState(decodeState(await api('/api/state')));
    message('会场已接收请求，正在等待更新。'); return true;
  } catch (err) {
    message(err.message,true,editor.open ? channel : 'feedback');
    return false;
  } finally {
    busy=false; el('save').disabled=false; el('cancel').disabled=false; render();
  }
}
function submit(event) {
  event.preventDefault();
  if (role!=='control' || busy || !form.reportValidity()) return;
  const item = cleanRow(editIndex===null?{uid:makeUid()}:rows[editIndex]);
  for (const key of textFields) item[key]=form.elements.namedItem(key).value.trim();
  item.seconds=Number(form.elements.namedItem('seconds').value);
  if (!item.title || !Number.isInteger(item.seconds) || item.seconds<0 || item.seconds>86400 ||
      (item.file && (/[\\/:]/.test(item.file) || item.file==='.' || item.file==='..'))) {
    message('请填写议题、有效的计时秒数和不含路径的文件名。',true,'edit-message');return;
  }
  const next=rows.slice();
  if (editIndex===null) next.push(item); else next[editIndex]=item;
  saveRows(next);
}
function move(index, step) {
  const next=rows.slice(), target=index+step;
  if (target<0 || target>=next.length) return;
  [next[target],next[index]]=[next[index],next[target]];
  saveRows(next);
}
function remove(index) {
  if (confirm('删除议程“'+rows[index].title+'”？')) saveRows(rows.filter((_,i)=>i!==index));
}
el('filters').addEventListener('input',render);
el('filters').addEventListener('submit',ev=>ev.preventDefault());
el('reset').addEventListener('click',()=>{el('filters').reset();render();});
el('view').addEventListener('change',render);
el('add').addEventListener('click',()=>openEditor(null));
el('cancel').addEventListener('click',()=>editor.close());
editor.addEventListener('cancel',event=>{if (busy) event.preventDefault();});
form.addEventListener('submit',submit);
refresh();
setInterval(refresh,4000);
"""
