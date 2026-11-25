"""Appearance controls embedded in the authenticated meeting operator's desk."""

from overlay_options import SIZE_MODES


def _number(key, label, minimum, maximum):
    return (f'<label>{label}<input name="{key}" type="number" step="1" '
            f'min="{minimum}" max="{maximum}" required></label>')


def _flag(key, label):
    return f'<label class="ap-flag"><input type="checkbox" name="{key}">{label}</label>'


APPEARANCE_HTML = (
    '<details id="ap-panel"><summary>放映条外观</summary><form id="ap-form">'
    '<fieldset disabled id="ap-fields"><legend>大小与位置</legend><div class="ap-grid">'
    '<label>放映条大小<select name="size_mode">'
    + ''.join(f'<option value="{key}">{label}</option>' for key, label in SIZE_MODES.items())
    + '</select></label><label>放映条位置<select name="anchor">'
    + ''.join(f'<option value="{key}">{label}</option>' for key, label in (
        ("top-center", "顶部居中"), ("top-left", "左上"), ("top-right", "右上"),
        ("bottom-center", "底部居中"), ("bottom-left", "左下"), ("bottom-right", "右下")))
    + '</select></label>'
    + _number("x_offset", "水平偏移（px）", -4096, 4096)
    + _number("edge_gap", "距上下边缘（−1 随样式）", -1, 2048)
    + _number("radius", "圆角（−1 随样式）", -1, 96)
    + _flag("secondary_offsets", "扩展屏单独设置偏移")
    + _number("secondary_x", "扩展屏水平偏移（px）", -4096, 4096)
    + _number("secondary_gap", "扩展屏边距（−1 随样式）", -1, 2048)
    + '</div><h3>数字和颜色</h3><div class="ap-grid">'
    + _number("time_size", "时间字号（px）", 10, 80) + _number("page_size", "页码字号（px）", 10, 80)
    + '<label>字体<input name="font" maxlength="128" placeholder="随系统" autocomplete="off"></label>'
    + _flag("lcd", "七段液晶数字") + _flag("bold", "加粗数字")
    + ''.join(f'<label>{label}<input name="{key}" pattern="#[0-9A-Fa-f]{{6}}" maxlength="7" '
              f'placeholder="随基础设置" aria-description="六位十六进制颜色，例如 #ffffff，留空沿用基础设置。"></label>'
              for key, label in (("foreground", "文字颜色（#RRGGBB）"), ("background", "背景颜色（#RRGGBB）")))
    + '</div><h3>提示与录屏</h3><div class="ap-grid">'
    + _flag("flash_last_minute", "最后 60 秒闪动") + _flag("flash_expiry", "到时闪动 4 秒")
    + _flag("smart_fade", "智能闪避")
    + _number("fade_percent", "闪避时保留的不透明度（%）", 0, 95)
    + _flag("exclude_capture", "在兼容的录屏工具中隐藏")
    + '</div></fieldset><div class="actions"><button id="ap-apply" class="primary" disabled>应用到讲台</button>'
    '<button id="ap-reset" type="button" disabled>放弃修改</button></div>'
    '<output id="ap-status" aria-live="polite">正在等待讲台外观设置。</output></form></details>'
)

APPEARANCE_CSS = """
#ap-panel{border-top:1px solid #d5dedb;margin-top:16px;padding-top:12px}
#ap-panel summary{cursor:pointer;color:#174644;font-weight:600}
#ap-form{display:block;margin:10px 0 0}#ap-fields{border:0;margin:0;padding:0;min-width:0}
#ap-fields legend,#ap-fields h3{font-size:14px;font-weight:600;margin:14px 0 7px;padding:0;color:#315455}
.ap-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,156px),1fr));gap:9px 12px}
.ap-grid label{display:flex;flex-direction:column;align-items:stretch;gap:4px;min-width:0;color:#315455;font-size:13px}
.ap-grid input:not([type=checkbox]),.ap-grid select{width:100%;min-width:0;padding:6px 8px}
.ap-grid .ap-flag{flex-direction:row;align-items:center;line-height:1.4;gap:6px;min-height:34px}
.ap-grid input[type=checkbox]{width:16px;height:16px;flex-shrink:0;accent-color:#126455}
#ap-status{font-size:13px;color:#426463}#ap-form button:disabled{opacity:.55;cursor:default}
"""

# Mutating a single control sends only that field. Polling still updates every
# untouched setting, including when a second operator edits the desktop dialog.
APPEARANCE_JS = r"""
const apForm = el('ap-form'), apFields = el('ap-fields');
const apControls = Array.from(apFields.querySelectorAll('[name]'));
let apLatest = null, apPending = null;
const apDirty = new Set();
function apValue(control) {
  if(control.type === 'checkbox')return control.checked;
  if(control.type === 'number')return control.valueAsNumber;
  return ['foreground','background'].includes(control.name) ? control.value.toLowerCase() : control.value;
}
function apSet(control,value) {
  if(control.type === 'checkbox')control.checked=value === true;
  else control.value=value;
}
function apButtons() {
  el('ap-apply').disabled=!apLatest || !apDirty.size || !!apPending;
  el('ap-reset').disabled=!apDirty.size || !!apPending;
  apFields.disabled=!apLatest || !!apPending;
  const field=name=>apForm.elements.namedItem(name);
  field('font').disabled=field('lcd').checked;
  field('secondary_x').disabled=field('secondary_gap').disabled=!field('secondary_offsets').checked;
  field('fade_percent').disabled=!field('smart_fade').checked;
}
function loadAppearance(state) {
  if(!state || typeof state !== 'object' || Array.isArray(state))return;
  if(!apControls.every(c=>Object.prototype.hasOwnProperty.call(state,c.name)))return;
  const first=apLatest===null;
  apLatest=state;
  if(apPending && Object.entries(apPending.patch).every(([key,v])=>state[key]===v)) {
    apPending=null;apDirty.clear();el('ap-status').textContent='讲台已应用。';
  } else if(apPending && performance.now()-apPending.started>10000) {
    apPending=null;el('ap-status').textContent='尚未收到讲台确认。可重试或放弃修改。';
  } else if(first)el('ap-status').textContent='';
  for(const c of apControls)if(!apDirty.has(c.name))apSet(c,state[c.name]);
  apButtons();
}
apForm.addEventListener('input',event=>{
  const c=event.target;if(!apControls.includes(c) || !apLatest || apPending)return;
  if(apValue(c)===apLatest[c.name])apDirty.delete(c.name);else apDirty.add(c.name);
  el('ap-status').textContent=apDirty.size?'有尚未应用的修改。':'';
  apButtons();
});
el('ap-reset').addEventListener('click',()=>{
  if(!apLatest || apPending)return;
  apDirty.clear();el('ap-status').textContent='';loadAppearance(apLatest);
});
apForm.addEventListener('submit',async event=>{
  event.preventDefault();
  if(!apLatest || !apDirty.size || apPending || !apForm.reportValidity())return;
  const patch={};
  for(const c of apControls)if(apDirty.has(c.name)) {
    if(c.type==='number' && !Number.isInteger(apValue(c)))return;
    patch[c.name]=apValue(c);
  }
  const pending={patch,started:performance.now()};apPending=pending;apButtons();
  el('ap-status').textContent='正在应用到讲台…';
  try {
    await api('/api/action',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({action:'set_appearance',value:patch})});
    if(apPending===pending)el('ap-status').textContent='已送达，等待讲台确认。';
  } catch(e) {
    if(apPending===pending){apPending=null;apButtons();el('ap-status').textContent=e.message;}
  }
});
"""
