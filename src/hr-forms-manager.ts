import * as XLSX from 'xlsx';

const LEGACY_KEY = 'hr-saudi-forms-v1';
const KEY = 'hr-saudi-forms-v2';
const TYPES = ['عرض عمل','مباشرة عمل','إخلاء طرف','طلب إجازة','طلب سلفة','تقييم التجربة'];
type Row = { id:string; type:string; employeeId:string; employeeName:string; status:string; createdAt:string; updatedAt:string; data:Record<string,any> };

const read = (key:string, fallback:any) => { try { const v=localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch { return fallback; } };
const write = (key:string, value:any) => localStorage.setItem(key, JSON.stringify(value));
const esc = (v:any) => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c] || c));
const makeId = () => `FORM-${new Date().toISOString().replace(/\D/g,'').slice(0,14)}-${Math.floor(1000+Math.random()*9000)}`;
const fmt = (v:string) => v ? new Date(v).toLocaleString('ar-SA') : '—';

function migrate():Row[] {
  const current = read(KEY, []);
  if (Array.isArray(current)) return current;
  const legacy = read(LEGACY_KEY, {});
  const rows:Row[] = [];
  for (const type of TYPES) {
    const data = legacy?.[type];
    if (!data) continue;
    rows.push({ id:makeId(), type, employeeId:data.employeeId||'', employeeName:data.name||'', status:data.status||'مسودة', createdAt:data.updatedAt||new Date().toISOString(), updatedAt:data.updatedAt||new Date().toISOString(), data });
  }
  write(KEY, rows);
  return rows;
}

function sync():Row[] {
  const legacy = read(LEGACY_KEY, {});
  let rows = migrate();
  let changed = false;
  for (const type of TYPES) {
    const data = legacy?.[type];
    if (!data) continue;
    const existing = rows.find(r => r.type === type);
    if (!existing) {
      rows.push({id:makeId(), type, employeeId:data.employeeId||'', employeeName:data.name||'', status:data.status||'مسودة', createdAt:new Date().toISOString(), updatedAt:new Date().toISOString(), data});
      changed = true;
    } else if (JSON.stringify(existing.data) !== JSON.stringify(data)) {
      existing.data = data;
      existing.employeeId = data.employeeId || '';
      existing.employeeName = data.name || '';
      existing.status = data.status || 'مسودة';
      existing.updatedAt = new Date().toISOString();
      changed = true;
    }
  }
  if (changed) write(KEY, rows);
  return rows;
}

function download(name:string, blob:Blob) {
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}

function exportAll() {
  const rows = sync();
  const data = rows.map(r => ({'رقم النموذج':r.id,'نوع النموذج':r.type,'الموظف':r.employeeName,'رقم الموظف':r.employeeId,'الحالة':r.status,'تاريخ الإنشاء':fmt(r.createdAt),'آخر تعديل':fmt(r.updatedAt), ...r.data}));
  const ws=XLSX.utils.json_to_sheet(data.length?data:[{'رسالة':'لا توجد نماذج محفوظة'}]);
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'النماذج'); XLSX.writeFile(wb,`HR-Forms-${new Date().toISOString().slice(0,10)}.xlsx`);
}

function exportOne(r:Row) {
  const row:any={'رقم النموذج':r.id,'نوع النموذج':r.type,'الموظف':r.employeeName,'رقم الموظف':r.employeeId,'الحالة':r.status,'تاريخ الإنشاء':fmt(r.createdAt),'آخر تعديل':fmt(r.updatedAt),...r.data};
  const ws=XLSX.utils.json_to_sheet([row]); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'النموذج'); XLSX.writeFile(wb,`${r.id}.xlsx`);
}

function printOne(r:Row) {
  const entries=Object.entries(r.data||{}).filter(([,v])=>v!==undefined&&v!==null&&String(v)!=='').map(([k,v])=>`<tr><td>${esc(k)}</td><td>${esc(v)}</td></tr>`).join('');
  const w=window.open('','_blank','width=1000,height=800');
  if(!w){alert('يرجى السماح بالنوافذ المنبثقة للطباعة');return;}
  w.document.write(`<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>${esc(r.type)}</title><style>body{font-family:Arial,sans-serif;padding:32px;color:#111}h1{font-size:24px;margin-bottom:6px}p{color:#555}table{width:100%;border-collapse:collapse;margin-top:22px}td{border:1px solid #ccc;padding:10px}td:first-child{width:32%;font-weight:bold;background:#f5f5f5}</style></head><body><h1>${esc(r.type)}</h1><p>رقم النموذج: ${esc(r.id)} — الموظف: ${esc(r.employeeName||'—')} — آخر تعديل: ${esc(fmt(r.updatedAt))}</p><table>${entries}</table><script>setTimeout(()=>window.print(),250)<\/script></body></html>`);
  w.document.close();
}

function removeRow(id:string) {
  const rows=sync().filter(r=>r.id!==id); write(KEY,rows);
  const legacy=read(LEGACY_KEY,{}); const target=sync().find(r=>r.id===id); if(target) { delete legacy[target.type]; write(LEGACY_KEY,legacy); }
  window.location.reload();
}

function openEditor(type:string) {
  const cards=[...document.querySelectorAll('.report')];
  const card=cards.find(c=>c.querySelector('b')?.textContent?.trim()===type) as HTMLElement|undefined;
  (card?.querySelector('button') as HTMLButtonElement|undefined)?.click();
}

function openSaved() {
  const rows=sync().sort((a,b)=>b.updatedAt.localeCompare(a.updatedAt));
  const overlay=document.createElement('div'); overlay.className='modal-overlay'; overlay.innerHTML=`<div class="modal modal-wide" dir="rtl"><div class="modal-head"><div><h2>النماذج المحفوظة</h2><small>${rows.length} نموذج محفوظ</small></div><button class="modal-close">×</button></div><div class="modal-body"><div class="table-wrap"><table class="data-table"><thead><tr><th>الرقم</th><th>النموذج</th><th>الموظف</th><th>الحالة</th><th>آخر تعديل</th><th>إجراءات</th></tr></thead><tbody>${rows.length?rows.map(r=>`<tr><td><code>${esc(r.id)}</code></td><td>${esc(r.type)}</td><td>${esc(r.employeeName||'—')}</td><td><span class="badge">${esc(r.status)}</span></td><td>${esc(fmt(r.updatedAt))}</td><td><div class="row-actions"><button data-open="${esc(r.id)}">فتح</button><button data-print="${esc(r.id)}">طباعة</button><button data-xlsx="${esc(r.id)}">Excel</button><button data-del="${esc(r.id)}" class="danger">حذف</button></div></td></tr>`).join(''):`<tr><td colspan="6" style="text-align:center;padding:30px">لا توجد نماذج محفوظة حتى الآن</td></tr>`}</tbody></table></div></div></div>`;
  document.body.appendChild(overlay);
  (overlay.querySelector('.modal-close') as HTMLElement).onclick=()=>overlay.remove();
  overlay.addEventListener('click',e=>{if(e.target===overlay)overlay.remove();});
  overlay.querySelectorAll('[data-open]').forEach(b=>(b as HTMLElement).onclick=()=>{const r=rows.find(x=>x.id===(b as HTMLElement).dataset.open); overlay.remove(); if(r) openEditor(r.type);});
  overlay.querySelectorAll('[data-print]').forEach(b=>(b as HTMLElement).onclick=()=>{const r=rows.find(x=>x.id===(b as HTMLElement).dataset.print); if(r) printOne(r);});
  overlay.querySelectorAll('[data-xlsx]').forEach(b=>(b as HTMLElement).onclick=()=>{const r=rows.find(x=>x.id===(b as HTMLElement).dataset.xlsx); if(r) exportOne(r);});
  overlay.querySelectorAll('[data-del]').forEach(b=>(b as HTMLElement).onclick=()=>{const r=rows.find(x=>x.id===(b as HTMLElement).dataset.del); if(r&&confirm(`حذف نموذج ${r.type}؟`)) removeRow(r.id);});
}

function jsonBackup() {
  const payload={version:2,exportedAt:new Date().toISOString(),forms:sync()};
  download(`HR-Forms-${new Date().toISOString().slice(0,10)}.json`,new Blob([JSON.stringify(payload,null,2)],{type:'application/json'}));
}

function render() {
  const heading=[...document.querySelectorAll('h1')].find(h=>h.textContent?.includes('النماذج الإلكترونية'));
  const grid=document.querySelector('.report-grid');
  if(!heading||!grid) return;
  const section=grid.parentElement; if(!section) return;
  let bar=section.querySelector('[data-forms-manager]') as HTMLElement|null;
  if(!bar){
    bar=document.createElement('div'); bar.setAttribute('data-forms-manager','1'); bar.className='hr-module-toolbar';
    bar.innerHTML='<div class="hr-module-toolbar-title">إدارة النماذج</div><button data-saved>النماذج المحفوظة</button><button data-export>تصدير Excel</button><button data-json>نسخة JSON</button>';
    section.insertBefore(bar,grid);
    (bar.querySelector('[data-saved]') as HTMLElement).onclick=openSaved;
    (bar.querySelector('[data-export]') as HTMLElement).onclick=exportAll;
    (bar.querySelector('[data-json]') as HTMLElement).onclick=jsonBackup;
  }
  let summary=section.querySelector('[data-forms-summary]') as HTMLElement|null;
  const rows=sync();
  const byType=new Map(rows.map(r=>[r.type,r]));
  const html=TYPES.map(type=>{const r=byType.get(type); return `<div class="report" style="min-height:auto"><b>${esc(type)}</b><span>${r?`محفوظ • ${esc(fmt(r.updatedAt))}`:'غير محفوظ بعد'}</span>${r?`<button data-summary-open="${esc(type)}">فتح المحفوظ</button>`:''}</div>`}).join('');
  if(!summary){summary=document.createElement('div'); summary.setAttribute('data-forms-summary','1'); section.appendChild(summary);}
  summary.className='report-grid'; summary.innerHTML=html;
  summary.querySelectorAll('[data-summary-open]').forEach(b=>(b as HTMLElement).onclick=()=>openEditor((b as HTMLElement).dataset.summaryOpen||''));
}

const patchedKey='__hr_forms_storage_patched__';
if(!(window as any)[patchedKey]){
  const original=Storage.prototype.setItem;
  Storage.prototype.setItem=function(key:string,value:string){original.call(this,key,value);if(this===localStorage&&key===LEGACY_KEY)window.dispatchEvent(new Event('hr:forms-updated'));};
  (window as any)[patchedKey]=true;
}
window.addEventListener('hr:forms-updated',render);
new MutationObserver(()=>render()).observe(document.body,{childList:true,subtree:true});
setTimeout(render,300);
