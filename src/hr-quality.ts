import * as XLSX from 'xlsx';

type AnyRecord = Record<string, any>;

const EMP_KEY = 'hr-saudi-employees-v3';
const AUDIT_KEY = 'hr-saudi-audit-v1';

const read = <T,>(key: string, fallback: T): T => {
  try { return JSON.parse(localStorage.getItem(key) || '') as T; } catch { return fallback; }
};

const employees = (): AnyRecord[] => read<AnyRecord[]>(EMP_KEY, []);
const clean = (v: any) => String(v ?? '').trim();
const iban = (v: any) => clean(v).replace(/\s+/g, '').toUpperCase();
const validSaudiIban = (v: any) => /^SA\d{22}$/.test(iban(v));
const dateValue = (v: any) => {
  const t = Date.parse(clean(v));
  return Number.isNaN(t) ? null : new Date(t);
};
const daysTo = (v: any) => {
  const d = dateValue(v); if (!d) return null;
  return Math.ceil((d.getTime() - Date.now()) / 86400000);
};
const audit = (action: string, details: string) => {
  const rows = read<AnyRecord[]>(AUDIT_KEY, []);
  rows.unshift({ at: new Date().toISOString(), action, details });
  localStorage.setItem(AUDIT_KEY, JSON.stringify(rows.slice(0, 500)));
};

function issuesFor(e: AnyRecord, duplicateId: boolean) {
  const issues: string[] = [];
  if (!clean(e.name)) issues.push('الاسم مفقود');
  if (!clean(e.id)) issues.push('رقم الموظف مفقود');
  if (duplicateId && clean(e.id)) issues.push('رقم الموظف مكرر');
  if (!clean(e.job)) issues.push('المسمى الوظيفي مفقود');
  if (!clean(e.dept)) issues.push('الإدارة/القسم مفقود');
  if (!clean(e.identity)) issues.push('رقم الهوية/الإقامة مفقود');
  if (!clean(e.joiningDate)) issues.push('تاريخ المباشرة مفقود');
  if (!clean(e.status)) issues.push('الحالة الوظيفية مفقودة');
  if (!clean(e.bank)) issues.push('البنك مفقود');
  if (!clean(e.iban)) issues.push('IBAN مفقود');
  else if (!validSaudiIban(e.iban)) issues.push('IBAN سعودي غير صحيح');
  if (Number(e.gosiWage || 0) <= 0) issues.push('أجر الاشتراك بالتأمينات مفقود/صفر');

  const docs: [string,string][] = [['iqama','الإقامة'],['passport','الجواز'],['workPermit','رخصة/تصريح العمل'],['insurance','التأمين الطبي']];
  for (const [key,label] of docs) {
    const v = e[key];
    if (!clean(v)) { issues.push(`${label}: تاريخ غير مسجل`); continue; }
    const d = daysTo(v);
    if (d !== null && d < 0) issues.push(`${label}: منتهية`);
    else if (d !== null && d <= 30) issues.push(`${label}: تنتهي خلال ${d} يوم`);
  }
  return issues;
}

function downloadWorkbook(rows: AnyRecord[]) {
  const data = rows.map(r => ({ 'رقم الموظف': r.id, 'اسم الموظف': r.name, 'القسم': r.dept, 'المشكلة': r.issue }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), 'جودة البيانات');
  XLSX.writeFile(wb, `HR-Data-Quality-${new Date().toISOString().slice(0,10)}.xlsx`);
}

function openQuality() {
  const es = employees();
  const counts: Record<string, number> = {};
  es.forEach(e => { const id = clean(e.id); counts[id] = (counts[id] || 0) + 1; });
  const rows = es.flatMap(e => issuesFor(e, !!clean(e.id) && counts[clean(e.id)] > 1).map(issue => ({ ...e, issue })));
  const uniqueEmployees = new Set(rows.map(r => clean(r.id) || r.name)).size;
  const expired = rows.filter(r => r.issue.includes('منتهية')).length;
  const invalidIban = rows.filter(r => r.issue.includes('IBAN سعودي غير صحيح')).length;

  const overlay = document.createElement('div');
  overlay.className = 'drawer-backdrop';
  overlay.innerHTML = `<div class="modal modal-wide" style="max-width:1100px">
    <div class="modal-head"><div><h3>مركز جودة بيانات الموارد البشرية</h3><small>${es.length} موظف — ${uniqueEmployees} موظف لديهم ملاحظات</small></div><button class="btn ghost" data-close>إغلاق</button></div>
    <div class="profile-grid" style="margin:14px 0">
      <div class="info"><span>إجمالي الملاحظات</span><b>${rows.length}</b></div>
      <div class="info"><span>مستندات منتهية</span><b>${expired}</b></div>
      <div class="info"><span>IBAN غير صحيح</span><b>${invalidIban}</b></div>
      <div class="info"><span>موظفون يحتاجون مراجعة</span><b>${uniqueEmployees}</b></div>
    </div>
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px"><button class="btn primary" data-export>تصدير تقرير الجودة Excel</button><button class="btn ghost" data-refresh>إعادة الفحص</button></div>
    <div style="max-height:55vh;overflow:auto"><table class="data-table"><thead><tr><th>الموظف</th><th>القسم</th><th>المشكلة</th></tr></thead><tbody>${rows.length ? rows.map(r => `<tr><td>${clean(r.name) || '-'} <small>${clean(r.id)}</small></td><td>${clean(r.dept) || '-'}</td><td>${clean(r.issue)}</td></tr>`).join('') : '<tr><td colspan="3" style="text-align:center;padding:30px">لا توجد ملاحظات — البيانات الأساسية جيدة.</td></tr>'}</tbody></table></div>
  </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector('[data-close]')?.addEventListener('click', () => overlay.remove());
  overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
  overlay.querySelector('[data-export]')?.addEventListener('click', () => { downloadWorkbook(rows); audit('تصدير جودة البيانات', `تم تصدير ${rows.length} ملاحظة`); });
  overlay.querySelector('[data-refresh]')?.addEventListener('click', () => { overlay.remove(); openQuality(); });
}

function mount() {
  const toolbar = document.querySelector('.hr-module-toolbar');
  if (!toolbar || toolbar.querySelector('[data-quality]')) return;
  const btn = document.createElement('button');
  btn.className = 'btn ghost'; btn.dataset.quality = '1'; btn.textContent = 'جودة البيانات';
  btn.addEventListener('click', openQuality);
  toolbar.appendChild(btn);
}

const observer = new MutationObserver(mount);
observer.observe(document.body, { childList: true, subtree: true });
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mount); else mount();
