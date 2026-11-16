// 管理後台：單一用戶的額度（伺服器記錄）——顯示已用／總額度、加額度、記錄付款、設定總額度、看變動紀錄
import React, { useState } from 'react';
import { api, errorMessage } from '../utils/api';
import { db, collection, getDocs } from '../firebase';

export interface ServerCredit { uid: string; email: string | null; quota: number; used: number; updated_at?: string }
interface LedgerEntry { id: number; kind: string; files: number; amount: number | null; note: string | null; by_email: string | null; created_at: string }

type Kind = 'grant' | 'payment' | 'adjust';
const KIND_LABEL: Record<string, string> = { init: '建立帳戶', use: '翻譯', grant: '加額度', payment: '付款', adjust: '設定總額度', refund: '退回' };

interface Props {
  uid: string;
  email?: string | null;
  credit?: ServerCredit;          // 還沒有帳戶時是 undefined（第一次翻譯時才建立）
  legacyQuota?: number;           // Firestore 上的舊額度，只拿來顯示
  freeQuota: number;
  onChanged: (c: ServerCredit) => void;
}

export function CreditControls({ uid, email, credit, legacyQuota, freeQuota, onChanged }: Props) {
  const [open, setOpen] = useState(false);
  const [kind, setKind] = useState<Kind>('payment');
  const [files, setFiles] = useState('');
  const [amount, setAmount] = useState('');
  const [note, setNote] = useState('');
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);
  const [ledger, setLedger] = useState<LedgerEntry[] | null>(null);

  const loadLedger = async () => {
    const r = await api<{ ledger: LedgerEntry[] }>(`/api/admin/credits?uid=${encodeURIComponent(uid)}`);
    setLedger(r.ok ? r.data.ledger : []);
  };

  const toggle = () => {
    const next = !open;
    setOpen(next);
    setMsg(null);
    if (next && credit) loadLedger();
  };

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true);
    setMsg(null);
    try {
      // 還沒有帳戶的人：用他在 Firestore 的翻譯紀錄筆數當「已用份數」
      let seedUsed: number | undefined;
      if (!credit) {
        try { seedUsed = (await getDocs(collection(db, 'users', uid, 'history'))).size; } catch { seedUsed = 0; }
      }
      const body = kind === 'adjust'
        ? { uid, email, kind, setQuota: Number(files), note, seedUsed }
        : { uid, email, kind, files: Number(files), amount: kind === 'payment' ? Number(amount) : undefined, note, seedUsed };
      const r = await api<{ account: ServerCredit }>('/api/admin/credits', { method: 'POST', body: JSON.stringify(body) });
      if (!r.ok) throw new Error(errorMessage(r.data, '調整失敗'));
      onChanged(r.data.account);
      setMsg({ ok: true, text: '已更新' });
      setFiles(''); setAmount(''); setNote('');
      loadLedger();
    } catch (err) {
      setMsg({ ok: false, text: (err as Error).message });
    } finally {
      setBusy(false);
    }
  };

  const total = credit?.quota ?? legacyQuota ?? freeQuota;
  const used = credit?.used;
  const input = 'w-full min-w-0 text-xs bg-white border border-[var(--rule)] rounded-lg px-2 py-1.5 text-[var(--ink)] focus:outline-none focus:ring-2 focus:ring-[var(--signal)]/40';

  return (
    <div className="flex flex-col gap-1">
      <span className="text-[10px] uppercase tracking-wider text-[var(--muted)] font-bold">額度</span>
      <div className="flex items-center gap-2">
        <span className="text-xs font-bold text-[var(--ink)] whitespace-nowrap">
          {credit ? <>已用 {used} / 共 {total} 份</> : <span className="text-[var(--muted)] font-medium">共 {total} 份（還沒翻過）</span>}
        </span>
        <button type="button" onClick={toggle}
          className="px-2.5 py-1 rounded-lg text-xs font-bold border border-[var(--rule)] bg-white hover:border-[var(--signal)] text-[var(--signal-ink)]">
          {open ? '收起' : '調整'}
        </button>
      </div>

      {open && (
        <div className="mt-2 p-3 rounded-xl border border-[var(--rule)] bg-[var(--paper)] w-full min-w-[260px] max-w-[420px]">
          <form onSubmit={submit} className="flex flex-col gap-2">
            <div className="flex gap-1" role="radiogroup" aria-label="調整方式">
              {(['payment', 'grant', 'adjust'] as Kind[]).map(k => (
                <button key={k} type="button" role="radio" aria-checked={kind === k} onClick={() => setKind(k)}
                  className={`flex-1 px-2 py-1 rounded-lg text-xs font-bold border ${kind === k ? 'bg-[var(--signal)] text-white border-[var(--signal)]' : 'bg-white border-[var(--rule)] text-[var(--muted)]'}`}>
                  {k === 'payment' ? '記錄付款' : k === 'grant' ? '加額度' : '設定總額度'}
                </button>
              ))}
            </div>
            <div className="flex gap-2">
              <label className="flex-1 min-w-0 flex flex-col gap-0.5 text-[11px] text-[var(--muted)]">
                {kind === 'adjust' ? '總額度改成（份）' : '加幾份（扣回填負數）'}
                <input type="number" inputMode="numeric" required value={files} onChange={e => setFiles(e.target.value)} className={input} />
              </label>
              {kind === 'payment' && (
                <label className="flex-1 min-w-0 flex flex-col gap-0.5 text-[11px] text-[var(--muted)]">
                  收款金額（NT$）
                  <input type="number" inputMode="numeric" required min={1} value={amount} onChange={e => setAmount(e.target.value)} className={input} />
                </label>
              )}
            </div>
            <label className="flex flex-col gap-0.5 text-[11px] text-[var(--muted)]">
              備註（例如：11/16 轉帳末五碼 12345）
              <input type="text" maxLength={300} value={note} onChange={e => setNote(e.target.value)} className={input} />
            </label>
            <div className="flex items-center gap-2">
              <button type="submit" disabled={busy}
                className="px-3 py-1.5 rounded-lg text-xs font-bold bg-[var(--ink)] text-white disabled:opacity-50">
                {busy ? '處理中…' : '確定'}
              </button>
              {msg && <span className={`text-xs font-medium ${msg.ok ? 'text-emerald-700' : 'text-red-600'}`}>{msg.text}</span>}
            </div>
          </form>

          {ledger && (
            <div className="mt-3 border-t border-[var(--rule)] pt-2">
              <p className="text-[11px] font-bold text-[var(--muted)] mb-1">變動紀錄</p>
              {ledger.length === 0 ? <p className="text-[11px] text-[var(--muted)]">還沒有紀錄</p> : (
                <ul className="flex flex-col gap-1 max-h-40 overflow-y-auto">
                  {ledger.map(l => (
                    <li key={l.id} className="text-[11px] text-[var(--ink)] flex gap-2">
                      <span className="text-[var(--muted)] whitespace-nowrap">{new Date(l.created_at).toLocaleDateString('zh-TW', { month: 'numeric', day: 'numeric' })}</span>
                      <span className="font-bold whitespace-nowrap">{KIND_LABEL[l.kind] || l.kind}</span>
                      <span className="whitespace-nowrap">{l.files > 0 ? `+${l.files}` : l.files} 份</span>
                      {l.amount ? <span className="whitespace-nowrap">NT${l.amount.toLocaleString('zh-TW')}</span> : null}
                      <span className="truncate text-[var(--muted)]" title={l.note || ''}>{l.note}</span>
                    </li>
                  ))}
                </ul>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
