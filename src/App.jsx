import { useState, useEffect } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

/* ============ Supabase ============ */
const SUPABASE_URL = "https://cnhntxoxrvjajeyvibnu.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImNuaG50eG94cnZqYWpleXZpYm51Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODMxNTE4NzUsImV4cCI6MjA5ODcyNzg3NX0.gskPbydZS8zsuE8GdCpYR_RVS0Ta1104HFImDjVP-LQ";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

/* ============ GAS連携 ============ */
const GAS_URL = "https://script.google.com/macros/s/AKfycby4juYNfxT3ob39J5RqtrXOSbcLKZ-H-9dowoe4_v0UcbRz4obDXraDz8mvtcXRkc3C/exec";

/**
 * GAS にメール送信リクエストを送る（補助機能）。
 * no-cors モードで送信し、エラーが起きても予約処理には影響させない。
 */
async function sendToGAS(payload) {
  try {
    await fetch(GAS_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain" },
      body: JSON.stringify(payload),
    });
  } catch (e) {
    console.warn("GAS通知の送信に失敗しましたが、予約処理は正常に完了しています。", e);
  }
}

/* ============ 基本設定 ============ */
const FARM = {
  name: "やんばるいちご園",
  email: "yanbaru.1066@gmail.com",
  tel: "080-7806-1077",
  mapUrl: "https://www.google.com/maps/place/%E3%82%84%E3%82%93%E3%81%B0%E3%82%8B%E3%81%84%E3%81%A1%E3%81%94%E5%9C%92/@26.4844974,127.9817773,17z/data=!3m1!4b1!4m6!3m5!1s0x34e501daa8c7b7f7:0xa8d5085d9fda8e06!8m2!3d26.4844974!4d127.9817773!16s%2Fg%2F11fqg74ds2?hl=ja",
  instagram: "https://www.instagram.com/yanbaruichigoen/",
};
const SLOTS = ["10:00", "11:00", "12:00", "13:00", "13:30", "14:00", "15:00", "16:00", "17:00", "18:00", "19:00", "20:00"];
const INITIAL_ACTIVE_SLOTS = ["10:00", "11:00", "12:00", "13:30", "14:00"];
const CATS = [
  { key: "adult", label: "大人（中学生以上）", price: 2500 },
  { key: "elem", label: "小学生", price: 2000 },
  { key: "senior", label: "シニア（75歳以上）", price: 2000 },
  { key: "child", label: "幼児（2歳〜6歳）", price: 1500 },
];
const TERMS = [
  "40分間の時間の中で食べ放題です",
  "転倒防止のためサンダルやハイヒールでのご来園はご遠慮ください",
  "いちごの生育状況などにより実施できない場合がありますのでご了承ください",
];
const DEFAULT_SETTINGS = {
  activeSlots: [...INITIAL_ACTIVE_SLOTS],
  defaultCapacity: 20,
  defaultDayCapacity: "",
  monthSettings: {},
  dateSettings: {},
  adminPassword: "1066",
};
const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

/* ============ Supabase ストレージ ============ */

/* DB行 → アプリ内オブジェクトへの変換（snake_case → camelCase） */
function rowToReservation(row) {
  return {
    id: row.id,
    date: row.date,
    slot: row.slot,
    counts: row.counts,
    total: row.total,
    people: row.people,
    name: row.name,
    phone: row.phone,
    email: row.email,
    note: row.note || "",
    status: row.status,
    checkedIn: row.checked_in,
    createdAt: row.created_at,
  };
}

/* アプリ内オブジェクト → DB行への変換（camelCase → snake_case） */
function reservationToRow(r) {
  return {
    id: r.id,
    date: r.date,
    slot: r.slot,
    counts: r.counts,
    total: r.total,
    people: r.people,
    name: r.name,
    phone: r.phone,
    email: r.email,
    note: r.note || "",
    status: r.status,
    checked_in: r.checkedIn ?? false,
    created_at: r.createdAt,
  };
}

async function fetchReservations() {
  const { data, error } = await supabase.from("reservations").select("*");
  if (error) {
    console.error("reservations fetch error", error);
    return [];
  }
  return (data || []).map(rowToReservation);
}

async function insertReservation(r) {
  const { error } = await supabase.from("reservations").insert(reservationToRow(r));
  if (error) {
    console.error("reservation insert error", error);
    return false;
  }
  return true;
}

async function updateReservation(id, fields) {
  // fields は DB カラム名（snake_case）で渡す
  const { error } = await supabase.from("reservations").update(fields).eq("id", id);
  if (error) {
    console.error("reservation update error", error);
    return false;
  }
  return true;
}

async function fetchSettings() {
  const { data, error } = await supabase
    .from("settings")
    .select("*")
    .eq("id", "main")
    .single();
  if (error || !data) {
    // 設定が未登録のときはデフォルト値を返す
    return { ...DEFAULT_SETTINGS };
  }
  return { ...DEFAULT_SETTINGS, ...data.data };
}

async function upsertSettings(settingsData) {
  const { error } = await supabase
    .from("settings")
    .upsert({ id: "main", data: settingsData });
  if (error) {
    console.error("settings upsert error", error);
    return false;
  }
  return true;
}

/* ============ 日付ヘルパー ============ */
function fmtDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${dd}`;
}
function todayStr() {
  return fmtDate(new Date());
}
function jpDate(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-").map(Number);
  const w = WEEKDAYS[new Date(y, m - 1, d).getDay()];
  return `${y}年${m}月${d}日（${w}）`;
}
function yen(n) {
  return n.toLocaleString("ja-JP") + "円";
}

/* ============ 予約計算 ============ */
function calcTotal(counts) {
  return CATS.reduce((s, c) => s + (counts[c.key] || 0) * c.price, 0);
}
function calcPeople(counts) {
  return CATS.reduce((s, c) => s + (counts[c.key] || 0), 0);
}
function breakdownLines(counts) {
  return CATS.filter((c) => (counts[c.key] || 0) > 0).map(
    (c) => `${c.label} × ${counts[c.key]}名 … ${yen(counts[c.key] * c.price)}`
  );
}

/* 設定の優先度：日別 ＞ 月別 ＞ 基本 */
function numOr(v, fallback) {
  return v === "" || v == null ? fallback : Number(v);
}
function wdOf(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d).getDay();
}
function resolveDay(dateIso, settings) {
  const ds = settings.dateSettings[dateIso];
  const ms = settings.monthSettings?.[dateIso.slice(0, 7)];
  const weekdayClosed = !!ms?.closedWeekdays?.includes(wdOf(dateIso));
  const dayClosed = !!ds?.closedAll || (weekdayClosed && !ds?.openOverride);
  const dayCapacity = numOr(ds?.dayCapacity, numOr(ms?.dayCapacity, numOr(settings.defaultDayCapacity, null)));
  const slotDefault = numOr(ms?.slotCapacity, settings.defaultCapacity);
  return { ds, dayClosed, weekdayClosed, dayCapacity, slotDefault };
}
function slotStatus(dateIso, slot, reservations, settings) {
  const { ds, dayClosed, dayCapacity, slotDefault } = resolveDay(dateIso, settings);
  const slotSetting = ds?.slots?.[slot];
  const mode = slotSetting?.mode || (slotSetting?.closed ? "closed" : "default");
  const defaultOpen = !settings.activeSlots || settings.activeSlots.includes(slot);
  const slotClosed = mode === "closed" || (mode === "default" && !defaultOpen);
  const closed = dayClosed || slotClosed;
  const capacity = numOr(slotSetting?.capacity, slotDefault);
  const active = reservations.filter((r) => r.status === "active" && r.date === dateIso);
  const booked = active.filter((r) => r.slot === slot).reduce((s, r) => s + r.people, 0);
  const dayBooked = active.reduce((s, r) => s + r.people, 0);
  let remaining = Math.max(0, capacity - booked);
  if (dayCapacity != null) remaining = Math.min(remaining, Math.max(0, dayCapacity - dayBooked));
  return { closed, capacity, booked, remaining, dayCapacity, dayBooked };
}
function dayAvailability(dateIso, reservations, settings) {
  let anyOpen = false;
  let totalRemaining = 0;
  for (const slot of SLOTS) {
    const st = slotStatus(dateIso, slot, reservations, settings);
    if (!st.closed && st.remaining > 0) {
      anyOpen = true;
      totalRemaining += st.remaining;
    }
  }
  return { anyOpen, totalRemaining };
}
function genId() {
  const s = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let id = "";
  for (let i = 0; i < 6; i++) id += s[Math.floor(Math.random() * s.length)];
  return "YIG-" + id;
}

/* ============ メール文面 ============ */
function contactBlock() {
  return `${FARM.name}
メール：${FARM.email}
電話：${FARM.tel}
📍 Googleマップ: ${FARM.mapUrl}
📷 Instagram: ${FARM.instagram}`;
}
function customerMail(r) {
  return `${FARM.name}へのご予約、ありがとうございました。
下記にてご予約を承りました。

予約番号：${r.id}
予約名：${r.name} 様
予約日時：${jpDate(r.date)} ${r.slot}
予約内訳：
${breakdownLines(r.counts).map((l) => "　" + l).join("\n")}
　合計：${r.people}名　${yen(r.total)}

キャンセルの場合は下記よりお願いいたします。
キャンセルURL: https://yanbaru-ichigo.example.com/cancel/${r.id}

${contactBlock()}`;
}
function adminNewMail(r) {
  return `【新規予約】${FARM.name} 予約システム

予約番号：${r.id}
予約日時：${jpDate(r.date)} ${r.slot}
予約名：${r.name} 様
電話番号：${r.phone}
メール：${r.email}
内訳：
${breakdownLines(r.counts).map((l) => "　" + l).join("\n")}
合計：${r.people}名　${yen(r.total)}
備考：${r.note || "（なし）"}`;
}
function customerCancelMail(r) {
  return `${FARM.name}です。
下記のご予約のキャンセルを承りました。

予約番号：${r.id}
予約名：${r.name} 様
予約日時：${jpDate(r.date)} ${r.slot}

またのご予約を心よりお待ちしております。

${contactBlock()}`;
}
function adminCancelMail(r) {
  return `【キャンセル】${FARM.name} 予約システム

予約番号：${r.id}
予約日時：${jpDate(r.date)} ${r.slot}
予約名：${r.name} 様（${r.phone}）
人数：${r.people}名　${yen(r.total)}`;
}

/* ============ 管理者キャンセル通知メール ============ */
function adminCancelNotifyMail(r, reason) {
  return `まことに申し訳ございませんが、${reason}によりこのご予約はキャンセルさせていただきました。

このキャンセルに身に覚えがない場合はお手数ですが、やんばるいちご園までご連絡ください。

やんばるいちご園
電話: ${FARM.tel}
メール: ${FARM.email}`;
}

/* ============ 共通UI ============ */
function Stepper({ value, onChange }) {
  return (
    <div className="stepper">
      <button type="button" className="step-btn" onClick={() => onChange(Math.max(0, value - 1))} aria-label="減らす">−</button>
      <span className="step-num">{value}</span>
      <button type="button" className="step-btn" onClick={() => onChange(Math.min(30, value + 1))} aria-label="増やす">＋</button>
    </div>
  );
}

function linkify(text) {
  const parts = text.split(/(\[[^\]]+\]\(https?:\/\/[^\s)]+\)|https?:\/\/[^\s]+)/g);
  return parts.map((p, i) => {
    const labeled = p.match(/^\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)$/);
    if (labeled) {
      return (
        <a key={i} href={labeled[2]} target="_blank" rel="noreferrer" className="mail-link">{labeled[1]}</a>
      );
    }
    if (/^https?:\/\//.test(p)) {
      return (
        <a key={i} href={p} target="_blank" rel="noreferrer" className="mail-link">{p}</a>
      );
    }
    return p;
  });
}

function MailPreview({ title, to, body }) {
  return (
    <div className="mail-box">
      <div className="mail-head">
        <span className="mail-tag">メール送信内容</span>
        <div className="mail-title">{title}</div>
        <div className="mail-to">宛先：{to}</div>
      </div>
      <pre className="mail-body">{linkify(body)}</pre>
    </div>
  );
}

/* カレンダー */
function Calendar({ ym, setYm, selected, onSelect, reservations, settings }) {
  const [year, month] = ym;
  const first = new Date(year, month, 1);
  const startDow = first.getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const today = todayStr();
  const cells = [];
  for (let i = 0; i < startDow; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);

  const prev = () => setYm(month === 0 ? [year - 1, 11] : [year, month - 1]);
  const next = () => setYm(month === 11 ? [year + 1, 0] : [year, month + 1]);

  return (
    <div className="calendar">
      <div className="cal-nav">
        <button type="button" className="cal-arrow" onClick={prev}>◀</button>
        <div className="cal-title">{year}年 {month + 1}月</div>
        <button type="button" className="cal-arrow" onClick={next}>▶</button>
      </div>
      <div className="cal-grid cal-week">
        {WEEKDAYS.map((w, i) => (
          <div key={w} className={"cal-wd" + (i === 0 ? " sun" : i === 6 ? " sat" : "")}>{w}</div>
        ))}
      </div>
      <div className="cal-grid">
        {cells.map((d, i) => {
          if (d === null) return <div key={"e" + i} />;
          const iso = fmtDate(new Date(year, month, d));
          const isPast = iso < today;
          const av = dayAvailability(iso, reservations, settings);
          const disabled = isPast || !av.anyOpen;
          const mark = isPast ? "" : av.anyOpen ? (av.totalRemaining <= 10 ? "△" : "○") : "×";
          return (
            <button
              type="button"
              key={iso}
              disabled={disabled}
              onClick={() => onSelect(iso)}
              className={
                "cal-day" +
                (selected === iso ? " selected" : "") +
                (disabled ? " disabled" : "") +
                (iso === today ? " today" : "")
              }
            >
              <span className="cal-num">{d}</span>
              <span className={"cal-mark" + (mark === "×" ? " full" : mark === "△" ? " few" : "")}>{mark}</span>
            </button>
          );
        })}
      </div>
      <div className="cal-legend">○ 空きあり　△ 残りわずか　× 予約不可・満員</div>
    </div>
  );
}

/* ============ 予約フロー ============ */
function BookingApp({ reservations, settings, refresh, goHome }) {
  const [step, setStep] = useState(1);
  const [ym, setYm] = useState(() => { const n = new Date(); return [n.getFullYear(), n.getMonth()]; });
  const [date, setDate] = useState("");
  const [slot, setSlot] = useState("");
  const [counts, setCounts] = useState({ adult: 0, elem: 0, senior: 0, child: 0 });
  const [name, setName] = useState("");
  const [phone, setPhone] = useState("");
  const [email, setEmail] = useState("");
  const [note, setNote] = useState("");
  const [agreed, setAgreed] = useState(false);
  const [error, setError] = useState("");
  const [sending, setSending] = useState(false);
  const [done, setDone] = useState(null);

  const total = calcTotal(counts);
  const people = calcPeople(counts);
  const st = date && slot ? slotStatus(date, slot, reservations, settings) : null;

  const submit = async () => {
    setSending(true);
    setError("");
    // 最新データで空き再確認
    const latest = await fetchReservations();
    const check = slotStatus(date, slot, latest, settings);
    if (check.closed || check.remaining < people) {
      setError("申し訳ありません。ただいまこの時間帯の空きが不足しています。別の日時をお選びください。");
      setSending(false);
      setStep(1);
      await refresh();
      return;
    }
    const r = {
      id: genId(),
      date, slot, counts, total, people,
      name: name.trim(), phone: phone.trim(), email: email.trim(),
      note: note.trim(),
      status: "active",
      checkedIn: false,
      createdAt: new Date().toISOString(),
    };
    const ok = await insertReservation(r);
    if (!ok) {
      setError("保存に失敗しました。時間をおいて再度お試しください。");
      setSending(false);
      return;
    }
    await refresh();
    // GAS にメール送信リクエストを送る（予約完了通知）
    sendToGAS({
      type: "booking",
      id: r.id,
      date: r.date,
      slot: r.slot,
      name: r.name,
      phone: r.phone,
      email: r.email,
      counts: r.counts,
      total: r.total,
      people: r.people,
      note: r.note,
      createdAt: r.createdAt,
    });
    setDone(r);
    setSending(false);
  };

  if (done) {
    return (
      <div className="page">
        <div className="card center-card">
          <div className="big-icon">🍓</div>
          <h2 className="done-title">ご予約を承りました</h2>
          <p className="done-sub">予約番号：<b className="rid">{done.id}</b></p>
          <p className="done-note">キャンセルの際に必要になりますので、予約番号をお控えください。</p>
        </div>
        <button type="button" className="btn primary wide" onClick={goHome}>トップへ戻る</button>
      </div>
    );
  }

  return (
    <div className="page">
      <div className="steps-bar">
        {["日時", "人数", "お客様情報", "確認"].map((s, i) => (
          <div key={s} className={"steps-item" + (step === i + 1 ? " on" : step > i + 1 ? " past" : "")}>
            <span className="steps-dot">{i + 1}</span>{s}
          </div>
        ))}
      </div>
      {error && <div className="alert">{error}</div>}

      {step === 1 && (
        <div className="card">
          <h2 className="sec-title">ご希望の日付を選択</h2>
          <Calendar ym={ym} setYm={setYm} selected={date} onSelect={(d) => { setDate(d); setSlot(""); }} reservations={reservations} settings={settings} />
          {date && (
            <>
              <h2 className="sec-title">{jpDate(date)} の時間を選択</h2>
              <div className="slot-list">
                {SLOTS.filter((s) => !slotStatus(date, s, reservations, settings).closed).map((s) => {
                  const ss = slotStatus(date, s, reservations, settings);
                  const dis = ss.remaining <= 0;
                  return (
                    <button type="button" key={s} disabled={dis}
                      className={"slot-btn" + (slot === s ? " selected" : "") + (dis ? " disabled" : "")}
                      onClick={() => setSlot(s)}>
                      <span className="slot-time">{s}</span>
                      <span className="slot-info">{dis ? "満員" : `残り${ss.remaining}名`}</span>
                    </button>
                  );
                })}
              </div>
            </>
          )}
          <div className="btn-row">
            <button type="button" className="btn ghost" onClick={goHome}>戻る</button>
            <button type="button" className="btn primary" disabled={!date || !slot} onClick={() => setStep(2)}>次へ</button>
          </div>
        </div>
      )}

      {step === 2 && (
        <div className="card">
          <h2 className="sec-title">人数の入力</h2>
          <p className="muted">{jpDate(date)} {slot}（残り {st?.remaining}名）</p>
          {CATS.map((c) => (
            <div key={c.key} className="cat-row">
              <div>
                <div className="cat-label">{c.label}</div>
                <div className="cat-price">{yen(c.price)}</div>
              </div>
              <Stepper value={counts[c.key]} onChange={(v) => setCounts({ ...counts, [c.key]: v })} />
            </div>
          ))}
          <div className="total-box">
            <span>合計 {people}名</span>
            <span className="total-yen">{yen(total)}</span>
          </div>
          {st && people > st.remaining && (
            <div className="alert">この時間帯の残り枠（{st.remaining}名）を超えています。人数を調整するか、別の時間をお選びください。</div>
          )}
          <div className="btn-row">
            <button type="button" className="btn ghost" onClick={() => setStep(1)}>戻る</button>
            <button type="button" className="btn primary" disabled={people === 0 || (st && people > st.remaining)} onClick={() => setStep(3)}>次へ</button>
          </div>
        </div>
      )}

      {step === 3 && (
        <div className="card">
          <h2 className="sec-title">お客様情報の入力</h2>
          <label className="fld">お名前<span className="req">必須</span>
            <input value={name} onChange={(e) => setName(e.target.value)} placeholder="例）山田 花子" />
          </label>
          <label className="fld">電話番号<span className="req">必須</span>
            <input value={phone} onChange={(e) => setPhone(e.target.value)} inputMode="tel" placeholder="例）090-1234-5678" />
          </label>
          <label className="fld">メールアドレス<span className="req">必須</span>
            <input value={email} onChange={(e) => setEmail(e.target.value)} inputMode="email" placeholder="例）hanako@example.com" />
          </label>
          <label className="fld">備考（自由入力）
            <textarea value={note} onChange={(e) => setNote(e.target.value)} rows={3} placeholder="車いすでの来園、アレルギーのご相談など" />
          </label>
          <div className="btn-row">
            <button type="button" className="btn ghost" onClick={() => setStep(2)}>戻る</button>
            <button type="button" className="btn primary"
              disabled={!name.trim() || !phone.trim() || !email.trim() || !email.includes("@")}
              onClick={() => setStep(4)}>入力内容を確認</button>
          </div>
        </div>
      )}

      {step === 4 && (
        <div className="card">
          <h2 className="sec-title">ご予約内容の確認</h2>
          <table className="conf-table"><tbody>
            <tr><th>予約日時</th><td>{jpDate(date)} {slot}〜</td></tr>
            <tr><th>人数・料金</th><td>{breakdownLines(counts).map((l) => <div key={l}>{l}</div>)}<div className="conf-total">合計 {people}名　{yen(total)}</div></td></tr>
            <tr><th>お名前</th><td>{name} 様</td></tr>
            <tr><th>電話番号</th><td>{phone}</td></tr>
            <tr><th>メール</th><td>{email}</td></tr>
            <tr><th>備考</th><td>{note || "（なし）"}</td></tr>
          </tbody></table>
          <div className="terms">
            <div className="terms-title">ご来園にあたって</div>
            <ul>{TERMS.map((t) => <li key={t}>{t}</li>)}</ul>
            <label className="agree">
              <input type="checkbox" checked={agreed} onChange={(e) => setAgreed(e.target.checked)} />
              上記の内容に同意します
            </label>
          </div>
          <div className="btn-row">
            <button type="button" className="btn ghost" onClick={() => setStep(3)}>戻る</button>
            <button type="button" className="btn primary" disabled={!agreed || sending} onClick={submit}>
              {sending ? "送信中…" : "同意して予約を送信"}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============ キャンセル ============ */
function CancelApp({ settings, refresh, goHome }) {
  const [rid, setRid] = useState("");
  const [phone, setPhone] = useState("");
  const [target, setTarget] = useState(null);
  const [error, setError] = useState("");
  const [confirm, setConfirm] = useState(false);
  const [cancelled, setCancelled] = useState(null);
  const [busy, setBusy] = useState(false);

  const search = async () => {
    setError("");
    const latest = await fetchReservations();
    const r = latest.find(
      (x) => x.id.toUpperCase() === rid.trim().toUpperCase() &&
        x.phone.replace(/[-\s]/g, "") === phone.trim().replace(/[-\s]/g, "")
    );
    if (!r) { setError("予約が見つかりませんでした。予約番号と電話番号をご確認ください。"); return; }
    if (r.status === "cancelled") { setError("この予約はすでにキャンセルされています。"); return; }
    setTarget(r);
  };

  const doCancel = async () => {
    setBusy(true);
    const ok = await updateReservation(target.id, {
      status: "cancelled",
    });
    if (ok) {
      await refresh();
      // GAS にキャンセル通知を送る（お客様キャンセル）
      sendToGAS({
        type: "cancel",
        id: target.id,
        date: target.date,
        slot: target.slot,
        name: target.name,
        phone: target.phone,
        email: target.email,
        counts: target.counts,
        total: target.total,
        people: target.people,
        note: target.note,
      });
      setCancelled({ ...target, status: "cancelled" });
    }
    setConfirm(false);
    setBusy(false);
  };

  if (cancelled) {
    return (
      <div className="page">
        <div className="card center-card">
          <div className="big-icon">🌱</div>
          <h2 className="done-title">予約はキャンセルされました</h2>
          <p className="done-sub">またのご予約お待ちしております。</p>
        </div>
        <button type="button" className="btn primary wide" onClick={goHome}>トップへ戻る</button>
      </div>
    );
  }

  return (
    <div className="page">
      {!target ? (
        <div className="card">
          <h2 className="sec-title">ご予約のキャンセル</h2>
          <p className="muted">確認メールに記載の予約番号と、ご予約時の電話番号を入力してください。</p>
          {error && <div className="alert">{error}</div>}
          <label className="fld">予約番号
            <input value={rid} onChange={(e) => setRid(e.target.value)} placeholder="例）YIG-ABC123" />
          </label>
          <label className="fld">電話番号
            <input value={phone} onChange={(e) => setPhone(e.target.value)} inputMode="tel" placeholder="ご予約時の電話番号" />
          </label>
          <div className="btn-row">
            <button type="button" className="btn ghost" onClick={goHome}>戻る</button>
            <button type="button" className="btn primary" disabled={!rid.trim() || !phone.trim()} onClick={search}>予約を照会する</button>
          </div>
        </div>
      ) : (
        <div className="card">
          <h2 className="sec-title">ご予約内容</h2>
          <table className="conf-table"><tbody>
            <tr><th>予約番号</th><td>{target.id}</td></tr>
            <tr><th>予約日時</th><td>{jpDate(target.date)} {target.slot}〜</td></tr>
            <tr><th>お名前</th><td>{target.name} 様</td></tr>
            <tr><th>人数・料金</th><td>{breakdownLines(target.counts).map((l) => <div key={l}>{l}</div>)}<div className="conf-total">合計 {target.people}名　{yen(target.total)}</div></td></tr>
          </tbody></table>
          <div className="btn-row">
            <button type="button" className="btn ghost" onClick={() => setTarget(null)}>戻る</button>
            <button type="button" className="btn danger" onClick={() => setConfirm(true)}>予約をキャンセルする</button>
          </div>
        </div>
      )}
      {confirm && (
        <div className="modal-bg">
          <div className="modal">
            <p className="modal-msg">キャンセルしてよろしいですか？</p>
            <div className="btn-row modal-btns">
              <button type="button" className="btn ghost" onClick={() => setConfirm(false)}>いいえ</button>
              <button type="button" className="btn danger" disabled={busy} onClick={doCancel}>{busy ? "処理中…" : "はい"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============ 管理者アプリ ============ */
function AdminApp({ reservations, settings, refresh, saveSettings, goHome }) {
  const [authed, setAuthed] = useState(false);
  const [pw, setPw] = useState("");
  const [pwErr, setPwErr] = useState("");
  const [tab, setTab] = useState("dash");

  if (!authed) {
    return (
      <div className="page">
        <div className="card center-card">
          <h2 className="sec-title">管理者ログイン</h2>
          {pwErr && <div className="alert">{pwErr}</div>}
          <label className="fld">パスワード
            <input type="password" value={pw} onChange={(e) => setPw(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter" && pw === settings.adminPassword) setAuthed(true); }} />
          </label>
          <p className="muted small">初期パスワード：1066（設定タブで変更できます）</p>
          <div className="btn-row">
            <button type="button" className="btn ghost" onClick={goHome}>戻る</button>
            <button type="button" className="btn primary" onClick={() => {
              if (pw === settings.adminPassword) { setAuthed(true); setPwErr(""); }
              else setPwErr("パスワードが違います");
            }}>ログイン</button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="page">
      <div className="admin-tabs">
        {[["dash", "予約状況"], ["checkin", "本日の受付"], ["settings", "設定"]].map(([k, l]) => (
          <button type="button" key={k} className={"tab-btn" + (tab === k ? " on" : "")} onClick={() => setTab(k)}>{l}</button>
        ))}
        <button type="button" className="tab-btn exit" onClick={goHome}>閉じる</button>
      </div>
      {tab === "dash" && <AdminDash reservations={reservations} refresh={refresh} settings={settings} />}
      {tab === "checkin" && <AdminCheckin reservations={reservations} refresh={refresh} />}
      {tab === "settings" && <AdminSettings settings={settings} saveSettings={saveSettings} reservations={reservations} />}
    </div>
  );
}

/* --- 予約状況（月別・日別・時間帯別 + Excel出力） --- */
function AdminDash({ reservations, refresh, settings }) {
  const n = new Date();
  const [ym, setYm] = useState([n.getFullYear(), n.getMonth()]);
  const [selDate, setSelDate] = useState("");
  const [year, month] = ym;

  /* --- 管理者キャンセル --- */
  const [cancelTarget, setCancelTarget] = useState(null);
  const [cancelReason, setCancelReason] = useState("");
  const [cancelStep, setCancelStep] = useState("reason"); // "reason" | "confirm" | "password"
  const [cancelPw, setCancelPw] = useState("");
  const [cancelPwErr, setCancelPwErr] = useState("");
  const [cancelBusy, setCancelBusy] = useState(false);
  const [cancelDone, setCancelDone] = useState(null);

  const openCancel = (r) => {
    setCancelTarget(r);
    setCancelReason("");
    setCancelStep("reason");
    setCancelPw("");
    setCancelPwErr("");
    setCancelDone(null);
  };
  const closeCancel = () => {
    setCancelTarget(null);
    setCancelDone(null);
  };
  const proceedToConfirm = () => setCancelStep("confirm");
  const proceedToPassword = () => { setCancelStep("password"); setCancelPw(""); setCancelPwErr(""); };
  const execCancel = async () => {
    if (cancelPw !== (settings.adminPassword || "1066")) {
      setCancelPwErr("パスワードが違います");
      return;
    }
    setCancelBusy(true);
    const noteValue = cancelTarget.note
      ? `${cancelTarget.note}｜管理者キャンセル: ${cancelReason}`
      : `管理者キャンセル: ${cancelReason}`;
    const ok = await updateReservation(cancelTarget.id, {
      status: "cancelled",
      note: noteValue,
    });
    if (ok) {
      await refresh();
      // GAS に管理者キャンセル通知を送る
      sendToGAS({
        type: "admin_cancel",
        id: cancelTarget.id,
        date: cancelTarget.date,
        slot: cancelTarget.slot,
        name: cancelTarget.name,
        phone: cancelTarget.phone,
        email: cancelTarget.email,
        counts: cancelTarget.counts,
        total: cancelTarget.total,
        people: cancelTarget.people,
        note: noteValue,
        reason: cancelReason,
      });
      setCancelDone({ ...cancelTarget, status: "cancelled", note: noteValue });
    }
    setCancelBusy(false);
  };
  const ymStr = `${year}-${String(month + 1).padStart(2, "0")}`;

  const monthRes = reservations.filter((r) => r.date.startsWith(ymStr));
  const active = monthRes.filter((r) => r.status === "active");
  const mPeople = active.reduce((s, r) => s + r.people, 0);
  const mSales = active.reduce((s, r) => s + r.total, 0);

  const byDate = {};
  for (const r of active) {
    byDate[r.date] = byDate[r.date] || { people: 0, sales: 0, count: 0 };
    byDate[r.date].people += r.people;
    byDate[r.date].sales += r.total;
    byDate[r.date].count += 1;
  }
  const dates = Object.keys(byDate).sort();

  const exportExcel = () => {
    const rows = monthRes
      .sort((a, b) => (a.date + a.slot).localeCompare(b.date + b.slot))
      .map((r) => ({
        "予約番号": r.id, "予約日": r.date, "時間": r.slot,
        "状態": r.status === "active" ? "予約中" : "キャンセル",
        "お名前": r.name, "電話番号": r.phone, "メール": r.email,
        "大人": r.counts.adult || 0, "小学生": r.counts.elem || 0,
        "シニア": r.counts.senior || 0, "幼児": r.counts.child || 0,
        "合計人数": r.people, "料金合計": r.total,
        "来園チェック": r.checkedIn ? "来園済" : "",
        "備考": r.note || "", "予約受付日時": r.createdAt,
      }));
    const daily = dates.map((d) => ({
      "日付": d, "予約件数": byDate[d].count, "合計人数": byDate[d].people, "売上": byDate[d].sales,
    }));
    daily.push({ "日付": "月合計", "予約件数": active.length, "合計人数": mPeople, "売上": mSales });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{ "予約番号": "データなし" }]), "予約一覧");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(daily), "日別集計");
    XLSX.writeFile(wb, `やんばるいちご園_予約_${ymStr}.xlsx`);
  };

  return (
    <>
      <div className="card">
        <div className="cal-nav">
          <button type="button" className="cal-arrow" onClick={() => { setYm(month === 0 ? [year - 1, 11] : [year, month - 1]); setSelDate(""); }}>◀</button>
          <div className="cal-title">{year}年 {month + 1}月の予約状況</div>
          <button type="button" className="cal-arrow" onClick={() => { setYm(month === 11 ? [year + 1, 0] : [year, month + 1]); setSelDate(""); }}>▶</button>
        </div>
        <div className="stat-row">
          <div className="stat"><div className="stat-label">予約件数</div><div className="stat-val">{active.length}<small>件</small></div></div>
          <div className="stat"><div className="stat-label">合計人数</div><div className="stat-val">{mPeople}<small>名</small></div></div>
          <div className="stat"><div className="stat-label">売上見込</div><div className="stat-val">{mSales.toLocaleString()}<small>円</small></div></div>
        </div>
        <button type="button" className="btn green wide" onClick={exportExcel}>この月をExcel出力（.xlsx）</button>
      </div>

      <div className="card">
        <h2 className="sec-title">日別の予約状況</h2>
        {dates.length === 0 ? (
          <p className="muted">この月の予約はまだありません。</p>
        ) : (
          <table className="list-table">
            <thead><tr><th>日付</th><th>件数</th><th>人数</th><th>売上</th></tr></thead>
            <tbody>
              {dates.map((d) => (
                <tr key={d} className={"clickable" + (selDate === d ? " sel" : "")} onClick={() => setSelDate(selDate === d ? "" : d)}>
                  <td>{jpDate(d)}</td><td>{byDate[d].count}件</td><td>{byDate[d].people}名</td><td>{yen(byDate[d].sales)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        <p className="muted small">日付をタップすると時間帯別の内訳が表示されます。</p>
      </div>

      {selDate && (
        <div className="card">
          <h2 className="sec-title">{jpDate(selDate)}｜時間帯別</h2>
          {SLOTS.map((s) => {
            const list = reservations.filter((r) => r.date === selDate && r.slot === s);
            const act = list.filter((r) => r.status === "active");
            if (list.length === 0) return null;
            return (
              <div key={s} className="slot-block">
                <div className="slot-block-head">
                  <b>{s}〜</b>
                  <span>{act.reduce((x, r) => x + r.people, 0)}名 ／ {yen(act.reduce((x, r) => x + r.total, 0))}</span>
                </div>
                {list.map((r) => (
                  <div key={r.id} className={"res-line" + (r.status === "cancelled" ? " cancelled" : "")}>
                    <div className="res-main">
                      <b>{r.name} 様</b>（{r.people}名・{yen(r.total)}）
                      {r.status === "cancelled" && <span className="badge red">キャンセル</span>}
                      {r.checkedIn && <span className="badge green">来園済</span>}
                      {r.status === "active" && (
                        <button type="button" className="btn small admin-cancel-btn" onClick={() => openCancel(r)}>キャンセル</button>
                      )}
                    </div>
                    <div className="res-sub">{r.id}｜{r.phone}｜{breakdownLines(r.counts).join("、")}{r.note ? `｜備考：${r.note}` : ""}</div>
                  </div>
                ))}
              </div>
            );
          })}
        </div>
      )}

      {/* 管理者キャンセルモーダル */}
      {cancelTarget && !cancelDone && (
        <div className="modal-bg">
          <div className="modal admin-cancel-modal">
            {cancelStep === "reason" && (
              <>
                <p className="modal-msg">予約のキャンセル</p>
                <div className="cancel-target-info">
                  <div><b>{cancelTarget.name} 様</b>（{cancelTarget.people}名・{yen(cancelTarget.total)}）</div>
                  <div className="res-sub">{jpDate(cancelTarget.date)} {cancelTarget.slot}〜｜{cancelTarget.id}</div>
                </div>
                <label className="fld cancel-reason-fld">キャンセル理由<span className="req">必須</span>
                  <textarea value={cancelReason} onChange={(e) => setCancelReason(e.target.value)} rows={3} placeholder="例）台風接近のため臨時休園" />
                </label>
                <div className="btn-row modal-btns">
                  <button type="button" className="btn ghost" onClick={closeCancel}>閉じる</button>
                  <button type="button" className="btn danger" disabled={!cancelReason.trim()} onClick={proceedToConfirm}>次へ</button>
                </div>
              </>
            )}
            {cancelStep === "confirm" && (
              <>
                <p className="modal-msg">本当にキャンセルしますか？</p>
                <div className="cancel-target-info">
                  <div><b>{cancelTarget.name} 様</b>（{cancelTarget.people}名・{yen(cancelTarget.total)}）</div>
                  <div className="res-sub">{jpDate(cancelTarget.date)} {cancelTarget.slot}〜｜{cancelTarget.id}</div>
                  <div className="cancel-reason-preview">理由：{cancelReason}</div>
                </div>
                <div className="btn-row modal-btns">
                  <button type="button" className="btn ghost" onClick={() => setCancelStep("reason")}>戻る</button>
                  <button type="button" className="btn danger" onClick={proceedToPassword}>はい</button>
                </div>
              </>
            )}
            {cancelStep === "password" && (
              <>
                <p className="modal-msg">管理者パスワードを入力</p>
                {cancelPwErr && <div className="alert">{cancelPwErr}</div>}
                <label className="fld cancel-reason-fld">パスワード
                  <input type="password" value={cancelPw} onChange={(e) => setCancelPw(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter" && cancelPw) execCancel(); }} />
                </label>
                <div className="btn-row modal-btns">
                  <button type="button" className="btn ghost" onClick={() => setCancelStep("confirm")}>戻る</button>
                  <button type="button" className="btn danger" disabled={!cancelPw || cancelBusy} onClick={execCancel}>
                    {cancelBusy ? "処理中…" : "キャンセルを実行"}
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* 管理者キャンセル完了 + メール通知 */}
      {cancelDone && (
        <div className="modal-bg">
          <div className="modal admin-cancel-modal cancel-done-modal">
            <div className="big-icon">🌱</div>
            <p className="modal-msg">キャンセルが完了しました</p>
            <div className="cancel-target-info">
              <div><b>{cancelDone.name} 様</b>（{cancelDone.people}名・{yen(cancelDone.total)}）</div>
              <div className="res-sub">{jpDate(cancelDone.date)} {cancelDone.slot}〜｜{cancelDone.id}</div>
            </div>
            <button type="button" className="btn primary wide" onClick={closeCancel}>閉じる</button>
          </div>
        </div>
      )}
    </>
  );
}

/* --- 本日の受付（来園チェック） --- */
function AdminCheckin({ reservations, refresh }) {
  const [date, setDate] = useState(todayStr());
  const [busyId, setBusyId] = useState("");
  const list = reservations.filter((r) => r.date === date && r.status === "active")
    .sort((a, b) => a.slot.localeCompare(b.slot));

  const toggle = async (id, currentCheckedIn) => {
    setBusyId(id);
    await updateReservation(id, { checked_in: !currentCheckedIn });
    await refresh();
    setBusyId("");
  };

  const doneCount = list.filter((r) => r.checkedIn).length;

  return (
    <div className="card">
      <h2 className="sec-title">受付・来園チェック</h2>
      <label className="fld">対象日
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} />
      </label>
      <p className="muted">{jpDate(date)}｜{list.length}組（来園済 {doneCount}組）</p>
      {list.length === 0 ? (
        <p className="muted">この日の予約はありません。</p>
      ) : (
        list.map((r) => (
          <div key={r.id} className={"checkin-row" + (r.checkedIn ? " done" : "")}>
            <label className="checkin-label">
              <input type="checkbox" checked={!!r.checkedIn} disabled={busyId === r.id} onChange={() => toggle(r.id, r.checkedIn)} />
              <span>
                <b>{r.slot}〜　{r.name} 様</b>（{r.people}名・{yen(r.total)}）
                <span className="res-sub block">{r.phone}｜{breakdownLines(r.counts).join("、")}{r.note ? `｜備考：${r.note}` : ""}</span>
              </span>
            </label>
            {r.checkedIn && <span className="badge green">来園済</span>}
          </div>
        ))
      )}
    </div>
  );
}

/* --- 設定（基本・月別・カレンダー式の日別設定） --- */
function AdminSettings({ settings, saveSettings, reservations }) {
  const now0 = new Date();
  const [msg, setMsg] = useState("");
  const flash = (t) => { setMsg(t); setTimeout(() => setMsg(""), 2500); };

  const [cap, setCap] = useState(settings.defaultCapacity);
  const [dayCap, setDayCap] = useState(settings.defaultDayCapacity ?? "");
  const [actSlots, setActSlots] = useState(settings.activeSlots ?? [...SLOTS]);
  const toggleActSlot = (s) => {
    setActSlots(actSlots.includes(s) ? actSlots.filter((x) => x !== s) : [...SLOTS].filter((x) => actSlots.includes(x) || x === s));
  };
  const [ym, setYm] = useState([now0.getFullYear(), now0.getMonth()]);
  const [mDraft, setMDraft] = useState(null);
  const [selDate, setSelDate] = useState("");
  const [draft, setDraft] = useState(null);
  const [newPw, setNewPw] = useState("");

  const [year, month] = ym;
  const ymKey = `${year}-${String(month + 1).padStart(2, "0")}`;

  useEffect(() => {
    const m = settings.monthSettings?.[ymKey] || {};
    setMDraft({
      slotCapacity: m.slotCapacity ?? "",
      dayCapacity: m.dayCapacity ?? "",
      closedWeekdays: m.closedWeekdays ? [...m.closedWeekdays] : [],
    });
  }, [ymKey, settings]);

  /* --- 基本設定 --- */
  const saveDefault = async () => {
    await saveSettings({
      ...settings,
      activeSlots: actSlots,
      defaultCapacity: Math.max(0, Number(cap) || 0),
      defaultDayCapacity: dayCap === "" ? "" : Math.max(0, Number(dayCap) || 0),
    });
    flash("基本設定を保存しました");
  };

  /* --- 月別設定 --- */
  const toggleWd = (i) => {
    const list = mDraft.closedWeekdays.includes(i)
      ? mDraft.closedWeekdays.filter((x) => x !== i)
      : [...mDraft.closedWeekdays, i].sort();
    setMDraft({ ...mDraft, closedWeekdays: list });
  };
  const saveMonth = async () => {
    const all = { ...(settings.monthSettings || {}) };
    const entry = {};
    if (mDraft.slotCapacity !== "") entry.slotCapacity = Number(mDraft.slotCapacity);
    if (mDraft.dayCapacity !== "") entry.dayCapacity = Number(mDraft.dayCapacity);
    if (mDraft.closedWeekdays.length) entry.closedWeekdays = mDraft.closedWeekdays;
    if (Object.keys(entry).length === 0) delete all[ymKey];
    else all[ymKey] = entry;
    await saveSettings({ ...settings, monthSettings: all });
    flash(`${year}年${month + 1}月の設定を保存しました`);
  };
  const clearMonth = async () => {
    const all = { ...(settings.monthSettings || {}) };
    delete all[ymKey];
    await saveSettings({ ...settings, monthSettings: all });
    flash(`${year}年${month + 1}月の設定を削除しました`);
  };

  /* --- 日別設定 --- */
  const loadDate = (d) => {
    setSelDate(d);
    const ds = settings.dateSettings[d] || {};
    setDraft({
      closedAll: !!ds.closedAll,
      openOverride: !!ds.openOverride,
      dayCapacity: ds.dayCapacity ?? "",
      slots: SLOTS.reduce((o, s) => {
        const sl = ds.slots?.[s];
        o[s] = { capacity: sl?.capacity ?? "", mode: sl?.mode || (sl?.closed ? "closed" : "default") };
        return o;
      }, {}),
    });
  };
  const saveDate = async () => {
    const all = { ...settings.dateSettings };
    const slots = {};
    for (const s of SLOTS) {
      const v = draft.slots[s];
      if (v.mode !== "default" || v.capacity !== "") {
        slots[s] = { capacity: v.capacity === "" ? null : Number(v.capacity), mode: v.mode };
      }
    }
    const entry = {};
    if (draft.closedAll) entry.closedAll = true;
    if (draft.openOverride) entry.openOverride = true;
    if (draft.dayCapacity !== "") entry.dayCapacity = Number(draft.dayCapacity);
    if (Object.keys(slots).length) entry.slots = slots;
    if (Object.keys(entry).length === 0) delete all[selDate];
    else all[selDate] = entry;
    await saveSettings({ ...settings, dateSettings: all });
    flash(`${jpDate(selDate)} の設定を保存しました`);
  };
  const removeDate = async (d) => {
    const all = { ...settings.dateSettings };
    delete all[d];
    await saveSettings({ ...settings, dateSettings: all });
    if (d === selDate) loadDate(d);
    flash("個別設定を削除しました");
  };
  const savePw = async () => {
    if (newPw.trim().length < 4) { flash("パスワードは4文字以上にしてください"); return; }
    await saveSettings({ ...settings, adminPassword: newPw.trim() });
    setNewPw("");
    flash("パスワードを変更しました");
  };

  /* --- 営業日カレンダー --- */
  const first = new Date(year, month, 1);
  const startDow = first.getDay();
  const dim = new Date(year, month + 1, 0).getDate();
  const cells = [];
  for (let k = 0; k < startDow; k++) cells.push(null);
  for (let d = 1; d <= dim; d++) cells.push(d);

  const configured = Object.keys(settings.dateSettings).sort();
  const selWeekdayClosed = selDate
    ? !!settings.monthSettings?.[selDate.slice(0, 7)]?.closedWeekdays?.includes(wdOf(selDate))
    : false;

  return (
    <>
      {msg && <div className="toast">{msg}</div>}

      <div className="card">
        <h2 className="sec-title">基本設定（全期間共通の初期値）</h2>
        <div className="wd-title">営業する時間帯</div>
        <div className="wd-checks">
          {SLOTS.map((s) => (
            <label key={s} className={"wd-check" + (actSlots.includes(s) ? " on" : "")}>
              <input type="checkbox" checked={actSlots.includes(s)} onChange={() => toggleActSlot(s)} />{s}
            </label>
          ))}
        </div>
        <div className="set-grid">
          <label>各時間帯の最大予約人数
            <input type="number" min="0" className="num-input" value={cap} onChange={(e) => setCap(e.target.value)} />
          </label>
          <label>1日の最大予約人数（空欄＝制限なし）
            <input type="number" min="0" className="num-input" value={dayCap} placeholder="制限なし" onChange={(e) => setDayCap(e.target.value)} />
          </label>
        </div>
        <button type="button" className="btn primary" onClick={saveDefault}>基本設定を保存</button>
        <p className="muted small">チェックを外した時間帯は、全日程で予約受付を停止します（特定の日だけ営業・休止したい場合は、営業日カレンダーから日別に設定できます）。1日の合計予約人数が「1日の最大」に達すると、各時間帯に空きが残っていても、その日は予約できなくなります。</p>
      </div>

      <div className="card">
        <div className="cal-nav">
          <button type="button" className="cal-arrow" onClick={() => setYm(month === 0 ? [year - 1, 11] : [year, month - 1])}>◀</button>
          <div className="cal-title">{year}年 {month + 1}月</div>
          <button type="button" className="cal-arrow" onClick={() => setYm(month === 11 ? [year + 1, 0] : [year, month + 1])}>▶</button>
        </div>

        <h2 className="sec-title">この月のデフォルト設定</h2>
        {mDraft && (
          <>
            <div className="set-grid">
              <label>各時間帯の最大予約人数（空欄＝基本 {settings.defaultCapacity}名）
                <input type="number" min="0" className="num-input" value={mDraft.slotCapacity} placeholder={String(settings.defaultCapacity)}
                  onChange={(e) => setMDraft({ ...mDraft, slotCapacity: e.target.value })} />
              </label>
              <label>1日の最大予約人数（空欄＝基本{settings.defaultDayCapacity === "" ? "：制限なし" : ` ${settings.defaultDayCapacity}名`}）
                <input type="number" min="0" className="num-input" value={mDraft.dayCapacity}
                  placeholder={settings.defaultDayCapacity === "" ? "制限なし" : String(settings.defaultDayCapacity)}
                  onChange={(e) => setMDraft({ ...mDraft, dayCapacity: e.target.value })} />
              </label>
            </div>
            <div className="wd-title">定休日（毎週）</div>
            <div className="wd-checks">
              {WEEKDAYS.map((w, i2) => (
                <label key={w} className={"wd-check" + (mDraft.closedWeekdays.includes(i2) ? " on" : "")}>
                  <input type="checkbox" checked={mDraft.closedWeekdays.includes(i2)} onChange={() => toggleWd(i2)} />{w}
                </label>
              ))}
            </div>
            <div className="btn-row">
              <button type="button" className="btn ghost" onClick={clearMonth}>この月の設定を削除</button>
              <button type="button" className="btn primary" onClick={saveMonth}>この月の設定を保存</button>
            </div>
          </>
        )}

        <h2 className="sec-title cal-sec">営業日カレンダー</h2>
        <p className="muted small">日付をタップすると、その日だけの個別設定（休園・定員変更）ができます。●印は個別設定あり。</p>
        <div className="cal-grid cal-week">
          {WEEKDAYS.map((w, i2) => (
            <div key={w} className={"cal-wd" + (i2 === 0 ? " sun" : i2 === 6 ? " sat" : "")}>{w}</div>
          ))}
        </div>
        <div className="cal-grid">
          {cells.map((d, k) => {
            if (d === null) return <div key={"e" + k} />;
            const iso = fmtDate(new Date(year, month, d));
            const info = resolveDay(iso, settings);
            const booked = reservations.filter((r) => r.status === "active" && r.date === iso).reduce((s, r) => s + r.people, 0);
            const custom = !!settings.dateSettings[iso];
            return (
              <button type="button" key={iso}
                className={"adm-day" + (info.dayClosed ? " closedday" : "") + (selDate === iso ? " selected" : "") + (iso === todayStr() ? " today" : "")}
                onClick={() => loadDate(iso)}>
                <span className="cal-num">{d}</span>
                <span className="adm-info">{info.dayClosed ? "休" : `${booked}${info.dayCapacity != null ? "/" + info.dayCapacity : ""}名`}</span>
                {custom && <span className="adm-dot">●</span>}
              </button>
            );
          })}
        </div>
      </div>

      {selDate && draft && (
        <div className="card">
          <h2 className="sec-title">{jpDate(selDate)} の個別設定</h2>
          <label className="closed-all">
            <input type="checkbox" checked={draft.closedAll} onChange={(e) => setDraft({ ...draft, closedAll: e.target.checked })} />
            この日は終日予約不可にする（休園）
          </label>
          {selWeekdayClosed && !draft.closedAll && (
            <label className="closed-all">
              <input type="checkbox" checked={draft.openOverride} onChange={(e) => setDraft({ ...draft, openOverride: e.target.checked })} />
              毎週の定休日ですが、この日は営業する
            </label>
          )}
          {!draft.closedAll && (
            <>
              <label className="fld">この日の最大予約人数（空欄＝月別・基本設定に従う）
                <input type="number" min="0" className="num-input" value={draft.dayCapacity}
                  onChange={(e) => setDraft({ ...draft, dayCapacity: e.target.value })} />
              </label>
              <table className="list-table">
                <thead><tr><th>時間</th><th>営業</th><th>定員（空欄＝デフォルト）</th></tr></thead>
                <tbody>
                  {SLOTS.map((s) => {
                    const st = slotStatus(selDate, s, reservations, settings);
                    const defaultOpen = !settings.activeSlots || settings.activeSlots.includes(s);
                    return (
                      <tr key={s}>
                        <td><b>{s}</b><span className="res-sub block">予約済 {st.booked}名</span></td>
                        <td>
                          <select className="sel-input" value={draft.slots[s].mode}
                            onChange={(e) => setDraft({ ...draft, slots: { ...draft.slots, [s]: { ...draft.slots[s], mode: e.target.value } } })}>
                            <option value="default">デフォルト（{defaultOpen ? "営業" : "休止"}）</option>
                            <option value="open">この日は営業する</option>
                            <option value="closed">この日は休止する</option>
                          </select>
                        </td>
                        <td><input type="number" min="0" className="num-input" value={draft.slots[s].capacity}
                          onChange={(e) => setDraft({ ...draft, slots: { ...draft.slots, [s]: { ...draft.slots[s], capacity: e.target.value } } })} /></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </>
          )}
          <div className="btn-row">
            {settings.dateSettings[selDate] && (
              <button type="button" className="btn ghost" onClick={() => removeDate(selDate)}>個別設定を削除</button>
            )}
            <button type="button" className="btn primary" onClick={saveDate}>この日の設定を保存</button>
          </div>
        </div>
      )}

      {configured.length > 0 && (
        <div className="card">
          <h2 className="sec-title">個別設定のある日</h2>
          {configured.map((d) => {
            const ds = settings.dateSettings[d];
            const parts = [];
            if (ds.closedAll) parts.push("終日予約不可");
            if (ds.openOverride) parts.push("定休日を営業");
            if (ds.dayCapacity != null) parts.push(`1日最大${ds.dayCapacity}名`);
            for (const s of SLOTS) {
              const sl = ds.slots?.[s];
              if (!sl) continue;
              const m = sl.mode || (sl.closed ? "closed" : "default");
              const bits = [];
              if (m === "closed") bits.push("休止");
              if (m === "open") bits.push("営業");
              if (sl.capacity != null) bits.push(`定員${sl.capacity}名`);
              parts.push(`${s}${bits.join("・")}`);
            }
            return (
              <div key={d} className="cfg-row">
                <button type="button" className="link-btn" onClick={() => { const [yy, mm2] = d.split("-").map(Number); setYm([yy, mm2 - 1]); loadDate(d); }}>{jpDate(d)}</button>
                <span className="res-sub">{parts.join("、")}</span>
                <button type="button" className="btn small ghost" onClick={() => removeDate(d)}>削除</button>
              </div>
            );
          })}
        </div>
      )}

      <div className="card">
        <h2 className="sec-title">管理者パスワードの変更</h2>
        <div className="inline-row">
          <input type="password" className="pw-input" value={newPw} onChange={(e) => setNewPw(e.target.value)} placeholder="新しいパスワード" />
          <button type="button" className="btn primary" onClick={savePw}>変更</button>
        </div>
      </div>
    </>
  );
}

/* ============ トップ & ルート ============ */
export default function App() {
  const [view, setView] = useState("home");
  const [reservations, setReservations] = useState([]);
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState("");
  const [tapInfo, setTapInfo] = useState({ n: 0, t: 0 });
  const footerTap = () => {
    const nowT = Date.now();
    const cnt = nowT - tapInfo.t < 1500 ? tapInfo.n + 1 : 1;
    if (cnt >= 5) { setTapInfo({ n: 0, t: 0 }); setView("admin"); }
    else setTapInfo({ n: cnt, t: nowT });
  };

  const refresh = async () => {
    try {
      const [r, s] = await Promise.all([fetchReservations(), fetchSettings()]);
      setReservations(r);
      setSettings({ ...DEFAULT_SETTINGS, ...s });
      setLoadError("");
    } catch (e) {
      console.error("refresh error", e);
      setLoadError("データの取得に失敗しました。ページを再読み込みしてください。");
    }
  };
  useEffect(() => { (async () => { await refresh(); setLoading(false); })(); }, []);

  const saveSettingsHandler = async (next) => {
    setSettings(next);
    const ok = await upsertSettings(next);
    if (!ok) {
      setLoadError("設定の保存に失敗しました。時間をおいて再度お試しください。");
    }
  };

  return (
    <div className="app">
      <header className="header" onClick={() => setView("home")}>
        <div className="header-inner">
          <span className="header-berry">🍓</span>
          <div>
            <div className="header-name">{FARM.name}</div>
            <div className="header-sub">いちご狩り予約</div>
          </div>
        </div>
      </header>

      {loadError && <div className="page"><div className="alert">{loadError}</div></div>}

      {loading ? (
        <div className="page"><div className="card center-card"><p className="muted">読み込み中…</p></div></div>
      ) : view === "home" ? (
        <div className="page">
          <div className="hero">
            <p className="hero-lead">甘くておいしい、やんばる育ちのいちご。<br />40分間の食べ放題をお楽しみください。</p>
            <div className="hero-terms">
              {TERMS.map((t) => <div key={t} className="hero-term">・{t}</div>)}
            </div>
          </div>
          <button type="button" className="btn primary big" onClick={() => setView("book")}>🍓 いちご狩りを予約する</button>
          <button type="button" className="btn ghost wide" onClick={() => setView("cancel")}>予約のキャンセルはこちら</button>
          <div className="price-card card">
            <h2 className="sec-title">料金（お一人様・40分食べ放題）</h2>
            <table className="list-table">
              <tbody>
                {CATS.map((c) => (
                  <tr key={c.key}><td>{c.label}</td><td className="price-td">{yen(c.price)}</td></tr>
                ))}
              </tbody>
            </table>
            <p className="muted small">受付時間：{(settings.activeSlots ?? INITIAL_ACTIVE_SLOTS).join(" / ")}</p>
          </div>
          <div className="home-links">
            <a className="home-link" href={FARM.mapUrl} target="_blank" rel="noreferrer">📍 アクセス（Googleマップで開く）</a>
            <a className="home-link" href={FARM.instagram} target="_blank" rel="noreferrer">📷 Instagramはこちら</a>
          </div>
          <button type="button" className="admin-link" onClick={() => setView("admin")}>🔑 管理者の方はこちら（公開時は非表示にできます）</button>
        </div>
      ) : view === "book" ? (
        <BookingApp reservations={reservations} settings={settings} refresh={refresh} goHome={() => setView("home")} />
      ) : view === "cancel" ? (
        <CancelApp settings={settings} refresh={refresh} goHome={() => setView("home")} />
      ) : (
        <AdminApp reservations={reservations} settings={settings} refresh={refresh} saveSettings={saveSettingsHandler} goHome={() => setView("home")} />
      )}

      <footer className="footer" onClick={footerTap}>
        {FARM.name}｜{FARM.email}｜{FARM.tel}
      </footer>
    </div>
  );
}
