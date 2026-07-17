/* 空き状況・定員判定ロジック（フロント／サーバーAPIの両方から使う共通モジュール） */

export const SLOTS = ["10:00", "11:00", "12:00", "13:00", "13:30", "14:00", "15:00", "16:00", "17:00", "18:00", "19:00", "20:00"];
export const INITIAL_ACTIVE_SLOTS = ["10:00", "11:00", "12:00", "13:30", "14:00"];
export const FEW_LEFT_THRESHOLD = 5;

export const DEFAULT_SETTINGS = {
  activeSlots: [...INITIAL_ACTIVE_SLOTS],
  defaultCapacity: 20,
  defaultDayCapacity: "",
  monthSettings: {},
  dateSettings: {},
};

export function numOr(v, fallback) {
  return v === "" || v == null ? fallback : Number(v);
}
export function wdOf(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d).getDay();
}

/* 設定の優先度：日別 ＞ 月別 ＞ 基本 */
export function resolveDay(dateIso, settings) {
  const ds = settings.dateSettings[dateIso];
  const ms = settings.monthSettings?.[dateIso.slice(0, 7)];
  const weekdayClosed = !!ms?.closedWeekdays?.includes(wdOf(dateIso));
  const dayClosed = !!ds?.closedAll || (weekdayClosed && !ds?.openOverride);
  const dayCapacity = numOr(ds?.dayCapacity, numOr(ms?.dayCapacity, numOr(settings.defaultDayCapacity, null)));
  const slotDefault = numOr(ms?.slotCapacity, settings.defaultCapacity);
  return { ds, dayClosed, weekdayClosed, dayCapacity, slotDefault };
}

export function slotStatus(dateIso, slot, reservations, settings) {
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

export function dayAvailability(dateIso, reservations, settings) {
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

/* 顧客向けに数字を出さない判定（○/△/×相当）。remaining等の生データはここで止める */
export function publicSlotStatus(dateIso, slot, reservations, settings) {
  const st = slotStatus(dateIso, slot, reservations, settings);
  if (st.closed) return "closed";
  if (st.remaining <= 0) return "full";
  if (st.remaining <= FEW_LEFT_THRESHOLD) return "few";
  return "open";
}

export function publicDayStatus(dateIso, reservations, settings) {
  const av = dayAvailability(dateIso, reservations, settings);
  if (!av.anyOpen) return "full";
  if (av.totalRemaining <= FEW_LEFT_THRESHOLD) return "few";
  return "open";
}
