import { createClient } from "@supabase/supabase-js";
import { DEFAULT_SETTINGS, SLOTS, publicDayStatus, publicSlotStatus } from "../src/lib/availability.js";

const SUPABASE_URL = "https://cnhntxoxrvjajeyvibnu.supabase.co";

function daysInMonth(ym) {
  const [y, m] = ym.split("-").map(Number);
  return new Date(y, m, 0).getDate();
}

/* 顧客向け空き状況API。settings/reservationsの生データ(定員・人数)はここで止め、
   ○(open)/△(few)/×(full)/休園(closed) の判定結果だけを返す。 */
export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");

  const { month, date } = req.query;
  if (!month || !/^\d{4}-\d{2}$/.test(month)) {
    res.status(400).json({ error: "month (YYYY-MM) is required" });
    return;
  }

  const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!serviceKey) {
    console.error("SUPABASE_SERVICE_ROLE_KEY is not set");
    res.status(500).json({ error: "server not configured" });
    return;
  }
  const supabase = createClient(SUPABASE_URL, serviceKey);

  try {
    const [{ data: settingsRow }, { data: reservations, error: resErr }] = await Promise.all([
      supabase.from("settings").select("data").eq("id", "main").maybeSingle(),
      supabase
        .from("reservations")
        .select("date, slot, people, status")
        .eq("status", "active")
        .gte("date", `${month}-01`)
        .lt("date", `${month}-32`),
    ]);
    if (resErr) throw resErr;

    const settings = { ...DEFAULT_SETTINGS, ...(settingsRow?.data || {}) };
    const dim = daysInMonth(month);
    const days = {};
    for (let d = 1; d <= dim; d++) {
      const iso = `${month}-${String(d).padStart(2, "0")}`;
      days[iso] = publicDayStatus(iso, reservations || [], settings);
    }

    let slots = null;
    if (date && /^\d{4}-\d{2}-\d{2}$/.test(date) && date.startsWith(month)) {
      slots = {};
      for (const s of SLOTS) {
        slots[s] = publicSlotStatus(date, s, reservations || [], settings);
      }
    }

    res.status(200).json({ days, slots });
  } catch (e) {
    console.error("public-availability error", e);
    res.status(500).json({ error: "internal error" });
  }
}
