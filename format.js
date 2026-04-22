// THEMIS format helpers (stage 1 split)
window.__THEMIS_FORMAT_LOADED = true;

function formatDateTimeJa(value) {
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value || "");
  return d.toLocaleString("ja-JP");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function toNullableNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function normalizeStatus(status) {
  if (status === "done") return "done";
  if (status === "cancel") return "cancel";
  if (status === "assigned") return "assigned";
  return "pending";
}

function getStatusText(status) {
  const s = normalizeStatus(status);
  if (s === "done") return "完了";
  if (s === "cancel") return "キャンセル";
  if (s === "assigned") return "配車済";
  return "未完了";
}

function getHourLabel(hour) {
  const n = Number(hour);
  return `${n}時`;
}

function normalizeAddressText(address) {
  return String(address || "")
    .trim()
    .replace(/[　\s]+/g, "")
    .replace(/ヶ/g, "ケ")
    .replace(/之/g, "の")
    .replace(/−/g, "-")
    .replace(/ー/g, "-");
}

function detectPrefecture(address) {
  const a = normalizeAddressText(address);

  if (a.includes("東京都")) return "東京";
  if (a.includes("埼玉県")) return "埼玉";
  if (a.includes("千葉県")) return "千葉";
  if (a.includes("茨城県")) return "茨城";

  return "";
}

function normalizeCastInputValue(value) {
  return String(value || "").trim();
}

function normalizeAreaLabel(area) {
  const value = String(area || "").trim();
  if (!value) return "無し";
  return value;
}

function csvEscape(value) {
  const s = String(value ?? "");
  if (s.includes('"') || s.includes(",") || s.includes("\n")) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

function parseCsvLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const next = line[i + 1];

    if (char === '"' && inQuotes && next === '"') {
      current += '"';
      i++;
      continue;
    }
    if (char === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (char === "," && !inQuotes) {
      result.push(current.trim());
      current = "";
      continue;
    }
    current += char;
  }

  result.push(current.trim());
  return result;
}

function normalizeCsvHeader(header) {
  const raw = String(header || "").trim();
  const h = raw.toLowerCase();

  const map = {
    name: "name",
    名前: "name",
    cast_name: "name",

    phone: "phone",
    tel: "phone",
    telephone: "phone",
    電話: "phone",
    電話番号: "phone",

    address: "address",
    住所: "address",

    area: "area",
    方面: "area",
    地域: "area",

    latitude: "latitude",
    lat: "latitude",
    緯度: "latitude",

    longitude: "longitude",
    lng: "longitude",
    lon: "longitude",
    経度: "longitude",

    memo: "memo",
    メモ: "memo",
    note: "memo",

    distance_km: "distance_km",
    距離: "distance_km",
    想定距離: "distance_km",

    travel_minutes: "travel_minutes",
    travel_time: "travel_minutes",
    片道予想時間: "travel_minutes",
    "片道予想時間(分)": "travel_minutes",
    予想時間: "travel_minutes",
    時間: "travel_minutes",

    plate_number: "plate_number",
    vehicle_id: "plate_number",
    車両id: "plate_number",
    車両ID: "plate_number",
    車両: "plate_number",

    vehicle_area: "vehicle_area",
    担当方面: "vehicle_area",

    home_area: "home_area",
    帰宅方面: "home_area",

    seat_capacity: "seat_capacity",
    定員: "seat_capacity",
    乗車可能人員: "seat_capacity",

    driver_name: "driver_name",
    driver: "driver_name",
    ドライバー: "driver_name",
    ドライバー名: "driver_name",

    line_id: "line_id",
    line: "line_id",
    lineid: "line_id",
    "line id": "line_id",
    LINEID: "line_id",
    "LINE ID": "line_id",
    line_id_: "line_id",

    status: "status",
    状態: "status"
  };

  return map[raw] || map[h] || h;
}

function formatMileageSheetDate(date) {
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${m}/${d}`;
}