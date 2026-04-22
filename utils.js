// utils
// dashboard.js から安全に分離した共通ユーティリティ

function buildMapUrlFromAddressOrLatLng(address, lat, lng) {
  const numLat = toNullableNumber(lat);
  const numLng = toNullableNumber(lng);

  if (isValidLatLng(numLat, numLng)) {
    return `https://www.google.com/maps/search/?api=1&query=${numLat},${numLng}`;
  }

  const safeAddress = String(address || "").trim();
  if (safeAddress) {
    return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(safeAddress)}`;
  }

  return "";
}

function buildCastMapUrl(cast) {
  return buildMapUrlFromAddressOrLatLng(
    cast?.address,
    cast?.latitude,
    cast?.longitude
  );
}

function buildDispatchItemMapUrl(item) {
  return buildMapUrlFromAddressOrLatLng(
    item?.destination_address || item?.casts?.address || "",
    item?.casts?.latitude,
    item?.casts?.longitude
  );
}

function buildMapLinkHtml({ name, address, lat, lng, className = "map-name-link" }) {
  const safeName = escapeHtml(name || "-");
  const mapUrl = buildMapUrlFromAddressOrLatLng(address, lat, lng);

  if (!mapUrl) return safeName;

  return `<a href="${mapUrl}" target="_blank" rel="noopener noreferrer" class="${className}">${safeName} 📍</a>`;
}

function downloadTextFile(filename, text, mimeType = "text/plain;charset=utf-8") {
  const content = String(text ?? "");
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;

  document.body.appendChild(a);
  a.click();
  a.remove();

  window.setTimeout(() => {
    try {
      URL.revokeObjectURL(url);
    } catch (_) {}
  }, 1000);
}

function parseCsv(text) {
  const lines = text
    .replace(/^\uFEFF/, "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .filter(line => line.trim() !== "");

  if (!lines.length) return [];

  const headers = parseCsvLine(lines[0]).map(h => normalizeCsvHeader(h.trim()));
  const rows = [];

  for (let i = 1; i < lines.length; i++) {
    const values = parseCsvLine(lines[i]);
    const row = {};
    headers.forEach((header, idx) => {
      row[header] = values[idx] ?? "";
    });
    rows.push(row);
  }

  return rows;
}

async function readCsvFileAsText(file) {
  const buffer = await file.arrayBuffer();
  let text = new TextDecoder("utf-8").decode(buffer);

  const mojibakeLike =
    text.includes("�") ||
    (!text.includes("name") &&
      !text.includes("address") &&
      !text.includes("名前") &&
      !text.includes("住所"));

  if (mojibakeLike) {
    try {
      text = new TextDecoder("shift_jis").decode(buffer);
    } catch (e) {
      console.warn("shift_jis decode failed:", e);
    }
  }

  return text;
}

function normalizeCsvRows(rows) {
  return rows.map(row => {
    const normalized = {};
    Object.keys(row).forEach(key => {
      const nk = normalizeCsvHeader(key);
      normalized[nk] = row[key];
    });
    return normalized;
  });
}
