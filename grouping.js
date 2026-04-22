// grouping
// 方面判定・表示グループ・親和性

const AREA_CANONICAL_PATTERNS = [
  ["松戸近郊", ["松戸近郊", "松戸方面", "松戸"]],
  ["葛飾方面", ["葛飾方面", "葛飾区", "葛飾"]],
  ["足立方面", ["足立方面", "足立区", "足立"]],
  ["江戸川方面", ["江戸川方面", "江戸川区", "江戸川"]],
  ["市川方面", ["市川方面", "市川", "本八幡", "妙典", "行徳", "下総中山"]],
  ["船橋方面", ["船橋方面", "船橋", "習志野", "西船橋"]],
  ["鎌ヶ谷方面", ["鎌ヶ谷方面", "鎌ケ谷方面", "鎌ヶ谷", "鎌ケ谷", "新鎌ヶ谷", "新鎌ケ谷"]],
  ["我孫子方面", ["我孫子方面", "我孫子", "天王台", "湖北"]],
  ["取手方面", ["取手方面", "取手"]],
  ["藤代方面", ["藤代方面", "藤代", "桐木"]],
  ["守谷方面", ["守谷方面", "守谷"]],
  ["柏方面", ["柏方面", "柏", "南柏", "北柏"]],
  ["柏の葉方面", ["柏の葉方面", "柏の葉", "柏たなか"]],
  ["流山方面", ["流山方面", "流山", "南流山", "おおたかの森"]],
  ["野田方面", ["野田方面", "野田", "運河", "梅郷", "川間"]],
  ["三郷方面", ["三郷方面", "三郷"]],
  ["八潮方面", ["八潮方面", "八潮"]],
  ["草加方面", ["草加方面", "草加", "谷塚方面", "谷塚"]],
  ["吉川方面", ["吉川方面", "吉川"]],
  ["越谷方面", ["越谷方面", "越谷", "新越谷"]],
  ["白井方面", ["白井方面", "白井"]],
  ["印西方面", ["印西方面", "印西"]],
  ["浦安方面", ["浦安方面", "浦安"]],
  ["習志野方面", ["習志野方面", "習志野"]],
  ["八千代方面", ["八千代方面", "八千代"]],
  ["つくば方面", ["つくば方面", "つくば"]],
  ["つくばみらい方面", ["つくばみらい方面", "つくばみらい"]],
  ["牛久方面", ["牛久方面", "牛久"]],
  ["龍ヶ崎方面", ["龍ヶ崎方面", "龍ケ崎方面", "龍ケ崎", "龍ヶ崎"]],
  ["利根方面", ["利根方面", "利根"]],
  ["墨田方面", ["墨田方面", "墨田区", "墨田"]],
  ["江東方面", ["江東方面", "江東区", "江東"]],
  ["荒川方面", ["荒川方面", "荒川区", "荒川"]],
  ["台東方面", ["台東方面", "台東区", "台東"]],
  ["東京方面", ["東京方面", "東京"]],
  ["千葉方面", ["千葉方面", "千葉", "幕張", "蘇我", "稲毛", "都賀"]]
];

const AREA_AFFINITY_MAP = {
  "松戸近郊": { "葛飾方面": 80, "市川方面": 72, "柏方面": 60, "三郷方面": 62, "足立方面": 58 },
  "葛飾方面": { "松戸近郊": 80, "足立方面": 62, "江戸川方面": 58, "市川方面": 55 },
  "足立方面": { "葛飾方面": 62, "松戸近郊": 58, "八潮方面": 55, "草加方面": 52 },
  "江戸川方面": { "葛飾方面": 58, "市川方面": 54, "船橋方面": 46 },
  "市川方面": { "葛飾方面": 55, "松戸近郊": 72, "船橋方面": 68, "鎌ヶ谷方面": 66, "江戸川方面": 54 },
  "船橋方面": { "市川方面": 68, "鎌ヶ谷方面": 76, "千葉方面": 58, "江戸川方面": 46 },
  "鎌ヶ谷方面": { "船橋方面": 76, "市川方面": 66, "柏方面": 56, "我孫子方面": 42 },
  "我孫子方面": { "取手方面": 88, "藤代方面": 84, "柏方面": 70, "守谷方面": 60, "鎌ヶ谷方面": 42 },
  "取手方面": { "我孫子方面": 88, "藤代方面": 92, "守谷方面": 72, "柏方面": 44 },
  "藤代方面": { "我孫子方面": 84, "取手方面": 92, "守谷方面": 70 },
  "守谷方面": { "取手方面": 72, "藤代方面": 70, "我孫子方面": 60, "つくば方面": 62 },
  "柏方面": { "我孫子方面": 70, "流山方面": 66, "柏の葉方面": 64, "野田方面": 56, "鎌ヶ谷方面": 56, "松戸近郊": 60, "取手方面": 44, "白井方面": 60, "印西方面": 54 },
  "柏の葉方面": { "柏方面": 64, "流山方面": 62, "野田方面": 58 },
  "流山方面": { "柏方面": 66, "柏の葉方面": 62, "野田方面": 58, "三郷方面": 56, "吉川方面": 54 },
  "野田方面": { "柏方面": 56, "柏の葉方面": 58, "流山方面": 58, "吉川方面": 52 },
  "三郷方面": { "松戸近郊": 62, "流山方面": 56, "八潮方面": 62, "吉川方面": 54 },
  "八潮方面": { "三郷方面": 62, "草加方面": 56, "足立方面": 55 },
  "草加方面": { "八潮方面": 56, "足立方面": 52, "越谷方面": 52 },
  "吉川方面": { "流山方面": 54, "野田方面": 52, "三郷方面": 54, "越谷方面": 52 },
  "越谷方面": { "草加方面": 52, "吉川方面": 52 },
  "白井方面": { "柏方面": 60, "印西方面": 72, "鎌ヶ谷方面": 62, "船橋方面": 54 },
  "印西方面": { "白井方面": 72, "柏方面": 54, "我孫子方面": 48 },
  "千葉方面": { "船橋方面": 58 }
};

const THEMIS_DISPLAY_GROUPS = new Set([
  "松戸方面",
  "柏方面",
  "我孫子・取手方面",
  "市川・船橋方面",
  "東京方面",
  "埼玉方面",
  "流山・野田方面"
]);

function normalizeAddressText(value = "") {
  return String(value || "")
    .replace(/[　\s]+/g, "")
    .replace(/^日本/, "")
    .replace(/^〒\d{3}-?\d{4}/, "")
    .trim();
}

function normalizeAreaLabel(area = "") {
  return String(area || "").trim().replace(/[　\s]+/g, "");
}

function getCanonicalArea(area) {
  const normalized = normalizeAreaLabel(area);
  if (!normalized || normalized === "無し") return "";
  for (const [canonical, patterns] of AREA_CANONICAL_PATTERNS) {
    if (patterns.some(pattern => normalized.includes(pattern))) return canonical;
  }
  if (normalized.endsWith("方面")) return normalized;
  return normalized;
}

function getAreaDisplayGroup(area) {
  const a = normalizeAreaLabel(area || "");
  if (!a || a === "無し") return "東京方面";

  if (a.includes("松戸")) return "松戸方面";
  if (a.includes("柏") || a.includes("白井") || a.includes("印西")) return "柏方面";

  if (
    a.includes("我孫子") || a.includes("取手") || a.includes("守谷") ||
    a.includes("牛久") || a.includes("藤代") || a.includes("つくば") ||
    a.includes("龍ケ崎") || a.includes("龍ヶ崎") || a.includes("利根")
  ) return "我孫子・取手方面";

  if (
    a.includes("市川") || a.includes("船橋") || a.includes("鎌ヶ谷") ||
    a.includes("鎌ケ谷") || a.includes("浦安") || a.includes("習志野") ||
    a.includes("八千代") || a.includes("千葉")
  ) return "市川・船橋方面";

  if (
    a.includes("足立") || a.includes("葛飾") || a.includes("江戸川") ||
    a.includes("墨田") || a.includes("江東") || a.includes("荒川") ||
    a.includes("台東") || a.includes("中央") || a.includes("千代田") ||
    a.includes("港") || a.includes("新宿") || a.includes("文京") ||
    a.includes("品川") || a.includes("目黒") || a.includes("大田") ||
    a.includes("世田谷") || a.includes("渋谷") || a.includes("中野") ||
    a.includes("杉並") || a.includes("豊島") || a.includes("北") ||
    a.includes("板橋") || a.includes("練馬") || a.includes("東京")
  ) return "東京方面";

  if (
    a.includes("三郷") || a.includes("八潮") || a.includes("草加") ||
    a.includes("越谷") || a.includes("吉川")
  ) return "埼玉方面";

  if (a.includes("流山") || a.includes("野田")) return "流山・野田方面";

  return "東京方面";
}

function getGroupedAreaHeaderHtml(area) {
  const detailArea = normalizeAreaLabel(area || "無し");
  const displayGroup = getAreaDisplayGroup(detailArea);
  if (!detailArea || detailArea === "無し") {
    return `<span class="group-main">${escapeHtml(displayGroup)}</span>`;
  }
  if (displayGroup === detailArea) {
    return `<span class="group-main">${escapeHtml(displayGroup)}</span>`;
  }
  return `
    <div class="group-area-stack">
      <span class="group-main">${escapeHtml(displayGroup)}</span>
      <span class="group-sub">${escapeHtml(detailArea)}</span>
    </div>
  `;
}

function getGroupedAreasByDisplay(items, areaGetter) {
  const ordered = [];
  const seen = new Set();
  for (const item of items) {
    const detailArea = normalizeAreaLabel(areaGetter(item));
    const key = `${getAreaDisplayGroup(detailArea)}__${detailArea}`;
    if (seen.has(key)) continue;
    seen.add(key);
    ordered.push({
      displayGroup: getAreaDisplayGroup(detailArea),
      detailArea
    });
  }
  ordered.sort((a, b) => {
    const groupCompare = a.displayGroup.localeCompare(b.displayGroup, "ja");
    if (groupCompare !== 0) return groupCompare;
    return a.detailArea.localeCompare(b.detailArea, "ja");
  });
  return ordered;
}

function extractCityTownFromAddress(address = "") {
  const raw = String(address || "").trim();
  const a = normalizeAddressText(raw);
  if (!a) return { city: "", town: "" };

  const specificMatch = a.match(
    /(足立区|葛飾区|江戸川区|墨田区|江東区|荒川区|台東区|中央区|千代田区|港区|新宿区|文京区|品川区|目黒区|大田区|世田谷区|渋谷区|中野区|杉並区|豊島区|北区|板橋区|練馬区|松戸市|柏市|白井市|印西市|流山市|我孫子市|野田市|市川市|鎌ケ谷市|鎌ヶ谷市|船橋市|浦安市|習志野市|八千代市|三郷市|八潮市|草加市|越谷市|吉川市|守谷市|取手市|つくば市|つくばみらい市|牛久市|龍ケ崎市|龍ヶ崎市|利根町)(.+)?/
  );
  const genericMatch = a.match(/^(?:北海道|東京都|京都府|大阪府|.{2,3}県)?(.+?(?:市|区|町|村))(.+)?$/);
  const match = specificMatch || genericMatch;
  if (!match) return { city: "", town: "" };

  const cityRaw = match[1] || "";
  let townRaw = match[2] || "";
  townRaw = townRaw
    .replace(/丁目.*$/, "")
    .replace(/番地.*$/, "")
    .replace(/番.*$/, "")
    .replace(/号.*$/, "")
    .replace(/[0-9０-９]+/g, "")
    .replace(/[-‐‑–—―ー－]+/g, "")
    .replace(/^(?:北海道|東京都|京都府|大阪府|.{2,3}県)/, "")
    .replace(/^日本/, "")
    .trim();

  const city = cityRaw.replace(/(市|区|町|村)$/, "");
  return { city, town: townRaw };
}

const GENERIC_PREFECTURES = [
  "北海道", "青森", "岩手", "宮城", "秋田", "山形", "福島",
  "茨城", "栃木", "群馬", "埼玉", "千葉", "東京", "神奈川",
  "新潟", "富山", "石川", "福井", "山梨", "長野",
  "岐阜", "静岡", "愛知", "三重",
  "滋賀", "京都", "大阪", "兵庫", "奈良", "和歌山",
  "鳥取", "島根", "岡山", "広島", "山口",
  "徳島", "香川", "愛媛", "高知",
  "福岡", "佐賀", "長崎", "熊本", "大分", "宮崎", "鹿児島", "沖縄"
];

function detectPrefectureGeneric(address = "") {
  const a = normalizeAddressText(address);
  if (!a) return "";
  const match = a.match(/^(北海道|東京都|京都府|大阪府|.{2,3}県)/);
  if (!match) return "";
  return String(match[1] || "")
    .replace(/都$/, "")
    .replace(/道$/, "")
    .replace(/府$/, "")
    .replace(/県$/, "")
    .trim();
}

function makeCompactTownLabel(town = "") {
  return String(town || "")
    .replace(/(大字|字)/g, "")
    .replace(/(?:ケ丘|ヶ丘)/g, "ケ丘")
    .replace(/[0-9０-９]/g, "")
    .replace(/[\-‐‑–—―ー－]/g, "")
    .trim();
}

function buildGenericAreaByAddress(address = "") {
  const { city, town } = extractCityTownFromAddress(address);
  const normalizedCity = String(city || "").trim();
  const compactTown = makeCompactTownLabel(town);
  if (normalizedCity) {
    if (compactTown) return `${normalizedCity}${compactTown}方面`;
    return `${normalizedCity}方面`;
  }
  const pref = detectPrefectureGeneric(address);
  if (pref) return `${pref}方面`;
  return "";
}

function isCompassLikeAreaLabel(area = "") {
  const normalized = normalizeAreaLabel(area);
  if (!normalized || normalized === "無し") return false;
  if (/^(北|北東|東|南東|南|南西|西|北西)方面$/.test(normalized)) return true;
  if (GENERIC_PREFECTURES.some(pref => normalized === `${pref}方面` || normalized === `${pref}北方面` || normalized === `${pref}北東方面` || normalized === `${pref}東方面` || normalized === `${pref}南東方面` || normalized === `${pref}南方面` || normalized === `${pref}南西方面` || normalized === `${pref}西方面` || normalized === `${pref}北西方面`)) {
    return true;
  }
  return false;
}

function resolvePreferredAreaLabel(area = "", address = "", lat = null, lng = null) {
  const normalized = normalizeAreaLabel(area);
  const byAddress = buildGenericAreaByAddress(address);
  if (byAddress && (!normalized || normalized === "無し" || isCompassLikeAreaLabel(normalized))) {
    return byAddress;
  }
  if (normalized && normalized !== "無し") return normalized;
  return normalizeAreaLabel(guessArea(lat, lng, address) || "無し");
}

function getDetailedAreaByCity(city = "") {
  if (!city) return "東京方面";

  if (["松戸"].includes(city)) return "松戸近郊";

  if (["柏"].includes(city)) return "柏方面";
  if (["白井"].includes(city)) return "白井方面";
  if (["印西"].includes(city)) return "印西方面";

  if (["我孫子"].includes(city)) return "我孫子方面";
  if (["取手"].includes(city)) return "取手方面";
  if (["藤代"].includes(city)) return "藤代方面";
  if (["守谷"].includes(city)) return "守谷方面";
  if (["つくば"].includes(city)) return "つくば方面";
  if (["つくばみらい"].includes(city)) return "つくばみらい方面";
  if (["牛久"].includes(city)) return "牛久方面";
  if (["龍ケ崎", "龍ヶ崎"].includes(city)) return "龍ヶ崎方面";
  if (["利根"].includes(city)) return "利根方面";

  if (["市川"].includes(city)) return "市川方面";
  if (["船橋"].includes(city)) return "船橋方面";
  if (["鎌ケ谷", "鎌ヶ谷"].includes(city)) return "鎌ヶ谷方面";
  if (["浦安"].includes(city)) return "浦安方面";
  if (["習志野"].includes(city)) return "習志野方面";
  if (["八千代"].includes(city)) return "八千代方面";
  if (["千葉"].includes(city)) return "千葉方面";

  if (["足立"].includes(city)) return "足立方面";
  if (["葛飾"].includes(city)) return "葛飾方面";
  if (["江戸川"].includes(city)) return "江戸川方面";
  if (["墨田"].includes(city)) return "墨田方面";
  if (["江東"].includes(city)) return "江東方面";
  if (["荒川"].includes(city)) return "荒川方面";
  if (["台東"].includes(city)) return "台東方面";
  if ([
    "中央", "千代田", "港", "新宿", "文京", "品川", "目黒",
    "大田", "世田谷", "渋谷", "中野", "杉並", "豊島", "北",
    "板橋", "練馬"
  ].includes(city)) return "東京方面";

  if (["三郷"].includes(city)) return "三郷方面";
  if (["八潮"].includes(city)) return "八潮方面";
  if (["草加"].includes(city)) return "草加方面";
  if (["越谷"].includes(city)) return "越谷方面";
  if (["吉川"].includes(city)) return "吉川方面";

  if (["流山"].includes(city)) return "流山方面";
  if (["野田"].includes(city)) return "野田方面";

  return "東京方面";
}

function getTownSpecificArea(city = "", town = "") {
  const t = String(town || "").trim();
  if (!city || !t) return "";

  if (city === "柏") {
    if (/(南?逆井)/.test(t)) return "柏逆井方面";
    if (/根戸/.test(t)) return "柏根戸方面";
    if (/若柴/.test(t)) return "柏若柴方面";
    if (/あけぼの/.test(t)) return "柏あけぼの方面";
    if (/豊四季/.test(t)) return "柏豊四季方面";
    if (/旭町/.test(t)) return "柏旭町方面";
    if (/明原/.test(t)) return "柏明原方面";
    if (/高田/.test(t)) return "柏高田方面";
    if (/篠籠田/.test(t)) return "柏篠籠田方面";
    if (/(柏の葉|十余二|大室|小青田|正連寺)/.test(t)) return "柏の葉方面";
    if (/増尾/.test(t)) return "柏増尾方面";
    if (/中新宿/.test(t)) return "柏中新宿方面";
    if (/光ヶ丘/.test(t)) return "柏光ヶ丘方面";
    if (/酒井根/.test(t)) return "柏酒井根方面";
  }

  if (city === "市川") {
    if (/(堀之内|堀の内|堀ノ内)/.test(t)) return "市川堀の内方面";
    if (/南行徳/.test(t)) return "市川南行徳方面";
    if (/行徳/.test(t)) return "市川行徳方面";
    if (/妙典/.test(t)) return "市川妙典方面";
    if (/末広/.test(t)) return "市川末広方面";
    if (/香取/.test(t)) return "市川香取方面";
    if (/福栄/.test(t)) return "市川福栄方面";
    if (/相之川/.test(t)) return "市川相之川方面";
    if (/曽谷/.test(t)) return "市川曽谷方面";
    if (/宮久保/.test(t)) return "市川宮久保方面";
    if (/菅野/.test(t)) return "市川菅野方面";
    if (/真間/.test(t)) return "市川真間方面";
    if (/国府台/.test(t)) return "市川国府台方面";
  }

  if (city === "松戸") {
    if (/常盤平/.test(t)) return "松戸常盤平方面";
    if (/五香/.test(t)) return "松戸五香方面";
    if (/牧の原/.test(t)) return "松戸牧の原方面";
    if (/新松戸/.test(t)) return "松戸新松戸方面";
    if (/幸谷/.test(t)) return "松戸幸谷方面";
    if (/(北小金|小金)/.test(t)) return "松戸小金方面";
    if (/八柱/.test(t)) return "松戸八柱方面";
    if (/日暮/.test(t)) return "松戸日暮方面";
    if (/稔台/.test(t)) return "松戸稔台方面";
    if (/松飛台/.test(t)) return "松戸松飛台方面";
    if (/西馬橋/.test(t)) return "松戸西馬橋方面";
    if (/馬橋/.test(t)) return "松戸馬橋方面";
    if (/中和倉/.test(t)) return "松戸中和倉方面";
    if (/栄町/.test(t)) return "松戸栄町方面";
  }

  if (city === "船橋") {
    if (/西船/.test(t)) return "船橋西船方面";
    if (/本中山/.test(t)) return "船橋本中山方面";
    if (/東中山/.test(t)) return "船橋東中山方面";
    if (/二子町/.test(t)) return "船橋二子町方面";
    if (/前原/.test(t)) return "船橋前原方面";
    if (/津田沼/.test(t)) return "船橋津田沼方面";
  }

  if (city === "我孫子") {
    if (/天王台/.test(t)) return "我孫子天王台方面";
    if (/柴崎台/.test(t)) return "我孫子柴崎台方面";
    if (/柴崎/.test(t)) return "我孫子柴崎方面";
    if (/青山/.test(t)) return "我孫子青山方面";
    if (/湖北/.test(t)) return "我孫子湖北方面";
    if (/新木/.test(t)) return "我孫子新木方面";
    if (/布佐/.test(t)) return "我孫子布佐方面";
  }

  const compactTown = t
    .replace(/(大字|字)/g, "")
    .replace(/(?:ケ丘|ヶ丘)/g, "ケ丘")
    .replace(/[0-9０-９]/g, "")
    .replace(/[\-‐‑–—―ー－]/g, "")
    .trim();

  if (!compactTown) return "";
  return `${city}${compactTown}方面`;
}

function isBroadDisplayOrCityArea(area = "", city = "") {
  const canonical = getCanonicalArea(area);
  if (!canonical) return true;
  if (THEMIS_DISPLAY_GROUPS.has(canonical)) return true;
  const cityDetail = getDetailedAreaByCity(city);
  return !!cityDetail && canonical === cityDetail;
}

function _normalizeDetectedCandidate(area, city = "", town = "") {
  const canonical = getCanonicalArea(area);
  if (!canonical) return "";

  const townSpecific = getTownSpecificArea(city, town);
  if (townSpecific && isBroadDisplayOrCityArea(canonical, city)) {
    return townSpecific;
  }

  if (THEMIS_DISPLAY_GROUPS.has(canonical) && city) {
    return townSpecific || getDetailedAreaByCity(city);
  }
  return canonical;
}

function detectAreaGroup(address = "") {
  const { city, town } = extractCityTownFromAddress(address);
  const normalizedTown = String(town || "").trim();
  if (!city) return "";

  const townSpecific = getTownSpecificArea(city, normalizedTown);

  const allJobs = [
    ...(Array.isArray(window.currentPlansCache) ? window.currentPlansCache : []),
    ...(Array.isArray(window.currentActualsCache) ? window.currentActualsCache : []),
    ...(Array.isArray(window.allCastsCache) ? window.allCastsCache : [])
  ];

  const normalizedRows = allJobs
    .map(j => {
      const addr = j.address || j.destination_address || "";
      const parsed = extractCityTownFromAddress(addr);
      return {
        city: parsed.city,
        town: parsed.town,
        area: _normalizeDetectedCandidate(
          j.area || j.area_group || j.planned_area || j.destination_area || "",
          parsed.city,
          parsed.town
        )
      };
    })
    .filter(x => x.city && x.area);

  const sameTownAreas = normalizedRows.filter(
    x => x.city === city && normalizedTown && x.town === normalizedTown
  );
  if (sameTownAreas.length > 0) {
    const count = {};
    sameTownAreas.forEach(x => { count[x.area] = (count[x.area] || 0) + 1; });
    const picked = Object.keys(count).sort((a, b) => count[b] - count[a])[0] || "";
    return townSpecific || picked || getDetailedAreaByCity(city);
  }

  if (townSpecific) return townSpecific;

  const sameCityAreas = normalizedRows
    .filter(x => x.city === city)
    .filter(x => !isBroadDisplayOrCityArea(x.area, city));
  if (sameCityAreas.length > 0) {
    const count = {};
    sameCityAreas.forEach(x => { count[x.area] = (count[x.area] || 0) + 1; });
    return Object.keys(count).sort((a, b) => count[b] - count[a])[0] || getDetailedAreaByCity(city);
  }

  return getDetailedAreaByCity(city);
}

function classifyAreaByAddress(address) {
  return detectAreaGroup(address);
}

function getDirection8(lat, lng) {
  if (!isValidLatLng(lat, lng)) return "";
  const dLat = lat - ORIGIN_LAT;
  const dLng = lng - ORIGIN_LNG;
  const angle = Math.atan2(dLat, dLng) * 180 / Math.PI;

  if (angle >= -22.5 && angle < 22.5) return "東";
  if (angle >= 22.5 && angle < 67.5) return "北東";
  if (angle >= 67.5 && angle < 112.5) return "北";
  if (angle >= 112.5 && angle < 157.5) return "北西";
  if (angle >= -67.5 && angle < -22.5) return "南東";
  if (angle >= -112.5 && angle < -67.5) return "南";
  if (angle >= -157.5 && angle < -112.5) return "南西";
  return "西";
}

function classifyAreaByLatLng(lat, lng) {
  if (!isValidLatLng(lat, lng)) return "";

  if (lat >= 35.79 && lat <= 35.86 && lng >= 139.90 && lng <= 139.97) return "松戸近郊";
  if (lat >= 35.79 && lat <= 35.88 && lng >= 139.76 && lng <= 139.85) {
    if (lat >= 35.84) return "草加方面";
    if (lat >= 35.81) return "八潮方面";
    return "八潮方面";
  }
  if (lat >= 35.84 && lat <= 35.92 && lng >= 139.84 && lng <= 139.92) return "三郷方面";
  if (lat >= 35.86 && lat <= 35.91 && lng >= 139.82 && lng <= 139.87) return "吉川方面";
  if (lat >= 35.84 && lat <= 35.89 && lng >= 139.92 && lng <= 139.98) return "柏方面";
  if (lat >= 35.85 && lat <= 35.91 && lng > 139.98 && lng <= 140.05) return "柏の葉方面";
  if (lat >= 35.92 && lat <= 35.99 && lng >= 139.84 && lng <= 139.91) return "野田方面";
  if (lat >= 35.84 && lat <= 35.90 && lng >= 139.88 && lng <= 139.95) return "流山方面";
  if (lat >= 35.85 && lat <= 35.89 && lng > 140.00 && lng <= 140.08) return "我孫子方面";
  if (lat >= 35.70 && lat <= 35.78 && lng >= 139.78 && lng <= 139.86) return "墨田方面";
  if (lat >= 35.73 && lat <= 35.80 && lng >= 139.80 && lng <= 139.88) return "足立方面";
  if (lat >= 35.75 && lat <= 35.79 && lng >= 139.84 && lng <= 139.89) return "葛飾方面";
  if (lat >= 35.67 && lat <= 35.72 && lng >= 139.80 && lng <= 139.86) return "江東方面";
  if (lat >= 35.90 && lat <= 36.02 && lng >= 140.00 && lng <= 140.08) return "藤代方面";
  if (lat >= 35.88 && lat <= 35.95 && lng >= 140.03 && lng <= 140.10) return "取手方面";
  if (lat >= 35.93 && lat <= 36.02 && lng >= 139.97 && lng <= 140.05) return "守谷方面";
  if (lat >= 36.02 && lat <= 36.10 && lng >= 140.05 && lng <= 140.15) return "つくば方面";
  if (lat >= 35.95 && lat <= 36.02 && lng >= 140.12 && lng <= 140.20) return "牛久方面";

  return "";
}

function detectPrefecture(address = "") {
  const a = normalizeAddressText(address);
  if (!a) return "";
  if (a.includes("東京都")) return "東京";
  if (a.includes("千葉県")) return "千葉";
  if (a.includes("埼玉県")) return "埼玉";
  if (a.includes("茨城県")) return "茨城";
  return "";
}

function guessArea(lat, lng, address = "") {
  const byAddress = classifyAreaByAddress(address);
  if (byAddress) return byAddress;

  const genericByAddress = buildGenericAreaByAddress(address);
  if (genericByAddress) return genericByAddress;

  const byLatLng = classifyAreaByLatLng(lat, lng);
  if (byLatLng) return byLatLng;

  const pref = detectPrefectureGeneric(address) || detectPrefecture(address);
  if (pref) return `${pref}方面`;

  return "周辺";
}
