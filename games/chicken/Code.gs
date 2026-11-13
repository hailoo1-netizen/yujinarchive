/**
 * 사회 치킨 마을 창업 (학급별 연동)
 * 스프레드시트 연결형 Apps Script 웹앱 - 서버 코드
 *
 * 처음 한 번: 편집기에서 setup 함수를 실행하면 [설정]/[메뉴]/[주문] 시트가 만들어집니다.
 * 단계(설계 → 판매 → 종료)는 웹앱의 [교사용] 화면에서 학급별로 바꿉니다.
 */

const BASE_COST = 6000;
const SHEET_SETTINGS = '설정';
const SHEET_MENU = '메뉴';
const SHEET_ORDER = '주문';
const SHEET_STUDENT = '학생';
const PHASES = ['설계', '판매', '종료'];

/* ---------- 원가표 (PDF 기준) ---------- */
function items_(prefix, list) {
  return list.map(function (x, i) {
    const o = { id: prefix + (i + 1), name: x[0], cost: x[1] };
    if (x[2]) { o.custom = true; o.min = x[1]; o.ph = x[2]; o.prefix = x[3] || ''; }
    return o;
  });
}

const OPTION_GROUPS = [
  { id: 'form', no: '①', title: '닭의 형태', min: 1, max: 1, items: items_('fo', [
    ['뼈 한 마리', 0], ['윙봉만', 1000], ['순살', 1500], ['닭다리만', 1200]]) },
  { id: 'cook', no: '②', title: '조리 방법', min: 1, max: 1, items: items_('co', [
    ['기본 튀김', 1000], ['오븐구이', 1000], ['바삭 튀김', 1500], ['숯불구이', 2000]]) },
  { id: 'flavor', no: '③', title: '맛', min: 1, max: 2, half: true, items: items_('fl', [
    ['기본 후라이드', 0], ['기본 양념', 500], ['간장마늘', 500], ['매운 양념', 500],
    ['허니버터', 700], ['갈릭버터', 700], ['파닭 소스', 700], ['불닭맛', 700],
    ['고추마요', 700], ['데리야키', 700], ['카레', 700], ['바비큐맛', 700],
    ['레몬페퍼', 700], ['허브솔트', 700], ['청양간장', 800], ['마라맛', 1000],
    ['깐풍맛', 800], ['양파크림', 800], ['스위트칠리', 800], ['고르곤졸라 꿀맛', 900],
    ['치즈시즈닝', 1000], ['콘치즈맛', 1000], ['불고기맛', 1000], ['짜장맛', 1000],
    ['잡채맛', 1000], ['크림치즈맛', 1000], ['민트초코맛', 1200], ['트러플맛', 1200],
    ['로제맛', 1200], ['시금치 크림맛', 1200], ['바질페스토맛', 1200], ['토마토치즈맛', 1000],
    ['고구마치즈맛', 1200], ['초코맛', 1200], ['캐비어맛', 1200],
    ['우리 기업만의 소스맛', 1700, '소스 이름']]) },
  { id: 'topping', no: '④', title: '토핑', min: 0, max: 4, items: items_('tp', [
    ['파', 500], ['마늘칩', 500], ['치즈', 1000], ['떡', 500], ['구운 채소', 700],
    ['베이컨', 1000], ['우동면', 1000], ['우리 기업만의 토핑', 1200, '토핑 이름']]) },
  { id: 'side', no: '⑤', title: '추가 사이드 메뉴', min: 0, max: 2, items: items_('sd', [
    ['소스 2종', 500], ['치킨무', 300], ['감자튀김', 1500], ['치즈볼', 2000],
    ['떡볶이', 2000], ['샐러드', 1000], ['주먹밥', 1500],
    ['우리 기업만의 사이드 메뉴', 1500, '사이드 메뉴 이름']]) },
  { id: 'drink', no: '⑥', title: '음료', min: 0, max: 1, items: items_('dr', [
    ['콜라 500ml', 500], ['콜라 1.25L', 1000], ['제로 콜라 500ml', 700],
    ['우리 기업만의 음료', 1200, '음료 이름']]) },
  { id: 'pack', no: '⑦', title: '포장', min: 1, max: 1, items: items_('pk', [
    ['기본 종이상자', 300], ['보온 포장', 700], ['친환경 포장', 800], ['고급 포장', 1200]]) },
  { id: 'delivery', no: '⑧', title: '배달 방식', min: 1, max: 1, items: items_('dv', [
    ['포장 주문만', 0], ['일반 배달 (1시간)', 1000], ['빠른 배달 (40분)', 1500], ['고속 배달 (20분)', 2000]]) },
  { id: 'goods', no: '⑨', title: '구매 사은품 (굿즈)', min: 0, max: 1, items: items_('gd', [
    ['포토 카드', 1500, '인물 이름', '포토 카드'], ['클리커', 2000], ['열쇠 고리', 1500],
    ['우리 기업만의 굿즈', 1500, '굿즈 이름']]) }
];

/* ---------- 웹앱 ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('사회 치킨 마을 창업')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(SHEET_SETTINGS)) {
    const s = ss.insertSheet(SHEET_SETTINGS);
    s.getRange(1, 1, 5, 2).setValues([
      ['항목', '값'],
      ['학급 목록', '1반,2반,3반,4반'],
      ['모둠 수', 6],
      ['시작 자금', 35000],
      ['교사 비밀번호', '1234']
    ]);
    s.setColumnWidth(1, 140).setColumnWidth(2, 240);
  }
  if (!ss.getSheetByName(SHEET_MENU)) {
    ss.insertSheet(SHEET_MENU).appendRow(
      ['학급', '모둠', '메뉴명', '선택 옵션', '옵션데이터', '총 원가', '이윤', '판매 가격', '특징', '수정 시각', '그림']);
  }
  if (!ss.getSheetByName(SHEET_STUDENT)) {
    ss.insertSheet(SHEET_STUDENT).appendRow(['학급', '번호', '모둠']);
  }
  if (!ss.getSheetByName(SHEET_ORDER)) {
    ss.insertSheet(SHEET_ORDER).appendRow(
      ['시각', '학급', '구매 모둠', '판매 모둠', '메뉴명', '가격', '이윤', '판매 학급', '구매 번호']);
  }
}

function sheet_(name) {
  let s = SpreadsheetApp.getActive().getSheetByName(name);
  if (!s) { setup(); s = SpreadsheetApp.getActive().getSheetByName(name); }
  return s;
}

function getSettings_() {
  const rows = sheet_(SHEET_SETTINGS).getDataRange().getValues();
  const m = {};
  rows.forEach(function (r) { m[String(r[0]).trim()] = r[1]; });
  return {
    classes: String(m['학급 목록'] || '1반').split(',').map(function (x) { return x.trim(); }).filter(String),
    teams: Math.max(1, Number(m['모둠 수']) || 6),
    startMoney: Number(m['시작 자금']) || 35000,
    password: String(m['교사 비밀번호'] || '1234')
  };
}

function checkClassTeam_(cls, team) {
  const st = getSettings_();
  if (st.classes.indexOf(cls) < 0) throw new Error('학급을 다시 선택해 주세요.');
  if (team !== undefined) {
    team = Number(team);
    if (!(team >= 1 && team <= st.teams)) throw new Error('모둠을 다시 선택해 주세요.');
  }
  return st;
}

function getPhase_(cls) {
  return PropertiesService.getScriptProperties().getProperty('phase_' + cls) || '설계';
}

function safe_(s, max) {
  s = String(s == null ? '' : s).trim().slice(0, max || 100);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function won_(n) {
  return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '원';
}

/* ---------- 원가 계산 (서버에서 다시 검증) ---------- */
function computeCost_(sel) {
  let cost = BASE_COST;
  const lines = [];
  const clean = {};
  OPTION_GROUPS.forEach(function (g) {
    const picks = (sel && Array.isArray(sel[g.id])) ? sel[g.id] : [];
    if (picks.length < g.min) throw new Error(g.title + ' 항목을 ' + g.min + '개 골라 주세요.');
    if (picks.length > g.max) throw new Error(g.title + ' 항목은 최대 ' + g.max + '개까지 고를 수 있어요.');
    const seen = {};
    clean[g.id] = picks.map(function (p) {
      const it = g.items.filter(function (i) { return i.id === p.id; })[0];
      if (!it) throw new Error('원가표에 없는 옵션이 들어 있어요. 새로고침 후 다시 골라 주세요.');
      if (seen[it.id]) throw new Error(g.title + '에 같은 옵션이 두 번 들어 있어요.');
      seen[it.id] = true;
      let name = it.name, c = it.cost, detail = '';
      if (it.custom) {
        detail = String(p.detail || '').trim().slice(0, 20);
        if (!detail) throw new Error(it.name + '의 ' + it.ph + '을(를) 적어 주세요.');
        c = Math.round(Number(p.cost));
        if (!isFinite(c) || c < it.min) throw new Error(it.name + ' 원가는 ' + won_(it.min) + ' 이상이어야 해요.');
        if (c > 30000) throw new Error(it.name + ' 원가가 너무 커요. 선생님과 다시 협의해 주세요.');
        name = it.prefix ? it.prefix + '(' + detail + ')' : detail;
      }
      cost += c;
      return { id: it.id, detail: detail, cost: c, name: name, label: name + (it.custom ? '*' : '') };
    });
    const labels = clean[g.id].map(function (x) { return x.label; });
    clean[g.id].forEach(function (x) { delete x.label; });
    // 맛을 2개 고르면 반반 치킨: 두 맛의 원가를 모두 더함
    if (g.half && labels.length === 2) lines.push('반반(' + labels.join('+') + ')');
    else Array.prototype.push.apply(lines, labels);
  });
  return { cost: cost, summary: lines.join(', '), clean: clean };
}

/* ---------- 데이터 읽기 ---------- */
function readAllMenus_() {
  const rows = sheet_(SHEET_MENU).getDataRange().getValues();
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    if (!rows[r][0]) continue;
    let opts = {};
    try { opts = JSON.parse(rows[r][4] || '{}'); } catch (e) {}
    out.push({
      cls: String(rows[r][0]), team: Number(rows[r][1]), name: String(rows[r][2]), summary: String(rows[r][3]),
      options: opts, cost: Number(rows[r][5]), profit: Number(rows[r][6]),
      price: Number(rows[r][7]), feature: String(rows[r][8]),
      imageId: String(rows[r][10] || '')
    });
  }
  return out;
}

function readMenus_(cls) {
  return readAllMenus_().filter(function (m) { return m.cls === cls; })
    .sort(function (a, b) { return a.team - b.team; });
}

// 주문 시트: 시각, 구매 학급, 구매 모둠, 판매 모둠, 메뉴명, 가격, 이윤, 판매 학급(비어 있으면 같은 반), 구매 번호
function readOrders_() {
  const rows = sheet_(SHEET_ORDER).getDataRange().getValues();
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    if (!rows[r][1]) continue;
    out.push({ buyerCls: String(rows[r][1]), buyerTeam: Number(rows[r][2]), buyerNo: Number(rows[r][8]) || 0,
               sellerCls: String(rows[r][7] || rows[r][1]), seller: Number(rows[r][3]),
               name: String(rows[r][4]), price: Number(rows[r][5]), profit: Number(rows[r][6]) });
  }
  return out;
}

function crossOn_() {
  return PropertiesService.getScriptProperties().getProperty('crossClass') !== 'off';
}

function phases_(st) {
  const props = PropertiesService.getScriptProperties().getProperties();
  const out = {};
  st.classes.forEach(function (c) { out[c] = props['phase_' + c] || '설계'; });
  return out;
}

function market_(st) {
  const menus = readAllMenus_();
  const orders = readOrders_();
  const T = {}, P = {};
  function team(c, t) {
    const k = c + '|' + t;
    return T[k] || (T[k] = { sold: 0, revenue: 0, profit: 0, buyers: [] });
  }
  function person(c, n) {
    const k = c + '|' + n;
    return P[k] || (P[k] = { no: n, team: 0, spent: 0, bought: [] });
  }
  orders.forEach(function (o) {
    const s = team(o.sellerCls, o.seller);
    s.sold += 1; s.revenue += o.price; s.profit += o.profit;
    if (o.buyerNo) {
      s.buyers.push(o.buyerCls === o.sellerCls ? o.buyerNo + '번' : o.buyerCls + ' ' + o.buyerNo + '번');
      const b = person(o.buyerCls, o.buyerNo);
      b.spent += o.price; b.team = o.buyerTeam;
      b.bought.push((o.sellerCls !== o.buyerCls ? o.sellerCls + ' ' : '') + o.seller + '모둠 「' + o.name + '」');
    }
  });
  menus.forEach(function (m) { m.sold = team(m.cls, m.team).sold; });
  return {
    menus: menus,
    team: function (c, t) {
      const x = team(c, t);
      return { team: t, sold: x.sold, revenue: x.revenue, profit: x.profit,
               buyers: x.buyers.filter(function (v, i, a) { return a.indexOf(v) === i; }) };
    },
    person: function (c, n) {
      const x = person(c, n);
      return { no: n, team: x.team, spent: x.spent, bought: x.bought, balance: st.startMoney - x.spent };
    },
    people: function (c) {
      return Object.keys(P).filter(function (k) { return k.split('|')[0] === c; })
        .map(function (k) { return this.person(c, P[k].no); }, this)
        .sort(function (a, b) { return a.no - b.no; });
    }
  };
}

function checkNo_(no) {
  no = Number(no);
  if (!(Number.isInteger(no) && no >= 1 && no <= 99)) throw new Error('번호를 1~99 사이로 적어 주세요.');
  return no;
}

// 학생 번호를 처음 들어온 모둠에 묶어, 다른 모둠으로 들어가 자기 메뉴를 사는 일을 막습니다.
function registerStudent_(cls, no, team) {
  const s = sheet_(SHEET_STUDENT);
  const find = function () {
    const rows = s.getDataRange().getValues();
    for (let r = 1; r < rows.length; r++) {
      if (String(rows[r][0]) === cls && Number(rows[r][1]) === no) return Number(rows[r][2]);
    }
    return 0;
  };
  let t = find();
  if (!t) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      t = find();
      if (!t) { s.appendRow([cls, no, team]); t = team; }
    } finally { lock.releaseLock(); }
  }
  if (t !== team) throw new Error(cls + ' ' + no + '번은 ' + t + '모둠으로 등록되어 있어요. 내 모둠으로 들어와 주세요.');
}

/* ---------- 학생용 API ---------- */
function getInit() {
  const st = getSettings_();
  return { groups: OPTION_GROUPS, base: BASE_COST, classes: st.classes, teams: st.teams, startMoney: st.startMoney };
}

function getClassData(cls, team, no) {
  const st = checkClassTeam_(cls, team);
  team = Number(team);
  no = checkNo_(no);
  registerStudent_(cls, no, team);
  const ph = phases_(st);
  const cross = crossOn_();
  const mk = market_(st);
  // 우리 반 메뉴는 항상, 다른 반 메뉴는 '다른 반 구매'가 켜져 있고 그 반이 설계 단계를 마쳤을 때만
  const menus = mk.menus.filter(function (m) {
    if (st.classes.indexOf(m.cls) < 0) return false;
    return m.cls === cls || (cross && ph[m.cls] !== '설계');
  }).sort(function (a, b) {
    if (a.cls !== b.cls) {
      if (a.cls === cls) return -1;
      if (b.cls === cls) return 1;
      return st.classes.indexOf(a.cls) - st.classes.indexOf(b.cls);
    }
    return a.team - b.team;
  });
  return { phase: ph[cls], cross: cross, startMoney: st.startMoney, menus: menus,
           me: mk.team(cls, team), wallet: mk.person(cls, no) };
}

function saveMenu(cls, team, payload, no) {
  checkClassTeam_(cls, team);
  team = Number(team);
  const phase = getPhase_(cls);
  if (phase !== '설계') throw new Error('지금은 ' + phase + ' 단계라서 메뉴를 바꿀 수 없어요.');
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('메뉴명을 적어 주세요.');
  const profit = Math.round(Number(payload.profit));
  if (!isFinite(profit) || profit < 0) throw new Error('이윤은 0원 이상으로 적어 주세요.');
  if (profit > 100000) throw new Error('이윤이 너무 커요. 10만 원 이하로 정해 주세요.');
  const c = computeCost_(payload.selections);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_(SHEET_MENU);
    const row = [cls, team, safe_(name, 30), c.summary, JSON.stringify(c.clean), c.cost, profit,
                 c.cost + profit, safe_(payload.feature, 200), new Date()];
    const rows = s.getDataRange().getValues();
    let found = false;
    for (let r = 1; r < rows.length; r++) {
      if (String(rows[r][0]) === cls && Number(rows[r][1]) === team) {
        s.getRange(r + 1, 1, 1, row.length).setValues([row]);
        found = true; break;
      }
    }
    if (!found) s.appendRow(row);
  } finally { lock.releaseLock(); }
  return getClassData(cls, team, no);
}


function buy(cls, team, no, sellerCls, seller) {
  const st = checkClassTeam_(cls, team);
  checkClassTeam_(sellerCls, seller);
  team = Number(team); seller = Number(seller); no = checkNo_(no);
  if (cls === sellerCls && team === seller) throw new Error('우리 모둠 치킨은 살 수 없어요.');
  registerStudent_(cls, no, team);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ph = phases_(st);
    if (ph[cls] !== '판매') throw new Error('지금은 판매 라운드가 아니에요.');
    if (sellerCls !== cls) {
      if (!crossOn_()) throw new Error('지금은 다른 반 치킨을 살 수 없어요.');
      if (ph[sellerCls] === '설계') throw new Error(sellerCls + '은(는) 아직 메뉴를 만드는 중이에요.');
    }
    const mk = market_(st);
    const menu = mk.menus.filter(function (m) { return m.cls === sellerCls && m.team === seller; })[0];
    if (!menu) throw new Error(sellerCls + ' ' + seller + '모둠은 아직 메뉴가 없어요.');
    if (mk.person(cls, no).balance < menu.price) throw new Error('내 지갑 잔액이 부족해요.');
    const s = sheet_(SHEET_ORDER);
    if (!s.getRange(1, 8).getValue()) s.getRange(1, 8).setValue('판매 학급');
    if (!s.getRange(1, 9).getValue()) s.getRange(1, 9).setValue('구매 번호');
    s.appendRow([new Date(), cls, team, seller, menu.name, menu.price, menu.profit, sellerCls, no]);
  } finally { lock.releaseLock(); }
  return getClassData(cls, team, no);
}

/* ---------- 메뉴 그림 ---------- */
const MAX_IMAGE_BYTES = 2 * 1024 * 1024;

function imageFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('imageFolderId');
  if (id) {
    try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {}
  }
  const folder = DriveApp.createFolder('치킨 마을 메뉴 그림');
  props.setProperty('imageFolderId', folder.getId());
  return folder;
}

function trashFile_(id) {
  if (!id) return;
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
}

function uploadMenuImage(cls, team, dataUrl, no) {
  checkClassTeam_(cls, team);
  team = Number(team);
  if (getPhase_(cls) === '종료') throw new Error('판매가 끝나서 그림을 바꿀 수 없어요.');
  const m = /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+\/=]+)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('그림 파일(jpg, png)만 올릴 수 있어요.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > MAX_IMAGE_BYTES) throw new Error('그림 용량이 너무 커요. 더 작은 사진으로 올려 주세요.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_(SHEET_MENU);
    if (!s.getRange(1, 11).getValue()) s.getRange(1, 11).setValue('그림');
    const rows = s.getDataRange().getValues();
    let rowNo = 0;
    for (let r = 1; r < rows.length; r++) {
      if (String(rows[r][0]) === cls && Number(rows[r][1]) === team) { rowNo = r + 1; break; }
    }
    if (!rowNo) throw new Error('메뉴를 먼저 저장한 뒤 그림을 올려 주세요.');
    const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
    const blob = Utilities.newBlob(bytes, m[1], cls + '_' + team + '모둠_메뉴그림.' + ext);
    const file = imageFolder_().createFile(blob);
    trashFile_(String(rows[rowNo - 1][10] || ''));
    s.getRange(rowNo, 11).setValue(file.getId());
  } finally { lock.releaseLock(); }
  return getClassData(cls, team, no);
}

function getMenuImage(cls, imageId) {
  checkClassTeam_(cls);
  const ok = readAllMenus_().some(function (m) { return m.imageId && m.imageId === imageId; });
  if (!ok) throw new Error('그림을 찾을 수 없어요.');
  const blob = DriveApp.getFileById(imageId).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

/* ---------- 교사용 API ---------- */
function checkPw_(pw) {
  if (String(pw) !== getSettings_().password) throw new Error('교사 비밀번호가 맞지 않아요.');
}


function teacherGetClass(pw, cls) {
  checkPw_(pw);
  const st = checkClassTeam_(cls);
  const mk = market_(st);
  const rows = [];
  for (let t = 1; t <= st.teams; t++) {
    const m = mk.menus.filter(function (x) { return x.cls === cls && x.team === t; })[0] || null;
    rows.push({ team: t, menu: m, stat: mk.team(cls, t) });
  }
  return { phase: phases_(st)[cls], cross: crossOn_(), startMoney: st.startMoney, rows: rows, people: mk.people(cls) };
}

function teacherSetPhase(pw, cls, phase) {
  checkPw_(pw);
  checkClassTeam_(cls);
  if (PHASES.indexOf(phase) < 0) throw new Error('알 수 없는 단계예요.');
  PropertiesService.getScriptProperties().setProperty('phase_' + cls, phase);
  return teacherGetClass(pw, cls);
}

function teacherSetCross(pw, cls, on) {
  checkPw_(pw);
  PropertiesService.getScriptProperties().setProperty('crossClass', on ? 'on' : 'off');
  return teacherGetClass(pw, cls);
}

// 이 학급의 메뉴·그림과, 이 학급이 사거나 판 주문(다른 반과의 거래 포함)을 지웁니다.
function teacherReset(pw, cls) {
  checkPw_(pw);
  checkClassTeam_(cls);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ms = sheet_(SHEET_MENU);
    let rows = ms.getDataRange().getValues();
    for (let r = rows.length - 1; r >= 1; r--) {
      if (String(rows[r][0]) === cls) { trashFile_(String(rows[r][10] || '')); ms.deleteRow(r + 1); }
    }
    const os = sheet_(SHEET_ORDER);
    rows = os.getDataRange().getValues();
    for (let r = rows.length - 1; r >= 1; r--) {
      const buyerCls = String(rows[r][1]), sellerCls = String(rows[r][7] || rows[r][1]);
      if (buyerCls === cls || sellerCls === cls) os.deleteRow(r + 1);
    }
    const ss2 = sheet_(SHEET_STUDENT);
    rows = ss2.getDataRange().getValues();
    for (let r = rows.length - 1; r >= 1; r--) {
      if (String(rows[r][0]) === cls) ss2.deleteRow(r + 1);
    }
    PropertiesService.getScriptProperties().setProperty('phase_' + cls, '설계');
  } finally { lock.releaseLock(); }
  return teacherGetClass(pw, cls);
}
