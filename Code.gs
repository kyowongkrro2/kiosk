/**
 * 학습개발2팀 키오스크 — 서버 코드 (최종본)
 * _data: 앱 원본 데이터(JSON). 절대 수정 금지. 탭 숨김 권장.
 * 대시보드/주문내역/축하카드/팀원/카페메뉴: 자동 정리되는 읽기 전용 시트.
 * 시트 직접 수정은 반영되지 않음. 모든 수정은 앱에서.
 *
 * 코드 수정 후에는 반드시:
 * 배포 → 배포 관리 → 연필 → 버전: "새 버전" → 배포
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "load") {
    return ContentService.createTextOutput(loadJson_())
      .setMimeType(ContentService.MimeType.TEXT);
  }
  return ContentService.createTextOutput("API 서버 동작 중입니다.");
}

function doPost(e) {
  saveJson_(e.postData.contents);
  return ContentService.createTextOutput("ok")
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── 데이터 저장 ────────────────────────────────────────
function ss_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("SS_ID");
  if (!id) {
    const ss = SpreadsheetApp.create("학습개발2팀 키오스크 데이터");
    id = ss.getId();
    props.setProperty("SS_ID", id);
  }
  return SpreadsheetApp.openById(id);
}

function sheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName("_data");
  if (!sh) sh = ss.insertSheet("_data");
  return sh;
}

const CHUNK = 40000;

function loadJson_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 1) return "";
  const values = sh.getRange(1, 1, last, 1).getValues();
  return values.map(function (r) { return r[0]; }).join("");
}

function saveJson_(json) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    sh.clearContents();
    const rows = [];
    for (let i = 0; i < json.length; i += CHUNK) {
      rows.push([json.substring(i, i + CHUNK)]);
    }
    if (rows.length > 0) sh.getRange(1, 1, rows.length, 1).setValues(rows);
    try { updateReadable_(json); } catch (err) { /* 정리 실패해도 저장은 유지 */ }
  } finally {
    lock.releaseLock();
  }
}

// ── 편집기에서 직접 실행하면 정리 시트를 즉시 갱신 ──────
function 수동갱신() {
  updateReadable_(loadJson_());
}

// ── 읽기 전용 시트 정리 ───────────────────────────────
function updateReadable_(json) {
  if (!json) return;
  let d;
  try { d = JSON.parse(json); } catch (err) { return; }
  const ss = ss_();

  const nameOf = function (id) {
    const m = (d.members || []).filter(function (x) { return x.id === id; })[0];
    return m ? m.name : id;
  };
  const cafeOf = function (id) {
    const c = (d.cafes || []).filter(function (x) { return x.id === id; })[0];
    return c ? c.name : "";
  };
  const fill = function (sheetName, rows) {
    let sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);
    sh.clearContents();
    sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sh.getRange(1, 1, 1, rows[0].length).setFontWeight("bold");
    sh.setFrozenRows(1);
  };
  const monthKeys = Object.keys(d.sessions || {}).sort();

  // 1) 대시보드
  const drows = [["월", "차수", "카페", "잔수", "금액"]];
  monthKeys.forEach(function (mk) {
    const s = d.sessions[mk];
    let mCups = 0, mAmt = 0;
    (s.rounds || []).forEach(function (r, i) {
      const cups = (r.orders || []).length;
      const amt = (r.orders || []).reduce(function (x, o) { return x + (o.price || 0); }, 0);
      mCups += cups; mAmt += amt;
      drows.push([mk, (i + 1) + "차", cafeOf(r.cafeId), cups, amt]);
    });
    drows.push([mk, "합계", "", mCups, mAmt]);
  });
  fill("대시보드", drows);

  // 2) 주문내역
  const orows = [["월", "차수", "카페", "이름", "상태", "메뉴", "옵션", "가격"]];
  monthKeys.forEach(function (mk) {
    const s = d.sessions[mk];
    (s.rounds || []).forEach(function (r, i) {
      (r.orders || []).forEach(function (o) {
        orows.push([mk, (i + 1) + "차", cafeOf(r.cafeId), nameOf(o.memberId),
          "주문", o.label, (o.options || []).join(", "), o.price]);
      });
      (r.absent || []).forEach(function (id) {
        orows.push([mk, (i + 1) + "차", cafeOf(r.cafeId), nameOf(id), "불참", "", "", ""]);
      });
    });
  });
  fill("주문내역", orows);

  // 3) 축하카드
  const crows = [["월", "내용", "작성시각"]];
  monthKeys.forEach(function (mk) {
    (d.sessions[mk].cards || []).forEach(function (c) {
      crows.push([mk, c.text, c.ts ? new Date(c.ts) : ""]);
    });
  });
  fill("축하카드", crows);

  // 4) 팀원
  const mrows = [["이름", "생일", "상태"]];
  (d.members || []).forEach(function (m) {
    mrows.push([m.name, m.bMonth ? m.bMonth + "/" + m.bDay : "", m.onLeave ? "휴직" : "재직"]);
  });
  fill("팀원", mrows);

  // 5) 카페메뉴
  const cfrows = [["카페", "카테고리", "메뉴", "가격", "온도"]];
  (d.cafes || []).forEach(function (c) {
    (c.menus || []).forEach(function (m) {
      cfrows.push([c.name, m.cat, m.name, m.price,
        m.temp === "both" ? "핫·아이스" : m.temp === "hot" ? "핫 전용" : "아이스 전용"]);
    });
    (c.options || []).forEach(function (o) {
      cfrows.push([c.name, "옵션", o.name, o.delta, ""]);
    });
  });
  fill("카페메뉴", cfrows);
}
