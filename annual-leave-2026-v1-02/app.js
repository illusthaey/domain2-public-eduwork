
'use strict';

/**
 * 2026학년도(2026-03-01 부여) 전용 연차유급휴가 부여일수 계산기 MVP
 * - 기준기간: 2025-03-01 ~ 2026-02-28
 * - 처리: 브라우저 로컬(서버 전송 없음)
 *
 * 의도적으로 “연도 가변화”를 하지 않습니다(일회성 배포).
 */

const FIXED = {
  workStart: '2025-03-01',
  workEnd: '2026-02-28',
  grantDate: '2026-03-01',
};

const CATEGORY = {
  DEEMED: 'DEEMED',      // 출근간주(분자 유지)
  EXCLUDE: 'EXCLUDE',    // 산정제외(재산정에서 분모/분자 제거)
  ABSENCE: 'ABSENCE',    // 결근성(분자 차감, 개근월수 깨짐)
  REVIEW: 'REVIEW',      // 검토필요(자동 산정 미반영 또는 보수적 처리)
};

let RULES = null;
let LAST_RESULT = null;
let STATE = null; // 파싱된 원천데이터 + 수동보정 상태(재계산용)

/* -------------------------
 * DOM
 * ------------------------- */
const el = (id) => document.getElementById(id);

const statusBox = el('status');
const summaryBox = el('summary');
const resultsBox = el('results');
const detailsBox = el('details');

function logStatus(msg) {
  statusBox.textContent = (statusBox.textContent ? statusBox.textContent + '\n' : '') + msg;
}

function resetUI() {
  statusBox.textContent = '';
  summaryBox.innerHTML = '';
  resultsBox.innerHTML = '';
  detailsBox.innerHTML = '';
  el('btnDownload').disabled = true;
  LAST_RESULT = null;
}

/* -------------------------
 * Utils: Date
 * ------------------------- */
function parseISODate(s) {
  // s: 'YYYY-MM-DD'
  if (!s) return null;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s).trim());
  if (!m) return null;
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  const dt = new Date(Date.UTC(y, mo - 1, d));
  // Validate roundtrip
  if (dt.getUTCFullYear() !== y || (dt.getUTCMonth() + 1) !== mo || dt.getUTCDate() !== d) return null;
  return dt;
}

function formatISODate(dt) {
  const y = dt.getUTCFullYear();
  const m = String(dt.getUTCMonth() + 1).padStart(2, '0');
  const d = String(dt.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function addDaysUTC(dt, days) {
  const r = new Date(dt.getTime());
  r.setUTCDate(r.getUTCDate() + days);
  return r;
}

function isSaturdayUTC(dt) {
  // JS: 0=Sun ... 6=Sat (UTC)
  return dt.getUTCDay() === 6;
}

function clampRangeToFixed(startDt, endDt) {
  const fixedStart = parseISODate(FIXED.workStart);
  const fixedEnd = parseISODate(FIXED.workEnd);
  if (!startDt || !endDt) return null;

  const s = startDt < fixedStart ? fixedStart : startDt;
  const e = endDt > fixedEnd ? fixedEnd : endDt;
  if (s > e) return null;
  return { start: s, end: e };
}

function iterateDaysExclSatUTC(startDt, endDt, fn) {
  let d = startDt;
  while (d <= endDt) {
    if (!isSaturdayUTC(d)) fn(d);
    d = addDaysUTC(d, 1);
  }
}

function buildDateSetExclSat(startDt, endDt) {
  const set = new Set();
  if (!startDt || !endDt) return set;
  iterateDaysExclSatUTC(startDt, endDt, (d) => set.add(formatISODate(d)));
  return set;
}

function countDaysExclSat(startDt, endDt) {
  let cnt = 0;
  iterateDaysExclSatUTC(startDt, endDt, () => cnt++);
  return cnt;
}

/* -------------------------
 * Utils: Parsing
 * ------------------------- */
function normalizeHeader(h) {
  return String(h ?? '')
    .replace(/\r?\n/g, '')
    .replace(/\s+/g, '')
    .trim();
}

function parseNameAndPersonalNo(v) {
  // "고남향\r\n(K109050178)" -> { name, personalNo }
  const s = String(v ?? '').trim();
  const name = s.split(/\r?\n/)[0].trim();
  const m = /\((K\d+)\)/.exec(s);
  return {
    name: name || null,
    personalNo: m ? m[1] : null,
    raw: s,
  };
}

function parsePeriodRange(v) {
  // "2026-02-05 09:00 ~ 2026-02-05 18:00"
  const s = String(v ?? '').trim();
  const parts = s.split('~').map(x => x.trim());
  if (parts.length !== 2) return null;
  const left = parts[0].slice(0, 10);
  const right = parts[1].slice(0, 10);
  const start = parseISODate(left);
  const end = parseISODate(right);
  if (!start || !end) return null;
  return { start, end };
}

function parseDaysHoursMinutes(v) {
  // "6일 0시간 30분", "0일 6시간 30분"
  const s = String(v ?? '').trim();
  const day = /(\d+)\s*일/.exec(s);
  const hour = /(\d+)\s*시간/.exec(s);
  const min = /(\d+)\s*분/.exec(s);
  return {
    days: day ? Number(day[1]) : 0,
    hours: hour ? Number(hour[1]) : 0,
    minutes: min ? Number(min[1]) : 0,
    raw: s,
  };
}

function parseWeeklyMinutes(v) {
  // "40시간00분", "14시간00분", "20시간00분"
  const s = String(v ?? '').trim();
  const hour = /(\d+)\s*시간/.exec(s);
  const min = /(\d+)\s*분/.exec(s);
  const h = hour ? Number(hour[1]) : 0;
  const m = min ? Number(min[1]) : 0;
  return h * 60 + m;
}

function safeText(v) {
  return String(v ?? '').trim();
}

/* -------------------------
 * Utils: Record ID (수동보정 키)
 * ------------------------- */
function fnv1a32(str) {
  // 32-bit FNV-1a
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    // multiply by prime 16777619 (overflow 32-bit)
    h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
  }
  return h >>> 0;
}

function makeRecordId(parts) {
  const s = parts.map(x => String(x ?? '')).join('|');
  return 'R' + fnv1a32(s).toString(36);
}

/* -------------------------
 * XLSX/CSV 읽기
 * ------------------------- */
async function readAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error);
    reader.onload = () => resolve(reader.result);
    reader.readAsArrayBuffer(file);
  });
}

async function readWorkbook(file) {
  if (!window.XLSX) {
    throw new Error('XLSX 라이브러리가 로드되지 않았습니다. xlsx.full.min.js를 포함하거나 index.html의 script src를 수정하세요.');
  }
  const buf = await readAsArrayBuffer(file);
  return XLSX.read(buf, { type: 'array' });
}

function sheetToObjects(sheet) {
  // header:1 방식으로 헤더 정규화
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (!rows || rows.length < 2) return [];
  const header = rows[0].map(normalizeHeader);
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    // 총계행 제거(근무상황목록)
    if (String(row[0] ?? '').includes('총계')) continue;
    const obj = {};
    for (let c = 0; c < header.length; c++) {
      const key = header[c] || `COL${c}`;
      obj[key] = row[c];
    }
    out.push(obj);
  }
  return out;
}

/* -------------------------
 * 규칙 로드/분류
 * ------------------------- */

async function loadRules() {
  // 1) rules_2026.json fetch (정적 서버 배포 시)
  try {
    const res = await fetch('rules_2026.json', { cache: 'no-store' });
    if (res.ok) {
      RULES = await res.json();
      el('appVersion').textContent = RULES?.app?.version ?? '-';
      return;
    }
  } catch (e) {
    // ignore and fallback
  }

  // 2) index.html 내장 JSON (file:// 실행 대응)
  const embedded = document.getElementById('rulesJson');
  if (embedded && embedded.textContent) {
    RULES = JSON.parse(embedded.textContent);
    el('appVersion').textContent = RULES?.app?.version ?? '-';
    return;
  }

  throw new Error('규칙(rules_2026.json) 로드 실패');
}


function classifyRecord(kind, reason) {
  const k = safeText(kind);
  const r = safeText(reason);

  // 1) excluded_kinds / absence_like_kinds 최우선
  if (RULES.classification.excluded_kinds.includes(k)) {
    return { category: CATEGORY.EXCLUDE, vacationCredit: false, tags: ['excluded_kind'] };
  }
  if (RULES.classification.absence_like_kinds.includes(k)) {
    return { category: CATEGORY.ABSENCE, vacationCredit: false, tags: ['absence_like_kind'] };
  }

  // 2) reason_rules (kind + reason)
  for (const rr of RULES.classification.reason_rules) {
    if (safeText(rr.when?.kind) !== k) continue;
    const includes = rr.when?.reason_includes_any ?? [];
    if (includes.some(word => r.includes(word))) {
      return {
        category: rr.category ?? CATEGORY.REVIEW,
        vacationCredit: Boolean(rr.vacation_credit),
        tags: rr.tags ?? [],
      };
    }
  }

  // 3) kind_rules
  for (const kr of RULES.classification.kind_rules) {
    if (safeText(kr.kind) === k) {
      return {
        category: kr.category ?? CATEGORY.REVIEW,
        vacationCredit: Boolean(kr.vacation_credit),
        tags: kr.tags ?? [],
        caps: kr.caps ?? null,
      };
    }
  }

  // 4) default
  return { category: CATEGORY.REVIEW, vacationCredit: false, tags: ['unknown_kind'] };
}

/* -------------------------
 * 인사기록 파싱
 * ------------------------- */
function extractSeniorityBaseDate(row) {
  // 후보 컬럼 중 가장 이른 날짜를 선택
  const candidates = [
    safeText(row['현소속교육청근무일']),
    safeText(row['최초계약일']),
    safeText(row['근무시작일']),
    safeText(row['재계약일']),
  ].filter(Boolean);

  const dates = candidates.map(parseISODate).filter(Boolean);
  if (dates.length === 0) return null;
  dates.sort((a,b)=>a-b);
  return dates[0];
}

function computeSeniorityYears(baseDate, asOfDate) {
  // 완전 연수(YYYY-MM-DD 기준)
  if (!baseDate || !asOfDate) return null;
  let years = asOfDate.getUTCFullYear() - baseDate.getUTCFullYear();
  const anniv = new Date(Date.UTC(asOfDate.getUTCFullYear(), baseDate.getUTCMonth(), baseDate.getUTCDate()));
  if (asOfDate < anniv) years -= 1;
  return Math.max(0, years);
}

function computeAddDaysBySeniorityYears(years) {
  // 최초 1년 초과 매 2년마다 1일
  if (years == null) return 0;
  const over1 = Math.max(0, years - 1);
  return Math.floor(over1 / 2);
}

function deriveWorkerCategory(hrRow) {
  const type = safeText(hrRow['직종구분']);     // 교육공무직/특수운영직/...
  const workForm = safeText(hrRow['근무형태']); // 상시근무자/방학중비근무자/...
  const workType = safeText(hrRow['근무유형']); // 전일제/시간제

  const isSpecial = (type === '특수운영직');
  const isEmergency = (workForm.includes('방학중') && workForm.includes('비근무')) || (workForm.includes('단시간방학중비근무자'));
  const isRegular = (workForm === '상시근무자');

  return {
    isSpecial,
    isEmergency,
    isRegular,
    hr_type: type,
    hr_workForm: workForm,
    hr_workType: workType,
  };
}

function parseHrPersons(rows) {
  const grantDate = parseISODate(FIXED.grantDate);

  const persons = [];
  for (const r of rows) {
    const personalNo = safeText(r['개인번호']) || null;
    const name = safeText(r['성명']) || null;
    if (!name) continue;

    const weeklyMinutes = parseWeeklyMinutes(r['주소정근로시간']);
    const workdaysPerWeek = Number(el('workdaysPerWeek').value || RULES.attendance.default_workdays_per_week);
    const dailyMinutes = workdaysPerWeek > 0 ? Math.round(weeklyMinutes / workdaysPerWeek) : 480;

    const baseDate = extractSeniorityBaseDate(r);
    const seniorityYears = computeSeniorityYears(baseDate, grantDate);
    const addDays = computeAddDaysBySeniorityYears(seniorityYears);

    const cat = deriveWorkerCategory(r);

    persons.push({
      key: personalNo || name, // join key
      personalNo,
      name,
      job: safeText(r['직종']),
      jobGroup: safeText(r['직종구분']),
      workForm: safeText(r['근무형태']),
      workType: safeText(r['근무유형']),
      weeklyMinutes,
      dailyMinutes,
      baseDate: baseDate ? formatISODate(baseDate) : null,
      seniorityYears,
      addDays,
      category: cat,
      raw: r,
    });
  }
  return persons;
}

/* -------------------------
 * 근무상황목록 파싱
 * ------------------------- */

function parseWorkRecords(rows) {
  const records = [];
  for (const r of rows) {
    const nameInfo = parseNameAndPersonalNo(r['성명']);
    const period = parsePeriodRange(r['기간']);
    if (!nameInfo.name || !period) continue;

    const kind = safeText(r['종별']);
    const reason = safeText(r['사유또는용무']); // "사유 또는 용무" -> normalizeHeader => "사유또는용무"
    const dur = parseDaysHoursMinutes(r['일수/기간']);
    const approval = safeText(r['결재상태']);

    const startISO = formatISODate(period.start);
    const endISO = formatISODate(period.end);

    const autoCls = classifyRecord(kind, reason);
    const id = makeRecordId([
      nameInfo.personalNo || nameInfo.name,
      startISO, endISO,
      kind, reason,
      dur.raw,
      approval,
    ]);

    records.push({
      id,
      key: nameInfo.personalNo || nameInfo.name,
      personalNo: nameInfo.personalNo,
      name: nameInfo.name,
      kind,
      reason,
      start: startISO,
      end: endISO,
      duration: dur, // days/hours/min
      approval,
      autoCls,
      cls: { ...autoCls }, // override 적용 후 cls를 갱신
      override: null,
      raw: r,
    });
  }
  return records;
}


/* -------------------------
 * 기록 분 단위로 일자별 분배
 * ------------------------- */
function allocateRecordToDateMinutes(rec, dailyMinutes) {
  // returns [{dateKey, minutes}]
  const period = clampRangeToFixed(parseISODate(rec.start), parseISODate(rec.end));
  if (!period) return [];

  // 날짜 목록(토 제외)
  const dates = [];
  iterateDaysExclSatUTC(period.start, period.end, (d) => dates.push(formatISODate(d)));
  if (dates.length === 0) return [];

  // 총 minutes 계산(일수 파트는 dailyMinutes 기준)
  const total = (rec.duration.days * dailyMinutes) + (rec.duration.hours * 60) + rec.duration.minutes;
  if (total <= 0) return [];

  // 단일일/다일 모두 순차 배분(일자당 dailyMinutes 상한)
  let remaining = total;
  const out = [];
  for (const dk of dates) {
    if (remaining <= 0) break;
    const m = Math.min(dailyMinutes, remaining);
    out.push({ dateKey: dk, minutes: m });
    remaining -= m;
  }

  // 남은 분이 있으면(입력 불일치) 마지막 일자에 합산(상한 초과 가능)하되 검토 대상으로 표시 가능
  if (remaining > 0 && out.length > 0) {
    out[out.length - 1].minutes += remaining;
  }
  return out;
}

/* -------------------------
 * 캘린더(학기/방학) 구성
 * ------------------------- */
function buildCalendar() {
  const fixedStart = parseISODate(FIXED.workStart);
  const fixedEnd = parseISODate(FIXED.workEnd);

  const yearDays = buildDateSetExclSat(fixedStart, fixedEnd);

  const summerStart = parseISODate(el('summerStart').value);
  const summerEnd = parseISODate(el('summerEnd').value);
  const winterStart = parseISODate(el('winterStart').value);
  const winterEnd = parseISODate(el('winterEnd').value);

  const vacationDays = new Set();
  const addRangeToSet = (s, e) => {
    if (!s || !e) return;
    const clamped = clampRangeToFixed(s, e);
    if (!clamped) return;
    iterateDaysExclSatUTC(clamped.start, clamped.end, (d) => vacationDays.add(formatISODate(d)));
  };
  addRangeToSet(summerStart, summerEnd);
  addRangeToSet(winterStart, winterEnd);

  const semesterDays = new Set([...yearDays].filter(d => !vacationDays.has(d)));

  // 참고 정보(연간 토요일 수/일수)
  const totalCalendarDays = (() => {
    const s = fixedStart, e = fixedEnd;
    return Math.round((e - s) / (24 * 3600 * 1000)) + 1;
  })();
  const satCount = (() => {
    let cnt = 0;
    let d = fixedStart;
    while (d <= fixedEnd) {
      if (isSaturdayUTC(d)) cnt++;
      d = addDaysUTC(d, 1);
    }
    return cnt;
  })();

  return {
    fixedStart: formatISODate(fixedStart),
    fixedEnd: formatISODate(fixedEnd),
    yearDays,
    semesterDays,
    vacationDays,
    info: { totalCalendarDays, satCount, yearDaysExclSat: yearDays.size, semesterDaysExclSat: semesterDays.size, vacationDaysExclSat: vacationDays.size }
  };
}

/* -------------------------
 * 출근율 계산(분 단위)
 * ------------------------- */
function sumMapValues(m) {
  let s = 0;
  for (const v of m.values()) s += v;
  return s;
}

function addMinutes(map, dateKey, minutes) {
  map.set(dateKey, (map.get(dateKey) || 0) + minutes);
}

function computeAttendanceForPerson(person, records, calendar, settings) {
  const dailyMinutes = person.dailyMinutes || 480;

  const excluded = new Map();      // dateKey -> minutes
  const absence = new Map();       // dateKey -> minutes
  const vacCredit = new Map();     // dateKey -> minutes
  const review = [];              // record refs

  const parentalCapDays = Number(settings.parentalDeemedCapDays || 365);
  let parentalUsedMinutes = 0; // MVP: 육아휴직 종류가 발견되면, cap 적용 필요(추후 룰 확장)

  for (const rec of records) {
    const alloc = allocateRecordToDateMinutes(rec, dailyMinutes);
    if (alloc.length === 0) continue;

    // 결재상태 미완결은 무조건 review에 올림(수치는 보수적으로 미반영)
    const approval = safeText(rec.approval);
    const isFinal = (approval === '완결');

    if (!isFinal) {
      review.push({ ...rec, note: '결재상태 미완결: 자동 산정 미반영(검토 필요)' });
      continue;
    }

    // MVP: 육아휴직 등 복잡 항목은 kind로 잡히면 추후 처리
    // 여기서는 RULES.classification.excluded_kinds / absence_like_kinds / kind_rules 중심으로 처리

    if (rec.cls.category === CATEGORY.EXCLUDE) {
      for (const { dateKey, minutes } of alloc) addMinutes(excluded, dateKey, minutes);
    } else if (rec.cls.category === CATEGORY.ABSENCE) {
      for (const { dateKey, minutes } of alloc) addMinutes(absence, dateKey, minutes);
    } else if (rec.cls.category === CATEGORY.DEEMED) {
      // 기본 출근 가정 모델에서는 변화 없음.
      // 단, 방학 중 근무 크레딧(vacation_credit)은 별도 반영
      if (rec.cls.vacationCredit) {
        for (const { dateKey, minutes } of alloc) {
          if (calendar.vacationDays.has(dateKey)) addMinutes(vacCredit, dateKey, minutes);
        }
      }
    } else {
      // REVIEW: 수치 미반영 + 검토 목록에 포함
      review.push({ ...rec, note: rec.cls.tags?.includes('unknown_kind') ? '종별 미분류: 검토 필요' : '기타(검토 필요)' });
    }
  }

  // 분모 구성
  const denomFull = calendar.yearDays.size * dailyMinutes;

  const denomSemester = calendar.semesterDays.size * dailyMinutes;

  // 제외/결근 분 계산(기간별)
  const sumInSet = (map, set) => {
    let s = 0;
    for (const [dk, min] of map.entries()) if (set.has(dk)) s += min;
    return s;
  };

  const excludedFull = sumMapValues(excluded);
  const absenceFull = sumMapValues(absence);

  const excludedSemester = sumInSet(excluded, calendar.semesterDays);
  const absenceSemester = sumInSet(absence, calendar.semesterDays);

  const vacCreditMinutes = sumInSet(vacCredit, calendar.vacationDays);

  const isEmergency = person.category.isEmergency;
  const isRegular = person.category.isRegular;

  // 상시 기준(연간): 상시근로자 = (분모 전체) - 제외/결근
  // 방학중비상시의 상시 기준(연간): 분모 전체, 분자 = 학기출근분 + 방학중유급근무분
  const numeratorFullRaw = isEmergency
    ? Math.max(0, (denomSemester - excludedSemester - absenceSemester)) + vacCreditMinutes
    : Math.max(0, denomFull - excludedFull - absenceFull);

  const rateFullRaw = denomFull > 0 ? (numeratorFullRaw / denomFull) : 0;

  // 재산정(제외기간 제거) - 상시 기준(상시근로자용)
  const denomFullRecalc = Math.max(0, denomFull - excludedFull);
  const numeratorFullRecalc = Math.max(0, denomFullRecalc - absenceFull);
  const rateFullRecalc = denomFullRecalc > 0 ? (numeratorFullRecalc / denomFullRecalc) : 0;

  // 비상시 기준(학기): 방학 제외
  const numeratorSemesterRaw = Math.max(0, denomSemester - excludedSemester - absenceSemester);
  const rateSemesterRaw = denomSemester > 0 ? (numeratorSemesterRaw / denomSemester) : 0;

  const denomSemesterRecalc = Math.max(0, denomSemester - excludedSemester);
  const numeratorSemesterRecalc = Math.max(0, denomSemesterRecalc - absenceSemester);
  const rateSemesterRecalc = denomSemesterRecalc > 0 ? (numeratorSemesterRecalc / denomSemesterRecalc) : 0;

  return {
    dailyMinutes,
    denomFull,
    denomSemester,
    excludedFull,
    excludedSemester,
    absenceFull,
    absenceSemester,
    vacCreditMinutes,
    numeratorFullRaw,
    rateFullRaw,
    rateFullRecalc,
    numeratorSemesterRaw,
    rateSemesterRaw,
    rateSemesterRecalc,
    reviewRecords: review,
    // 원천 기록 집계용
    maps: { excluded, absence, vacCredit },
  };
}

/* -------------------------
 * 연차 부여일수 산정(2026학년도 전용)
 * ------------------------- */
function determineBaseDays(person, att) {
  // 근로형태/직군별 기본일수(80% 충족 시)
  // - 상시근무자: 15
  // - 방학중비근무자(교육공무직): 12, 단 상시기준(연간) 80% 이상이면 15로 전환
  // - 방학중비근무자(특수운영직): 11, 단 상시기준 80% 이상이면 15로 전환(요구사항 반영)
  const isSpecial = person.category.isSpecial;
  const isEmergency = person.category.isEmergency;

  if (!isEmergency) return 15;

  if (!isSpecial) {
    // 교육공무직(비상시)
    if (att.rateFullRaw >= 0.8) return 15;
    return 12;
  } else {
    // 특수운영직군(비상시)
    if (att.rateFullRaw >= 0.8) return 15;
    return 11;
  }
}

function roundProrationToDaysAndHours(valueDays, dailyMinutes) {
  // 요구사항: 소수점 둘째 자리에서 반올림하여 시간 단위까지 부여
  // 구현: (1) 소수 첫째 자리까지 반올림 (2) 나머지를 "1일 소정근로시간" 기준으로 시간 환산
  const v = Math.round(valueDays * 10) / 10; // 0.1일 단위
  const days = Math.floor(v);
  const frac = v - days;
  const hoursPerDay = Math.round(dailyMinutes / 60); // 전일제 8
  let hours = Math.round(frac * hoursPerDay);
  let finalDays = days;
  if (hours >= hoursPerDay) {
    finalDays += 1;
    hours = 0;
  }
  return { days: finalDays, hours, roundedDays: v };
}

function computeMonthlyPerfectUnits(person, calendar, att) {
  // 80% 미달 시: 1개월 개근=1일
  // 상시: 12개월(2025.3~2026.2)
  // 비상시: 여름(7~8) 1개월, 겨울(12~2) 1개월로 묶어 9단위(3~6, 여름, 9~11, 겨울)
  const isEmergency = person.category.isEmergency;

  const fixedStart = parseISODate(FIXED.workStart);
  const fixedEnd = parseISODate(FIXED.workEnd);

  const units = []; // { label, start, end, daySet }
  const mkUnit = (label, start, end) => {
    const clamped = clampRangeToFixed(start, end);
    if (!clamped) return;
    const daySet = buildDateSetExclSat(clamped.start, clamped.end);
    units.push({ label, start: formatISODate(clamped.start), end: formatISODate(clamped.end), daySet });
  };

  if (!isEmergency) {
    // 12개월 단위
    for (let m = 3; m <= 12; m++) {
      const y = 2025;
      const start = new Date(Date.UTC(y, m - 1, 1));
      const end = new Date(Date.UTC(y, m, 0)); // last day
      mkUnit(`${y}-${String(m).padStart(2,'0')}`, start, end);
    }
    for (let m = 1; m <= 2; m++) {
      const y = 2026;
      const start = new Date(Date.UTC(y, m - 1, 1));
      const end = new Date(Date.UTC(y, m, 0));
      mkUnit(`${y}-${String(m).padStart(2,'0')}`, start, end);
    }
  } else {
    // 9단위
    mkUnit('2025-03', new Date(Date.UTC(2025,2,1)), new Date(Date.UTC(2025,2,31)));
    mkUnit('2025-04', new Date(Date.UTC(2025,3,1)), new Date(Date.UTC(2025,3,30)));
    mkUnit('2025-05', new Date(Date.UTC(2025,4,1)), new Date(Date.UTC(2025,4,31)));
    mkUnit('2025-06', new Date(Date.UTC(2025,5,1)), new Date(Date.UTC(2025,5,30)));

    // 여름방학: 사용자가 입력한 범위를 그대로 한 단위로 사용
    const ss = parseISODate(el('summerStart').value);
    const se = parseISODate(el('summerEnd').value);
    if (ss && se) mkUnit('여름방학(7~8월)', ss, se);

    mkUnit('2025-09', new Date(Date.UTC(2025,8,1)), new Date(Date.UTC(2025,8,30)));
    mkUnit('2025-10', new Date(Date.UTC(2025,9,1)), new Date(Date.UTC(2025,9,31)));
    mkUnit('2025-11', new Date(Date.UTC(2025,10,1)), new Date(Date.UTC(2025,10,30)));

    const ws = parseISODate(el('winterStart').value);
    const we = parseISODate(el('winterEnd').value);
    if (ws && we) mkUnit('겨울방학(12~2월)', ws, we);
  }

  // 개근 판정: ABSENCE가 해당 단위에 1분이라도 있으면 개근 아님.
  const absenceMap = att.maps.absence;

  const unitResults = units.map(u => {
    let hasAbsence = false;
    for (const dk of u.daySet) {
      if ((absenceMap.get(dk) || 0) > 0) { hasAbsence = true; break; }
    }
    // 산정제외(EXCLUDE)는 개근을 깨지 않는다고 가정(근로제공의무 자체가 정지/제외되는 기간)
    return { label: u.label, perfect: !hasAbsence };
  });

  const perfectCount = unitResults.filter(x=>x.perfect).length;

  // 상한: 상시 11, 비상시 9(요구사항 반영)
  const cap = isEmergency ? 9 : 11;
  return { units: unitResults, perfectCount, cappedDays: Math.min(perfectCount, cap), cap };
}

function computeGrant(person, att, calendar, settings) {
  const weeklyMin = person.weeklyMinutes || 0;
  const shortThreshold = RULES.attendance.short_time_threshold_weekly_minutes;
  if (weeklyMin < shortThreshold) {
    return {
      status: 'EXCLUDED',
      reason: `주 소정근로시간 ${Math.round(weeklyMin/60*10)/10}시간(= ${weeklyMin}분)으로 주 15시간 미만: 연차 산정 대상 제외`,
      baseDays: 0,
      addDays: 0,
      finalDays: 0,
      finalHours: 0,
    };
  }

  const grantDate = parseISODate(FIXED.grantDate);
  const baseDate = person.baseDate ? parseISODate(person.baseDate) : null;
  const years = person.seniorityYears ?? computeSeniorityYears(baseDate, grantDate);

  // 1년 미만(신규)
  if (years < 1) {
    const monthly = computeMonthlyPerfectUnits(person, calendar, att);
    return {
      status: 'NEW_UNDER_1Y',
      reason: '2026-03-01 기준 계속근로 1년 미만: 1개월 개근 시 1일(발생분 합산) 적용',
      baseDays: 0,
      addDays: 0,
      finalDays: monthly.cappedDays,
      finalHours: 0,
      monthly,
    };
  }

  // 기본/가산
  const baseDays = determineBaseDays(person, att);
  const addDays = person.addDays ?? computeAddDaysBySeniorityYears(years);
  const totalIfEligible = Math.min(25, baseDays + addDays);

  // 출근율 기준(상시=연간, 비상시=학기)
  const isEmergency = person.category.isEmergency;

  const rateRaw = isEmergency ? att.rateSemesterRaw : att.rateFullRaw;
  const rateRecalc = isEmergency ? att.rateSemesterRecalc : att.rateFullRecalc;

  // 정상 부여
  if (rateRaw >= 0.8) {
    return {
      status: 'NORMAL',
      reason: '전년도 출근율 80% 이상: 기본+가산 부여',
      baseDays,
      addDays,
      finalDays: totalIfEligible,
      finalHours: 0,
      totalIfEligible,
    };
  }

  // 비례 부여(제외기간 재산정 80% 충족)
  if (rateRaw < 0.8 && rateRecalc >= 0.8) {
    // 비례식: (기본+가산) × (연간총일수-제외기간) / 연간총일수
    // 연간총일수 = yearDays(토 제외) 기준
    const denomFull = att.denomFull;
    const excludedFull = att.excludedFull;
    const ratio = denomFull > 0 ? (denomFull - excludedFull) / denomFull : 0;
    const value = totalIfEligible * ratio;
    const yearTotalDays = calendar.yearDays.size;
    const excludedDaysApprox = Math.round(excludedFull / person.dailyMinutes);
    const rounded = roundProrationToDaysAndHours(value, person.dailyMinutes);

    return {
      status: 'PRORATED',
      reason: '전년도 출근율 80% 미만(전체)이나, 제외기간 제거 후 80% 이상: 비례 부여',
      baseDays,
      addDays,
      totalIfEligible,
      proration: {
        yearTotalDays,
        excludedDaysApprox,
        ratio,
        value,
        rounded,
      },
      finalDays: rounded.days,
      finalHours: rounded.hours,
    };
  }

  // 80% 미만(재산정도 미달): 개근월수
  const monthly = computeMonthlyPerfectUnits(person, calendar, att);
  return {
    status: 'MONTHLY',
    reason: '전년도 출근율 80% 미만(재산정 포함): 1개월 개근 시 1일(개근월수) 적용',
    baseDays,
    addDays,
    finalDays: monthly.cappedDays,
    finalHours: 0,
    monthly,
  };
}

/* -------------------------
 * 렌더링
 * ------------------------- */

function formatMinutesAsDayTime(totalMinutes, dailyMinutes) {
  if (totalMinutes == null) return '';
  const dm = dailyMinutes || 480;
  const sign = totalMinutes < 0 ? '-' : '';
  let m = Math.abs(Math.round(totalMinutes));

  const days = Math.floor(m / dm);
  m = m - days * dm;
  const hours = Math.floor(m / 60);
  const mins = m - hours * 60;

  const parts = [];
  if (days) parts.push(`${days}일`);
  if (hours) parts.push(`${hours}시간`);
  if (mins || parts.length === 0) parts.push(`${mins}분`);
  return sign + parts.join(' ');
}

function buildKindSummary(records, dailyMinutes, calendar) {
  const byKind = new Map(); // kind -> agg

  for (const rec of records) {
    const approval = safeText(rec.approval);
    const effectiveCategory = (approval === '완결') ? rec.cls.category : CATEGORY.REVIEW;

    const alloc = allocateRecordToDateMinutes(rec, dailyMinutes);
    const total = alloc.reduce((s, x) => s + x.minutes, 0);
    if (total <= 0) continue;

    const kind = rec.kind || '(종별없음)';
    if (!byKind.has(kind)) {
      byKind.set(kind, {
        kind,
        totalMinutes: 0,
        byCategory: { DEEMED: 0, EXCLUDE: 0, ABSENCE: 0, REVIEW: 0 },
        vacationCreditMinutes: 0,
        count: 0,
      });
    }
    const agg = byKind.get(kind);
    agg.totalMinutes += total;
    agg.count += 1;
    agg.byCategory[effectiveCategory] = (agg.byCategory[effectiveCategory] || 0) + total;

    if (effectiveCategory === CATEGORY.DEEMED && rec.cls.vacationCredit) {
      for (const { dateKey, minutes } of alloc) {
        if (calendar.vacationDays.has(dateKey)) agg.vacationCreditMinutes += minutes;
      }
    }
  }

  const list = Array.from(byKind.values());
  list.sort((a,b)=> b.totalMinutes - a.totalMinutes);
  return list;
}


function buildWarnings(kindSummary, dailyMinutes) {
  const warnings = [];
  const dm = dailyMinutes || 480;

  const getDays = (minutes) => minutes / dm;

  for (const k of (kindSummary || [])) {
    if (k.kind === '병가') {
      const days = getDays(k.totalMinutes || 0);
      if (days > 60.0001) {
        warnings.push(`병가 합계가 약 ${Math.round(days*10)/10}일로 60일 초과 가능: 유급(60일)·무급·휴직 구분 확인 필요`);
      }
    }
    if (k.kind === '자녀돌봄휴가' || k.kind === '가족돌봄휴가') {
      const days = getDays(k.totalMinutes || 0);
      if (days > 10.0001) {
        warnings.push(`${k.kind} 합계가 약 ${Math.round(days*10)/10}일로 10일 초과 가능: 출근간주 인정범위 확인 필요`);
      }
    }
  }

  return warnings;
}


function pct(x) {
  return `${Math.round(x * 1000) / 10}%`;
}

function renderSummary(calendar) {
  summaryBox.innerHTML = `
    <p>
      <b>기준기간</b>: ${calendar.fixedStart} ~ ${calendar.fixedEnd} (총 ${calendar.info.totalCalendarDays}일)
      / <b>토요일</b>: ${calendar.info.satCount}일
      / <b>토요일 제외</b>: 연간 ${calendar.info.yearDaysExclSat}일, 학기 ${calendar.info.semesterDaysExclSat}일, 방학 ${calendar.info.vacationDaysExclSat}일
    </p>
  `;
}

function renderResultsTable(rows) {
  const headers = [
    '성명','개인번호','직종구분','직종','근무형태','주소정(주)','1일(분)',
    '기준근속일','근속(년)','가산(일)',
    '출근율(비상시-학기)','출근율(상시-연간)','출근율(비상시-재산정)',
    '산정결과','부여(일)','부여(시간)','비고'
  ];
  const th = headers.map(h=>`<th>${h}</th>`).join('');
  const tr = rows.map(r => `
    <tr>
      <td>${r.name}</td>
      <td>${r.personalNo||''}</td>
      <td>${r.jobGroup||''}</td>
      <td>${r.job||''}</td>
      <td>${r.workForm||''}</td>
      <td>${Math.round((r.weeklyMinutes||0)/60*10)/10}</td>
      <td>${r.dailyMinutes||''}</td>
      <td>${r.baseDate||''}</td>
      <td>${r.seniorityYears ?? ''}</td>
      <td>${r.addDays ?? ''}</td>
      <td>${r.isEmergency ? pct(r.rateSemesterRaw) : '-'}</td>
      <td>${pct(r.rateFullRaw)}</td>
      <td>${r.isEmergency ? pct(r.rateSemesterRecalc) : pct(r.rateFullRecalc)}</td>
      <td>${r.grantStatus}</td>
      <td>${r.finalDays}</td>
      <td>${r.finalHours}</td>
      <td>${r.note||''}</td>
    </tr>
  `).join('');

  resultsBox.innerHTML = `
    <table>
      <thead><tr>${th}</tr></thead>
      <tbody>${tr}</tbody>
    </table>
  `;
}


function renderDetails(allDetails, openKeys = new Set()) {
  const CAT_LABEL = {
    DEEMED: '출근간주',
    EXCLUDE: '산정제외',
    ABSENCE: '결근성',
    REVIEW: '검토',
  };
  const catLabel = (c) => CAT_LABEL[c] || c || '';

  const html = (allDetails || []).map(d => {
    const dm = d.dailyMinutes || 480;

    const kindRows = (d.kindSummary || []).map(k => `
      <tr>
        <td>${k.kind}</td>
        <td>${formatMinutesAsDayTime(k.totalMinutes, dm)}</td>
        <td>${formatMinutesAsDayTime(k.byCategory.DEEMED || 0, dm)}</td>
        <td>${formatMinutesAsDayTime(k.byCategory.EXCLUDE || 0, dm)}</td>
        <td>${formatMinutesAsDayTime(k.byCategory.ABSENCE || 0, dm)}</td>
        <td>${formatMinutesAsDayTime(k.byCategory.REVIEW || 0, dm)}</td>
        <td>${formatMinutesAsDayTime(k.vacationCreditMinutes || 0, dm)}</td>
        <td>${k.count}</td>
      </tr>
    `).join('');

    const kindTable = kindRows
      ? `<table>
          <thead>
            <tr>
              <th>종별</th>
              <th>합계</th>
              <th>출근간주</th>
              <th>산정제외</th>
              <th>결근성</th>
              <th>검토</th>
              <th>방학근무크레딧</th>
              <th>건수</th>
            </tr>
          </thead>
          <tbody>${kindRows}</tbody>
        </table>`
      : `<p class="muted">취합할 복무 기록이 없습니다.</p>`;

    const warnings = (d.warnings || []).map(w => `<li>${w}</li>`).join('');
    const warningBox = warnings
      ? `<div class="warn"><b>경고</b><ul>${warnings}</ul></div>`
      : '';

    const reviewRows = (d.reviewRecords || []).map(r => `
      <tr>
        <td>${r.start}~${r.end}</td>
        <td>${r.kind}</td>
        <td>${r.duration.raw}</td>
        <td>${r.reason||''}</td>
        <td>${r.approval||''}</td>
        <td>${r.note||''}</td>
      </tr>
    `).join('');

    const reviewTable = reviewRows
      ? `<table>
          <thead><tr><th>기간</th><th>종별</th><th>일수/기간</th><th>사유</th><th>결재</th><th>비고</th></tr></thead>
          <tbody>${reviewRows}</tbody>
        </table>`
      : `<p class="muted">검토 필요 항목 없음</p>`;

    const allRecLimit = 300;
    const allRecs = d.allRecords || [];
    const showRecs = allRecs.slice(0, allRecLimit);
    const truncated = allRecs.length > allRecLimit;

    const ovCount = showRecs.filter(r => !!r.override).length;

    const allRows = showRecs.map(r => {
      const editable = (safeText(r.approval) === '완결');

      const autoCat = r.autoCls?.category || r.cls?.category || CATEGORY.REVIEW;
      const autoVac = Boolean(r.autoCls?.vacationCredit);
      const finalCat = r.cls?.category || CATEGORY.REVIEW;
      const finalVac = Boolean(r.cls?.vacationCredit);

      const isOverridden = Boolean(r.override);

      const selectDisabled = editable ? '' : 'disabled';
      const vacDisabled = editable ? '' : 'disabled';

      const badge = isOverridden ? '<span class="badge">수정</span>' : '';

      return `
        <tr>
          <td>${r.start}~${r.end}</td>
          <td>${r.kind}</td>
          <td>${r.duration.raw}</td>
          <td>${r.reason||''}</td>
          <td>${r.approval||''}</td>
          <td>${catLabel(autoCat)} / ${autoVac ? 'Y' : 'N'}</td>
          <td>
            <div class="row-inline">
              <select class="cls-select" data-rec-id="${r.id}" data-person-key="${d.key}" ${selectDisabled}>
                <option value="DEEMED" ${finalCat==='DEEMED' ? 'selected' : ''}>출근간주</option>
                <option value="EXCLUDE" ${finalCat==='EXCLUDE' ? 'selected' : ''}>산정제외</option>
                <option value="ABSENCE" ${finalCat==='ABSENCE' ? 'selected' : ''}>결근성</option>
                <option value="REVIEW" ${finalCat==='REVIEW' ? 'selected' : ''}>검토</option>
              </select>
              ${badge}
            </div>
          </td>
          <td>
            <label class="row-inline">
              <input type="checkbox" class="vac-credit" data-rec-id="${r.id}" data-person-key="${d.key}" ${finalVac ? 'checked' : ''} ${vacDisabled}>
              <span class="muted">방학크레딧</span>
            </label>
          </td>
        </tr>
      `;
    }).join('');

    const editTable = allRows
      ? `<table>
          <thead>
            <tr>
              <th>기간</th><th>종별</th><th>일수/기간</th><th>사유</th><th>결재</th>
              <th>자동(분류/크레딧)</th>
              <th>최종 분류</th>
              <th>최종 크레딧</th>
            </tr>
          </thead>
          <tbody>${allRows}</tbody>
        </table>`
      : `<p class="muted">기록 없음</p>`;

    const openAttr = openKeys && openKeys.has(d.key) ? 'open' : '';

    return `
      <details class="person-detail" data-person-key="${d.key}" ${openAttr}>
        <summary><b>${d.name}</b> (${d.personalNo||'개인번호 없음'}) - 총 ${d.recordCount}건 / 검토필요 ${d.reviewCount}건 / 수동수정(표시범위내) ${ovCount}건</summary>

        ${warningBox}

        <h4>종별 취합</h4>
        ${kindTable}

        <h4>검토 필요(자동 미반영)</h4>
        ${reviewTable}

        <details>
          <summary>전체 복무 목록(수동 보정 가능) (UI 표시 ${allRecLimit}건 제한${truncated ? ', 나머지는 엑셀 다운로드 확인' : ''})</summary>
          <p class="muted">
            ※ 분류/방학크레딧은 <b>결재상태=완결</b>인 건만 변경 가능합니다. 방학크레딧은 <b>출근간주</b>로 분류된 건이 방학 기간에 있을 때만 출근율(상시 기준) 계산에 반영됩니다.
          </p>
          ${editTable}
        </details>
      </details>
    `;
  }).join('');

  detailsBox.innerHTML = html || '<p class="muted">상세 데이터가 없습니다.</p>';
}


/* -------------------------
 * 수동보정(분류/방학근무크레딧) 적용
 * ------------------------- */

function getOverride(recId) {
  if (!STATE || !STATE.overrides) return null;
  return STATE.overrides[recId] || null;
}

function upsertOverride(recId, patch) {
  if (!STATE) return;
  STATE.overrides = STATE.overrides || {};
  const current = STATE.overrides[recId] || {};
  const next = { ...current, ...patch };

  // 자동값과 동일하면 override 제거(깔끔 유지)
  const rec = (STATE.allRecords || []).find(r => r.id === recId);
  if (rec && rec.autoCls) {
    const sameCat = (next.category ?? rec.autoCls.category) === rec.autoCls.category;
    const sameVac = (next.vacationCredit ?? rec.autoCls.vacationCredit) === rec.autoCls.vacationCredit;
    if (sameCat && sameVac) {
      delete STATE.overrides[recId];
      return;
    }
  }
  STATE.overrides[recId] = next;
}

function applyOverridesToRecords(allRecords, overrides) {
  const ov = overrides || {};
  for (const rec of allRecords) {
    const auto = rec.autoCls || rec.cls || { category: CATEGORY.REVIEW, vacationCredit: false, tags: ['missing_auto'] };
    rec.autoCls = auto;

    const o = ov[rec.id];
    if (!o) {
      rec.cls = { ...auto };
      rec.override = null;
      continue;
    }
    const merged = { ...auto };
    if (o.category) merged.category = o.category;
    if (typeof o.vacationCredit === 'boolean') merged.vacationCredit = o.vacationCredit;
    merged.tags = Array.from(new Set([...(auto.tags || []), 'manual_override']));
    rec.cls = merged;
    rec.override = o;
  }
}

function computeAllFromState() {
  if (!STATE) throw new Error('STATE가 없습니다. 먼저 계산 실행을 수행하세요.');

  const { persons, allRecords, calendar, settings } = STATE;

  applyOverridesToRecords(allRecords, STATE.overrides);

  // Group by key
  const recMap = new Map();
  for (const rec of allRecords) {
    const k = rec.key;
    if (!recMap.has(k)) recMap.set(k, []);
    recMap.get(k).push(rec);
  }

  const resultRows = [];
  const detailRows = [];

  for (const p of persons) {
    const recs = recMap.get(p.key) || [];
    const att = computeAttendanceForPerson(p, recs, calendar, settings);
    const grant = computeGrant(p, att, calendar, settings);

    resultRows.push({
      key: p.key,
      name: p.name,
      personalNo: p.personalNo,
      job: p.job,
      jobGroup: p.jobGroup,
      workForm: p.workForm,
      workType: p.workType,
      weeklyMinutes: p.weeklyMinutes,
      dailyMinutes: p.dailyMinutes,
      baseDate: p.baseDate,
      seniorityYears: p.seniorityYears,
      addDays: p.addDays,
      isEmergency: p.category.isEmergency,
      rateSemesterRaw: att.rateSemesterRaw,
      rateSemesterRecalc: att.rateSemesterRecalc,
      rateFullRaw: att.rateFullRaw,
      rateFullRecalc: att.rateFullRecalc,
      grantStatus: grant.status,
      finalDays: grant.finalDays,
      finalHours: grant.finalHours,
      note: grant.reason,
    });

    const kindSummary = buildKindSummary(recs, p.dailyMinutes, calendar);
    const warnings = buildWarnings(kindSummary, p.dailyMinutes);

    detailRows.push({
      key: p.key,
      name: p.name,
      personalNo: p.personalNo,
      dailyMinutes: p.dailyMinutes,
      recordCount: recs.length,
      kindSummary,
      reviewCount: (att.reviewRecords || []).length,
      reviewRecords: att.reviewRecords || [],
      allRecords: recs,
      warnings,
    });
  }

  return { calendar, rows: resultRows, allRecords, detailRows, overrides: STATE.overrides || {} };
}

function recomputeAndRender(focusPersonKey) {
  // 열린 사람 상세(접힘) 상태 보존
  const openKeys = new Set(
    Array.from(document.querySelectorAll('details.person-detail[open]')).map(d => d.dataset.personKey)
  );
  if (focusPersonKey) openKeys.add(focusPersonKey);

  const result = computeAllFromState();
  renderResultsTable(result.rows);
  renderDetails(result.detailRows, openKeys);

  LAST_RESULT = result;
  el('btnDownload').disabled = false;
}



/* -------------------------
 * 다운로드(엑셀)
 * ------------------------- */
function downloadWorkbook(result) {
  if (!window.XLSX) throw new Error('XLSX 라이브러리 없음');

  const wb = XLSX.utils.book_new();

  // Sheet 1: 연차 결과
  const sheet1 = result.rows.map(r => ({
    '성명': r.name,
    '개인번호': r.personalNo,
    '직종구분': r.jobGroup,
    '직종': r.job,
    '근무형태': r.workForm,
    '주 소정근로시간(분)': r.weeklyMinutes,
    '1일 소정근로시간(분)': r.dailyMinutes,
    '기준근속일': r.baseDate,
    '근속연수(년)': r.seniorityYears,
    '가산일수': r.addDays,
    '비상시-학기 출근율': r.isEmergency ? r.rateSemesterRaw : null,
    '상시-연간 출근율': r.rateFullRaw,
    '비상시-재산정 출근율': r.isEmergency ? r.rateSemesterRecalc : r.rateFullRecalc,
    '부여방식': r.grantStatus,
    '부여일수': r.finalDays,
    '부여시간': r.finalHours,
    '비고': r.note || '',
  }));
  const ws1 = XLSX.utils.json_to_sheet(sheet1);
  XLSX.utils.book_append_sheet(wb, ws1, '연차부여결과');

  // Sheet 2: 복무 레코드 전체(분류 포함) - 개인별 합산은 추후 확장 가능
  const sheet2 = result.allRecords.map(rec => ({
    '개인번호': rec.personalNo || '',
    '성명': rec.name,
    '기간': `${rec.start}~${rec.end}`,
    '일수/기간': rec.duration.raw,
    '종별': rec.kind,
    '사유': rec.reason || '',
    '결재상태': rec.approval || '',
    '자동분류': rec.autoCls?.category || '',
    '자동방학근무크레딧': rec.autoCls?.vacationCredit ? 'Y' : 'N',
    '최종분류': rec.cls.category,
    '최종방학근무크레딧': rec.cls.vacationCredit ? 'Y' : 'N',
    '수동수정여부': rec.override ? 'Y' : 'N',
    '태그': (rec.cls.tags || []).join(','),
  }));
  const ws2 = XLSX.utils.json_to_sheet(sheet2);
  XLSX.utils.book_append_sheet(wb, ws2, '복무취합(전체)');

  XLSX.writeFile(wb, '2026학년도_연차부여결과.xlsx');
}

/* -------------------------
 * 실행
 * ------------------------- */
async function run() {
  resetUI();

  try {
    if (!RULES) await loadRules();
    logStatus('규칙 로드 완료');

    const hrFile = el('hrFile').files?.[0];
    const workFiles = Array.from(el('workFiles').files || []);

    if (!hrFile) throw new Error('인사기록 조회 엑셀 파일을 업로드하세요.');
    if (workFiles.length === 0) throw new Error('근무상황목록 파일을 1개 이상 업로드하세요.');

    const calendar = buildCalendar();
    renderSummary(calendar);

    // HR
    logStatus(`인사기록 파일 읽는 중: ${hrFile.name}`);
    const hrWb = await readWorkbook(hrFile);
    const hrSheetName = hrWb.SheetNames.find(n => RULES.parsing.hr_sheet_name_candidates.includes(n)) || hrWb.SheetNames[0];
    const hrRowsRaw = sheetToObjects(hrWb.Sheets[hrSheetName]);
    logStatus(`인사기록 시트: ${hrSheetName}, 행수: ${hrRowsRaw.length}`);

    const persons = parseHrPersons(hrRowsRaw);
    logStatus(`인사기록 대상자: ${persons.length}명`);

    // Work records
    const allRecords = [];
    for (const f of workFiles) {
      logStatus(`근무상황 파일 읽는 중: ${f.name}`);
      const wb = await readWorkbook(f);
      const sname = wb.SheetNames.find(n => RULES.parsing.work_status_sheet_name_candidates.includes(n)) || wb.SheetNames[0];
      const rows = sheetToObjects(wb.Sheets[sname]);
      const recs = parseWorkRecords(rows);
      allRecords.push(...recs);
      logStatus(`- 시트: ${sname}, 레코드: ${recs.length}건`);
    }
    logStatus(`근무상황 총 레코드: ${allRecords.length}건`);

    const settings = {
      parentalDeemedCapDays: Number(el('parentalDeemedCapDays').value || 365),
    };

    // 전역 STATE 저장(수동보정/재계산 용)
    STATE = {
      persons,
      allRecords,
      calendar,
      settings,
      overrides: {}, // recId -> {category, vacationCredit}
    };

    // 최초 계산/렌더
    recomputeAndRender();

    logStatus('완료');
  } catch (e) {
    console.error(e);
    logStatus(`오류: ${e.message || e}`);
  }
}

el('btnRun').addEventListener('click', run);

// 수동보정 이벤트(복무 분류/방학크레딧 변경)
detailsBox.addEventListener('change', (e) => {
  const t = e.target;
  if (!STATE || !t) return;

  const recId = t.dataset?.recId;
  const personKey = t.dataset?.personKey || null;
  if (!recId) return;

  if (t.classList.contains('cls-select')) {
    upsertOverride(recId, { category: t.value });
    recomputeAndRender(personKey);
  } else if (t.classList.contains('vac-credit')) {
    upsertOverride(recId, { vacationCredit: Boolean(t.checked) });
    recomputeAndRender(personKey);
  }
});
el('btnDownload').addEventListener('click', () => {
  if (!LAST_RESULT) return;
  try {
    downloadWorkbook(LAST_RESULT);
  } catch (e) {
    logStatus(`다운로드 오류: ${e.message || e}`);
  }
});

// 초기 로드
loadRules().catch(e => logStatus(`rules_2026.json 로드 오류: ${e.message || e}`));
