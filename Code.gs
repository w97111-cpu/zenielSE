const LEDGER_TAB = '연차촉진_데이터';
const LEDGER_ID = 'LEDGER_SPREADSHEET_ID';
const LIVE = 'LEAVE_DELIVERY_ENABLED';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('제니엘 연차 사용촉진');
}

function initializeLedger() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  if (!workbook) throw Error('Google Sheets에서 연결된 Apps Script 프로젝트를 여세요.');
  PropertiesService.getScriptProperties().setProperty(LEDGER_ID,workbook.getId());
  if (!workbook.getSheetByName(LEDGER_TAB)) workbook.insertSheet(LEDGER_TAB);
  return readLedger().staff.length;
}

function ledgerSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(LEDGER_ID);
  if (!id) throw Error('Apps Script에서 initializeLedger를 먼저 실행해 주세요.');
  return SpreadsheetApp.openById(id).getSheetByName(LEDGER_TAB);
}

function readLedger() {
  const sheet = ledgerSheet_();
  if (!sheet.getLastRow()) return {staff:[],periods:[],log:[],rev:0};
  const chunks = sheet.getRange(1,1,sheet.getLastRow(),1).getValues();
  return JSON.parse(chunks.map(([v])=>String(v).slice(1)).join(''));
}

function lockLedger_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {return callback();}
  finally {lock.releaseLock();}
}

function writeLedger_(data) {
  data.rev = (data.rev||0)+1;
  const serialized = JSON.stringify(data);
  const chunks = [];
  for (let i=0;i<serialized.length;i+=20000) chunks.push(['x'+serialized.slice(i,i+20000)]);
  const sheet = ledgerSheet_();
  while(chunks.length<sheet.getLastRow()) chunks.push(['']);
  if(chunks.length>sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(),chunks.length-sheet.getMaxRows());
  sheet.getRange(1,1,chunks.length,1).setValues(chunks);
  SpreadsheetApp.flush();
}

function saveLedger(data, rev, importing) {
  return lockLedger_(()=>{
    const current = readLedger();
    if(current.rev!==rev) throw Error('다른 작업에서 변경된 내용이 있습니다. 새로고침 후 다시 저장해 주세요.');
    if (importing && current.log.length) throw Error('발송 이력이 있는 운영 데이터는 전체 가져오기로 바꿀 수 없습니다.');
    if (!data || !Array.isArray(data.staff) || !Array.isArray(data.periods)) throw Error('데이터 형식을 확인해 주세요.');
    const staff = new Set(data.staff.map(e=>e.id));
    if(staff.size!==data.staff.length || new Set(data.periods.map(p=>p.id)).size!==data.periods.length ||
        data.periods.some(p=>!staff.has(p.staff))) throw Error('중복 항목 또는 직원 연결을 확인해 주세요.');
    const next = {staff:data.staff,periods:data.periods,log:current.log,rev};
    writeLedger_(next);
    return next;
  });
}

function installDailyTrigger() {
  ledgerSheet_();
  if(!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='dailyLeaveNotices')) {
    ScriptApp.newTrigger('dailyLeaveNotices').timeBased().everyDays(1).atHour(9).inTimezone('Asia/Seoul').create();
  }
}

function previewDailyNotices() {
  const data = readLedger();
  const byId = new Map(data.staff.map(e=>[e.id,e]));
  const day = koreaToday();
  const summary = {};
  for(const p of data.periods){
    const {state} = decision(byId.get(p.staff),p,day);
    summary[state]=(summary[state]||0)+1;
  }
  console.log(JSON.stringify(summary));
  return summary;
}

function dailyLeaveNotices() {
  if (PropertiesService.getScriptProperties().getProperty(LIVE)!=='YES') return previewDailyNotices();
  return lockLedger_(()=>{
    const day = koreaToday();
    const data = readLedger();
    const byId = new Map(data.staff.map(e=>[e.id,e]));
    const processed = new Set(data.log.map(x=>x.key));
    let remaining = MailApp.getRemainingDailyQuota();
    const candidates = data.periods.map(p=>({p,e:byId.get(p.staff),due:decision(byId.get(p.staff),p,day)}))
      .filter(x=>x.due.stage).sort((a,b)=>(a.due.stage===1?a.due.last:a.due.second).localeCompare(b.due.stage===1?b.due.last:b.due.second));
    const began = Date.now();
    for(const {p,e,due} of candidates){
      if(!remaining || Date.now()-began>240000) break;
      const stage = due.stage;
      const key = p.id+'/'+stage;
      if(processed.has(key))continue;
      const message=notice(e,p,day,stage);
      const log={at:new Date().toISOString(),staff:e.id,key,stage,recipient:e.email,result:'발송 여부 확인 필요',subject:message.subject,body:message.body};
      data.log.push(log);
      writeLedger_(data);
      try {
        MailApp.sendEmail({to:e.email,subject:message.subject,body:message.body,name:'제니엘'});
        p['sent'+stage]=day;
        log.result='발송 완료';
      } catch(error) {
        log.result='실패 또는 발송 여부 확인 필요';
      }
      writeLedger_(data);
      remaining--;
    }
  });
}

function dateOK(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value||'')) return false;
  const parsed = new Date(value+'T00:00:00Z');
  return Number.isFinite(+parsed) && parsed.toISOString().slice(0,10) === value;
}
function daysAfter(value, count) {
  const valueDate = new Date(value+'T00:00:00Z');
  valueDate.setUTCDate(valueDate.getUTCDate()+count);
  return valueDate.toISOString().slice(0,10);
}
function monthsAfter(value, count) {
  const valueDate = new Date(value+'T00:00:00Z');
  const day = valueDate.getUTCDate();
  valueDate.setUTCDate(1);
  valueDate.setUTCMonth(valueDate.getUTCMonth()+count);
  const last = new Date(Date.UTC(valueDate.getUTCFullYear(),valueDate.getUTCMonth()+1,0)).getUTCDate();
  valueDate.setUTCDate(Math.min(day,last));
  return valueDate.toISOString().slice(0,10);
}
function koreaToday() {
  return new Date(Date.now()+9*3600000).toISOString().slice(0,10);
}
function deadline(p) {
  if (!dateOK(p.end) || !['일반','최초 1년','최초 1년 추가분'].includes(p.kind)) return {};
  const next = daysAfter(p.end,1);
  const extra = p.kind === '최초 1년 추가분';
  const annual = p.kind === '일반';
  const first = monthsAfter(next, annual ? -6 : extra ? -1 : -3);
  return {first, last:daysAfter(first,extra?4:9),
    second:extra?daysAfter(p.end,-10):daysAfter(monthsAfter(next,annual?-2:-1),-1),
    reply:dateOK(p.received)?daysAfter(p.received,10):''};
}
function decision(e,p,t) {
  const d = deadline(p);
  const result = (state,stage=0) => ({...d,state,stage});
  if (!e || e.status!=='재직' || (e.quit && e.quit<=t)) return result('발송 제외');
  if (p.balance !== null && p.balance<=0) return result('잔여 없음');
  if (!d.first || !dateOK(p.start) || p.start>p.end || !Number.isFinite(p.balance) ||
      !['입사일','회계연도'].includes(e.basis) || !p.ready) return result('대장 검토 필요');
  if (t>p.end) return result('사용기간 종료');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e.email||'')) return result('이메일 확인');
  if (p.sent2) return result('2차 발송 기록');
  if (!p.sent1) {
    if (p.old1 || p.old2) return result('기존 촉진 이력 확인');
    if (t<d.first) return result('1차 예정');
    if (t>d.last) return result('1차 기한 경과');
    return result('1차 발송 대상',1);
  }
  if (!dateOK(p.sent1) || p.sent1<d.first || p.sent1>d.last) return result('1차 통보일 확인');
  if (!d.reply || p.received<p.sent1 || p.received>t) return result('1차 수신일 확인');
  if (p.replied>=p.balance) return result('사용계획 접수 완료');
  if (t<=d.reply) return result('1차 회신 대기');
  if (t>d.second) return result('2차 기한 경과');
  const dates = (p.designated||'').trim().split(/[\s,]+/);
  if (new Set(dates).size!==dates.length ||
      dates.length!==Math.ceil(p.balance-(p.replied||0)) ||
      dates.some(value=>!dateOK(value)||value<=t||value<p.start||value>p.end)) return result('2차 사용일 지정 필요');
  return result('2차 발송 대상',2);
}
function notice(e,p,t,stage) {
  const heading = '연차유급휴가 사용촉진 '+stage+'차';
  const left = p.balance-(p.replied||0);
  const body = stage===1
    ? '미사용 연차 '+p.balance+'일의 사용 시기를 정해, 이 안내를 받은 날부터 10일 이내에 이 메일에 회신해 주세요.'
    : '사용계획이 제출되지 않은 잔여 연차 '+left+'일에 대해 회사가 다음과 같이 사용일을 지정하여 통보합니다.\n지정 사용일 '+p.designated+(left%1 ? '\n마지막 지정일 사용량 '+(left%1)+'일' : '');
  return {subject:'[제니엘] '+heading,body:e.name+' 님께\n\n'+heading+'\n발송일 '+t+'\n사용 가능 기간 '+p.start+' ~ '+p.end+'\n\n'+body+'\n\n제니엘'};
}
