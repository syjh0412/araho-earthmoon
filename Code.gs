/**
 * 아라호 관제센터 — Google Apps Script 백엔드 (구글 시트 + Gemini 피드백)
 * ------------------------------------------------------------------
 * 설치 (5분):
 *  1) 새 구글 시트 만들기 → 확장 프로그램 → Apps Script
 *  2) 기존 내용을 지우고 이 파일 전체를 붙여넣기 → 저장
 *  3) 배포 → 새 배포 → 유형: 웹 앱
 *       - 다음 사용자 인증 정보로 실행: 나
 *       - 액세스 권한이 있는 사용자: 모든 사용자
 *     → 배포 → 권한 허용 → "웹 앱 URL" 복사 (…/exec 로 끝남)
 *  4) index.html 맨 위 API = '여기에 붙여넣기'
 *  5) AI 키와 센터장 코드 — 둘 중 편한 곳에 넣으면 됩니다 (둘 다 있으면 스크립트 속성이 우선)
 *     (가) 스크립트 속성: 왼쪽 ⚙ 프로젝트 설정 → 맨 아래 '스크립트 속성' → 속성 추가
 *          GEMINI_KEY = 발급받은 키 / PIN = 센터장 코드 (예: 0726)
 *     (나) 시트 '설정' 탭 (웹 앱을 한 번 열면 자동 생성): GEMINI_KEY, PIN 의 B칸
 *  6) (선택) 학생 명단: 아무 탭에나 '반 | 번호 | 이름' 머리글로 명단을 붙여넣으면
 *     학생은 반·번호를 고르고 자기 이름을 확인한 뒤 '콜사인'(비밀번호)으로 입장합니다.
 *     콜사인은 첫 입장 때 학생이 정하고 명단 탭의 '콜사인' 열에 저장됩니다.
 * 이후 학생 기록은 '학생'(현재 상태) / '기록'(전체) 탭과
 *   한줄논술①~④ · 시험 문항1~3 · 수배파일 탭(문항별)에 자동으로 쌓이고,
 * 앱의 센터장 탭에서 실시간으로 볼 수 있습니다.
 *
 * ── 시즌 2 (Ⅶ-2 지구와 달, araho-earthmoon) ──────────────────────
 *  이 파일 하나가 시즌 1·2 두 앱을 함께 처리합니다. 시즌 2 앱은 요청마다 s=2를 보내고,
 *  그러면 '학생 S2' · '기록 S2' · '교신 S2' · '응답 S2' · '쿠폰 S2' 탭에 따로 저장됩니다.
 *  명단(콜사인)과 '설정' 탭(GEMINI_KEY·PIN)은 두 시즌이 같이 씁니다.
 *  s를 보내지 않는 시즌 1 앱은 예전과 똑같이 동작합니다.
 *  업데이트 방법: 이 파일로 Apps Script를 바꾼 뒤 배포 → 배포 관리 → ✏️ 수정 → 버전 '새 버전' → 배포
 *  (이렇게 하면 웹 앱 URL이 바뀌지 않습니다)
 */

var SHEET_STUDENTS = '학생';
var SHEET_LOG = '기록';
var SHEET_CONF = '설정';
var S2 = false;              // 요청마다 handle()에서 정함
var HOF_AT = 290;            // 명예의 전당 기준 (시즌 2는 250)
function seasonSetup(p){
  S2 = String(p.s||'')==='2';
  var x = S2 ? ' S2' : '';
  SHEET_STUDENTS = '학생'+x; SHEET_LOG = '기록'+x; SHEET_CP = '쿠폰'+x; SHEET_Q = '교신'+x; SHEET_A = '응답'+x;
  ITEM_SHEETS = S2 ? ITEM_SHEETS_S2 : ITEM_SHEETS_S1;
  HOF_AT = S2 ? 250 : 290;
}

function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet(name, header){
  var s = ss().getSheetByName(name);
  if(!s){ s = ss().insertSheet(name); if(header) s.appendRow(header); s.setFrozenRows(1); }
  return s;
}

function conf(key){
  // 1순위: 스크립트 속성(⚙ 프로젝트 설정 → 스크립트 속성). 2순위: 시트 '설정' 탭.
  try{ var sp = PropertiesService.getScriptProperties().getProperty(key); if(sp && String(sp).trim()) return String(sp).trim(); }catch(e){}
  var s = sheet(SHEET_CONF, ['항목','값','설명']);
  var v = s.getDataRange().getValues();
  if(v.length < 2){
    s.appendRow(['GEMINI_KEY','', 'Google AI Studio(aistudio.google.com)에서 발급한 키를 B열에 붙여넣기']);
    s.appendRow(['PIN','0726','센터장 탭 입장 코드']);
    s.appendRow(['MODEL','gemini-3.6-flash','피드백에 쓸 모델 (2026년 기준. 안 되면 AI Studio에서 지금 쓸 수 있는 flash 모델 이름으로)']);
    v = s.getDataRange().getValues();
  }
  for(var i=1;i<v.length;i++){ if(String(v[i][0]).trim()===key) return String(v[i][1]).trim(); }
  return '';
}

function out(obj, cb){
  var txt = JSON.stringify(obj);
  if(cb){ return ContentService.createTextOutput(cb+'('+txt+')').setMimeType(ContentService.MimeType.JAVASCRIPT); }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

/* ───────── 라우팅 ───────── */
function doGet(e){
  var p = (e && e.parameter) || {};
  try{ return out(handle(p), p.callback); }
  catch(err){ return out({ok:false, error:String(err && err.message || err)}, p.callback); }
}
function doPost(e){
  var p = {};
  try{ p = JSON.parse(e.postData.contents || '{}'); }catch(err){ p = (e && e.parameter) || {}; }
  try{ return out(handle(p)); }
  catch(err){ return out({ok:false, error:String(err && err.message || err)}); }
}

function handle(p){
  var a = p.action;
  seasonSetup(p);
  conf('PIN'); // 설정 탭 보장
  if(a==='ping')      return {ok:true, t:new Date().toISOString(), ai: !!conf('GEMINI_KEY')};
  if(a==='names')     return rosterPublic(p);
  if(a==='login')     return login(p);
  if(a==='resetcode') return resetCode(p);
  if(a==='save')      return save(p);
  if(a==='log')       return log(p);
  if(a==='feedback')  return feedback(p);
  if(a==='couponClaim') return couponClaim(p);
  if(a==='myCoupons')   return myCoupons(p);
  if(a==='couponList')  return couponList(p);
  if(a==='couponGive')  return couponGive(p);
  if(a==='qadd')      return qadd(p);
  if(a==='qlist')     return qlist(p);
  if(a==='qanswer')   return qanswer(p);
  if(a==='adel')      return adel(p);
  if(a==='moonUp')    return moonUp(p);
  if(a==='moonList')  return moonList(p);
  if(a==='moonVote')  return moonVote(p);
  if(a==='moonDel')   return moonDel(p);
  if(a==='tlogin')    return tlogin(p);
  if(a==='roster')    return roster(p);
  if(a==='logs')      return logs(p);
  if(a==='legends')   return legends(p);
  throw new Error('알 수 없는 요청: '+a);
}

/* ───────── 학생 ───────── */
function key(p){ return [String(p.cls||'').trim(), String(p.num||'').trim(), String(p.name||'').trim()].join('|'); }

function findRow(s, k){
  var v = s.getDataRange().getValues();
  for(var i=1;i<v.length;i++){ if(String(v[i][0])===k) return i+1; }
  return -1;
}

/* ───────── 명단(학생 명부) ─────────
   시트 아무 탭이든 첫 3줄 안에 '반' '번호' '이름' 머리글이 있으면 명단으로 인식합니다.
   '콜사인' 열이 없으면 자동으로 만들어 학생이 처음 입장할 때 정한 콜사인(비밀번호)을 저장합니다. */
/* 앱이 스스로 만드는 탭(학생·기록·교신·응답·쿠폰·설정, 문항별 탭)은 명단으로 착각하지 않는다 */
function isAppSheet(name){
  name = String(name);
  if(/^(학생|기록|교신|응답|쿠폰|설정)( S2)?$/.test(name)) return true;
  if(name==='달앨범') return true;
  if(/^(한줄논술|시험 문항|수배파일|S2 )/.test(name)) return true;
  return false;
}
function findRoster(){
  var sheets = ss().getSheets();
  for(var i=0;i<sheets.length;i++){
    var sh = sheets[i]; if(sh.getLastRow()<1 || isAppSheet(sh.getName())) continue;
    var top = sh.getRange(1,1,Math.min(3,sh.getLastRow()),Math.max(1,sh.getLastColumn())).getValues();
    for(var r=0;r<top.length;r++){
      var row = top[r].map(function(x){return String(x).replace(/\s/g,'');});
      var ci=row.indexOf('반'), ni=row.indexOf('번호'), mi=row.indexOf('이름');
      if(ci>=0 && ni>=0 && mi>=0){
        if(sh.getLastRow()<=r+1) continue;   // 머리글만 있고 학생이 없는 탭은 명단이 아님
        var ki = row.indexOf('콜사인'); if(ki<0) ki = row.indexOf('비밀번호');
        if(ki<0){ ki = row.length; sh.getRange(r+1, ki+1).setValue('콜사인'); }
        return {sh:sh, head:r+1, ci:ci, ni:ni, mi:mi, ki:ki};
      }
    }
  }
  return null;
}
/* 공개 응답에서는 학생 이름 가운데 글자를 가린다 (김민준 → 김○준, 이준 → 이○). 센터장 코드가 맞을 때만 전체 이름. */
function maskName(n){ n=String(n||'').trim(); if(!n || n==='익명 승무원' || n==='센터장') return n;
  var c=n.split(''); if(c.length<=1) return n; if(c.length===2) return c[0]+'○';
  return c[0]+new Array(c.length-1).join('○')+c[c.length-1]; }
function pinOk(p){ var pin=conf('PIN'); return !!(pin && String(p.pin||'')===pin); }
function rosterPublic(p){
  var R = findRoster(); if(!R) return {ok:true, roster:null};
  var v = R.sh.getRange(R.head+1,1,Math.max(0,R.sh.getLastRow()-R.head),Math.max(R.ki+1,R.sh.getLastColumn())).getValues();
  var out = {};
  for(var i=0;i<v.length;i++){ var c=String(v[i][R.ci]).trim(), n=String(v[i][R.ni]).trim(), m=String(v[i][R.mi]).trim();
    if(!c||!n||!m) continue; (out[c]=out[c]||[]).push({num:n,name:maskName(m),set:!!String(v[i][R.ki]||'').trim()}); }
  return {ok:true, roster:out};
}
function rosterRow(R, cls, num){
  var v = R.sh.getRange(R.head+1,1,Math.max(0,R.sh.getLastRow()-R.head),Math.max(R.ki+1,R.sh.getLastColumn())).getValues();
  for(var i=0;i<v.length;i++){ if(String(v[i][R.ci]).trim()===String(cls).trim() && String(v[i][R.ni]).trim()===String(num).trim()) return {row:R.head+1+i, name:String(v[i][R.mi]).trim(), code:String(v[i][R.ki]||'').trim()}; }
  return null;
}
function login(p){
  if(!p.cls || !p.num) throw new Error('반과 번호를 입력하세요');
  var R = findRoster(), name = String(p.name||'').trim();
  if(R){
    var row = rosterRow(R, p.cls, p.num);
    if(!row) throw new Error('명단에 없는 반·번호예요. 센터장에게 확인하세요');
    name = row.name;
    var code = String(p.code||'').trim();
    if(!row.code){
      if(code.length<2) throw new Error('처음 입장이에요. 콜사인(2~8자)을 정해 주세요');
      R.sh.getRange(row.row, R.ki+1).setValue(code);
    }else if(code!==row.code){ throw new Error('콜사인이 달라요'); }
  }else if(!name){ throw new Error('이름을 입력하세요'); }
  var s = sheet(SHEET_STUDENTS, ['key','반','번호','이름','점수','등급','진도','상태JSON','최근접속']);
  var k = key({cls:p.cls,num:p.num,name:name}), r = findRow(s,k), state = null;
  if(r<0){ s.appendRow([k, p.cls, p.num, name, 0, S2 ? '지상 관측병' : '지상 훈련생', '', '', new Date()]); }
  else{ var raw = s.getRange(r,8).getValue(); try{ state = raw ? JSON.parse(raw) : null; }catch(e){ state=null; } s.getRange(r,9).setValue(new Date()); }
  var res = {ok:true, name:name, state:state, adj:adjFor(p.cls,p.num)};
  if(S2) res.prev = prevSeason(k);
  return res;
}
/* 시즌 2 입장 때: 시즌 1('학생' 탭) 점수·등급·승무원 모습을 함께 돌려준다 (없으면 null) */
function prevSeason(k){
  try{
    var s1 = ss().getSheetByName('학생'); if(!s1) return null;
    var v = s1.getDataRange().getValues();
    for(var i=1;i<v.length;i++){ if(String(v[i][0])===k){
      var sty=null; try{ var st=v[i][7]?JSON.parse(v[i][7]):null; sty=st&&st.sty||null; }catch(e){}
      return {pts:Number(v[i][4]||0), rank:String(v[i][5]||''), sty:sty}; } }
  }catch(e){}
  return null;
}
/* 센터장 점수 조정 — '점수조정'(시즌 2는 '점수조정 S2') 탭에 줄을 넣으면 학생 앱이 다음 저장·입장 때 한 번만 반영합니다.
   점수: 더하거나 뺄 점수(예: -10) · 응답수: 응답 점수 상한(50건)에서 되돌려 줄 건수(예: -5) · id는 비워 두면 자동으로 채워져요 */
function adjFor(cls, num){
  try{
    var s = sheet('점수조정'+(S2?' S2':''), ['id','반','번호','이름','점수','응답수','까닭']);
    var v = s.getDataRange().getValues(), n = function(x){ return String(x).replace(/\.0$/,'').trim(); }, out = [];
    for(var i=1;i<v.length;i++){ var r=v[i];
      if(n(r[1])!==n(cls) || n(r[2])!==n(num)) continue;
      if(!Number(r[4]) && !Number(r[5])) continue;
      var id = String(r[0]||''); if(!id){ id = qid(); s.getRange(i+1,1).setValue(id); }
      out.push({id:id, pts:Number(r[4]||0), ans:Number(r[5]||0), why:String(r[6]||'')}); }
    return out;
  }catch(e){ return []; }
}
function resetCode(p){
  checkPin(p); var R = findRoster(); if(!R) throw new Error('명단 탭이 없어요');
  var row = rosterRow(R, p.cls, p.num); if(!row) throw new Error('명단에 없어요');
  R.sh.getRange(row.row, R.ki+1).setValue(''); return {ok:true};
}

function save(p){
  var s = sheet(SHEET_STUDENTS, ['key','반','번호','이름','점수','등급','진도','상태JSON','최근접속']);
  var k = key(p), r = findRow(s,k);
  if(r<0){ s.appendRow([k, p.cls, p.num, p.name, 0, '', '', '', new Date()]); r = s.getLastRow(); }
  var st = String(p.state||'');
  if(st.length > 45000) st = st.slice(0,45000);
  s.getRange(r,5,1,5).setValues([[Number(p.pts||0), String(p.rank||''), String(p.prog||''), st, new Date()]]);
  return {ok:true, adj:adjFor(p.cls,p.num)};
}

/* 기록은 전체 '기록' 탭 + 문항별 탭에 나눠 쌓입니다 */
var ITEM_SHEETS_S1 = {w_c1:'한줄논술① 태양계 구성원', w_c2:'한줄논술② 행성 분류', w_c3:'한줄논술③ 태양 활동', w_c4:'한줄논술④ 지구 영향'};
var ITEM_SHEETS_S2 = {r_c1:'S2 관측보고① 일주 운동', r_c2:'S2 관측보고② 별자리 변화', r_c3:'S2 관측보고③ 달의 위상', r_c4:'S2 관측보고④ 일식과 월식'};
var ITEM_SHEETS = ITEM_SHEETS_S1;
function itemSheetName(p){
  var it = String(p.item||''), ty = String(p.type||'');
  for(var k in ITEM_SHEETS){ if(it.indexOf(k)>=0) return ITEM_SHEETS[k]; }
  if(S2){ if(ty==='관측추리') return 'S2 관측추리'; return ''; }
  var m = it.match(/문항\s*([123])/); if(m) return '시험 문항'+m[1];
  if(ty==='수배파일') return '수배파일';
  return '';
}
function log(p){
  var row = [new Date(), p.cls||'', p.num||'', p.name||'', p.type||'', p.item||'', String(p.text||'').slice(0,5000), p.score||'', p.ai||'', String(p.aifb||'').slice(0,3000)];
  var head = ['시각','반','번호','이름','종류','항목','내용','자동점검','AI수준','AI피드백'];
  sheet(SHEET_LOG, head).appendRow(row);
  var name = itemSheetName(p);
  if(name){ var s2 = sheet(name, head); s2.appendRow(row); }
  return {ok:true};
}

/* ───────── 센터장 보급 쿠폰 (마이쭈) ─────────
   학생이 조건을 채우면 '보급 신청' → '쿠폰' 탭에 W(대기)로 기록 → 센터장 콘솔에서 지급(Y) */
var SHEET_CP='쿠폰';   // 시즌 2는 seasonSetup()에서 '쿠폰 S2'
function cpSheet(){ return sheet(SHEET_CP, ['key','반','번호','이름','쿠폰','상태','신청시각','지급시각']); }
function couponClaim(p){
  if(!p.cls||!p.num||!p.cid) throw new Error('신청 정보가 부족해요');
  var s=cpSheet(), k=[String(p.cls).trim(),String(p.num).trim(),String(p.cid).trim()].join('|');
  var v=s.getDataRange().getValues();
  for(var i=1;i<v.length;i++){ if(String(v[i][0])===k) return {ok:true, st:String(v[i][5])}; }
  s.appendRow([k, p.cls, p.num, p.name||'', p.cid, 'W', new Date(), '']);
  return {ok:true, st:'W'};
}
function myCoupons(p){
  var s=cpSheet(), v=s.getDataRange().getValues(), out={};
  for(var i=1;i<v.length;i++){ if(String(v[i][1])===String(p.cls)&&String(v[i][2])===String(p.num)) out[String(v[i][4])]=String(v[i][5]); }
  return {ok:true, coupons:out};
}
function couponList(p){
  checkPin(p); var s=cpSheet(), v=s.getDataRange().getValues(), rows=[];
  for(var i=1;i<v.length;i++){ if(p.cls&&String(v[i][1])!==String(p.cls)) continue;
    rows.push({key:v[i][0], cls:v[i][1], num:v[i][2], name:v[i][3], cid:v[i][4], st:v[i][5], t:v[i][6]?new Date(v[i][6]).toISOString():''}); }
  return {ok:true, rows:rows};
}
function couponGive(p){
  checkPin(p); var s=cpSheet(), v=s.getDataRange().getValues();
  for(var i=1;i<v.length;i++){ if(String(v[i][0])===String(p.key)){ s.getRange(i+1,6).setValue(p.undo?'W':'Y'); s.getRange(i+1,8).setValue(p.undo?'':new Date()); return {ok:true}; } }
  throw new Error('신청을 찾지 못했어요');
}

/* ───────── 교신(질문) 게시판 ───────── */
var SHEET_Q = '교신', SHEET_A = '응답';   // 시즌 2는 '교신 S2'·'응답 S2'
function qid(){ return Utilities.getUuid().slice(0,8); }
function qadd(p){
  if(!p.text || String(p.text).trim().length<5) throw new Error('질문을 조금 더 자세히 적어 주세요');
  var s = sheet(SHEET_Q, ['id','시각','반','번호','이름','미션','단계','내용','익명','응답수']);
  var id = qid();
  s.appendRow([id, new Date(), p.cls||'', p.num||'', p.name||'', p.mission||'', p.phase||'', String(p.text).slice(0,1000), p.anon==='1'?'1':'', 0]);
  return {ok:true, id:id};
}
function qanswer(p){
  if(!p.qid || !p.text || String(p.text).trim().length<2) throw new Error('응답 내용을 적어 주세요');
  var s = sheet(SHEET_A, ['id','qid','시각','반','번호','이름','내용','센터장']);
  var t = (p.teacher==='1' && pinOk(p));   // 센터장 표시는 센터장 코드가 맞을 때만
  s.appendRow([qid(), p.qid, new Date(), t?'':(p.cls||''), t?'':(p.num||''), t?'센터장':(p.name||''), String(p.text).slice(0,1000), t?'1':'']);
  // 응답수 갱신
  var q = sheet(SHEET_Q, ['id','시각','반','번호','이름','미션','단계','내용','익명','응답수']);
  var v = q.getDataRange().getValues();
  for(var i=1;i<v.length;i++){ if(String(v[i][0])===String(p.qid)){ q.getRange(i+1,10).setValue(Number(v[i][9]||0)+1); break; } }
  return {ok:true};
}
/* 센터장: 응답 하나 지우기 (교신 탭의 🗑 버튼) */
function adel(p){
  checkPin(p);
  var s = sheet(SHEET_A, ['id','qid','시각','반','번호','이름','내용','센터장']);
  var v = s.getDataRange().getValues(), qidv = null;
  for(var i=v.length-1;i>=1;i--){ if(String(v[i][0])===String(p.id)){ qidv=String(v[i][1]); s.deleteRow(i+1); break; } }
  if(qidv===null) return {ok:false, error:'이미 지워졌거나 찾을 수 없는 응답이에요'};
  var q = sheet(SHEET_Q, ['id','시각','반','번호','이름','미션','단계','내용','익명','응답수']);
  var qv = q.getDataRange().getValues();
  for(var j=1;j<qv.length;j++){ if(String(qv[j][0])===qidv){ q.getRange(j+1,10).setValue(Math.max(0,Number(qv[j][9]||0)-1)); break; } }
  return {ok:true};
}
function qlist(p){
  var q = sheet(SHEET_Q, ['id','시각','반','번호','이름','미션','단계','내용','익명','응답수']);
  var a = sheet(SHEET_A, ['id','qid','시각','반','번호','이름','내용','센터장']);
  var qv = q.getDataRange().getValues(), av = a.getDataRange().getValues();
  var full = pinOk(p), nm = function(x){ return full ? String(x) : maskName(x); };
  var ans = {};
  for(var j=1;j<av.length;j++){ var r=av[j]; (ans[String(r[1])]=ans[String(r[1])]||[]).push({id:String(r[0]),t:r[2]?new Date(r[2]).toISOString():'',cls:r[3],num:r[4],name:r[5],text:r[6],teacher:String(r[7])==='1'}); }
  for(var aq in ans){ ans[aq].forEach(function(a){ if(!a.teacher) a.name = nm(a.name); }); }
  var rows=[];
  for(var i=qv.length-1;i>=1 && rows.length<300;i--){ var x=qv[i];
    if(p.cls && String(x[2])!==String(p.cls)) continue;
    if(p.mission && String(x[5])!==String(p.mission)) continue;
    var anon = String(x[8])==='1';
    rows.push({id:x[0], t:x[1]?new Date(x[1]).toISOString():'', cls:x[2], num:anon?'':x[3], name:anon?'익명 승무원':nm(x[4]), mine:(String(x[2])===String(p.cls)&&String(x[3])===String(p.num)), mission:x[5], phase:x[6], text:x[7], answers:ans[String(x[0])]||[]});
  }
  return {ok:true, rows:rows};
}

/* ───────── Gemini 피드백 (키는 시트 '설정' 탭에만 있음) ───────── */
var COACH_RULES =
 "너는 한국 중학교 1학년 과학 수업의 논술 코치 '아라'다. 학생(13세)의 글을 아래 '채점 기준'에 비추어 피드백한다.\n"+
 "규칙:\n"+
 "1) 반드시 중1 교과서 범위 안의 말만 쓴다. 금지: 전리층, 자기권, 코로나 질량 방출, 플라즈마, 태양 복사압, 밀도 공식 변형, 고등학교 수준 용어.\n"+
 "2) 채점 기준에 적힌 '평가 요소'를 하나씩 확인해 있는 것/빠진 것을 가려낸다. 과학적 오개념이 있으면 가장 먼저 짚는다.\n"+
 "3) 정답 문장을 통째로 써 주지 않는다. 빠진 요소는 '자료의 어느 부분을 보면 되는지' 힌트로 준다.\n"+
 "4) 글의 짜임을 본다: 주장(판단·분류) → 자료 속 근거(왜냐하면) → 연결·결론(따라서). 근거가 자료에서 나왔는지 확인한다.\n"+
 "5) 말투는 따뜻하고 짧은 존댓말. 칭찬은 구체적으로 1~2개, 고칠 점은 가장 중요한 것부터 최대 2개.\n"+
 "6) 예상 수준은 채점 기준의 수행 수준(상·중·하)으로 판단하되, 점수 숫자는 말하지 않는다.\n";
var COACH_RULES_S2 =
 "너는 한국 중학교 1학년 과학 수업의 관측 보고 코치 '아라'다. 단원은 '지구와 달'(지구의 자전과 일주 운동, 공전과 별자리 변화·태양의 연주 운동, 달의 위상, 일식과 월식)이다. 학생(13세)의 짧은 글을 아래 '채점 기준'에 비추어 피드백한다.\n"+
 "규칙:\n"+
 "1) 반드시 중1 교과서 범위 안의 말만 쓴다. 금지: 적경·적위, 시간각, 세차 운동, 자전축 경사에 따른 계절 변화, 궤도 경사·승교점, 사로스 주기, 본그림자·반그림자, 위도에 따른 일주 운동 차이, 전향력, 고등학교 수준 용어.\n"+
 "2) 채점 기준에 적힌 요소를 하나씩 확인해 있는 것/빠진 것을 가려낸다. 방향(동→서, 서→동)이나 천체 배열 순서처럼 과학적으로 틀린 내용이 있으면 가장 먼저 짚는다.\n"+
 "3) 정답 문장을 통째로 써 주지 않는다. 빠진 요소는 '시뮬레이션이나 그림의 어느 부분을 보면 되는지' 힌트로 준다.\n"+
 "4) 한두 문장짜리 보고이므로 글의 짜임은 엄격히 따지지 않는다. 원인(지구·달의 운동, 배열)과 결과(보이는 현상)가 이어지는지만 본다.\n"+
 "5) 말투는 따뜻하고 짧은 존댓말. 칭찬은 구체적으로 1개, 고칠 점은 가장 중요한 것 최대 1~2개.\n"+
 "6) 예상 수준은 상·중·하로 판단하되, 점수 숫자는 말하지 않는다.\n";
function feedback(p){
  var k = conf('GEMINI_KEY');
  if(!k) return {ok:false, error:'AI 키가 아직 설정되지 않았어요 (시트 설정 탭)'};
  var models = [conf('MODEL')||'gemini-3.6-flash','gemini-3.6-flash','gemini-2.5-flash'];
  var prompt = (S2 ? COACH_RULES_S2 : COACH_RULES)+
    "\n[채점 기준]\n"+(p.rubric||'')+"\n\n[학생 글]\n\"\"\"\n"+(p.text||'')+"\n\"\"\"\n\n"+
    '다음 JSON 형식으로만 답하라: {"level":"상|중|하","good":["구체적 칭찬 1~2개"],"improve":["가장 중요한 고칠 점부터 최대 2개. 빠진 평가 요소가 있으면 자료의 어디를 보라고 힌트"],"next":"학생이 바로 이어 쓸 수 있는 시작 문장 (빈칸 ___ 포함, 정답은 넣지 않음)"}';
  var lastErr = '';
  for(var i=0;i<models.length;i++){
    var m = models[i]; if(!m) continue;
    try{
      var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/'+m+':generateContent?key='+encodeURIComponent(k),{
        method:'post', contentType:'application/json', muteHttpExceptions:true,
        payload: JSON.stringify({contents:[{parts:[{text:prompt}]}], generationConfig:{responseMimeType:'application/json', temperature:0.4}})
      });
      var code = res.getResponseCode();
      if(code===404) continue;
      if(code!==200){ lastErr = 'HTTP '+code; if(code===429) continue; break; }
      var j = JSON.parse(res.getContentText());
      var txt = (((j.candidates||[])[0]||{}).content||{}).parts; txt = (txt||[]).map(function(x){return x.text||'';}).join('');
      var fb = JSON.parse(txt.replace(/^```json\s*|```\s*$/g,''));
      // 기록
      if(String(p.nolog||'')!=='1') try{ log({cls:p.cls,num:p.num,name:p.name,type:'AI피드백',item:p.item||'',text:p.text,score:p.score||'',ai:fb.level||'',aifb:JSON.stringify(fb)}); }catch(e){}
      return {ok:true, fb:fb, model:m};
    }catch(err){ lastErr = String(err); }
  }
  return {ok:false, error:'AI 응답 실패 '+lastErr};
}

/* ───────── 명예의 전당: 시즌 1은 290점, 시즌 2는 250점 이상, 점수 높은 순 ───────── */
function legends(p){
  var s = sheet(SHEET_STUDENTS, ['key','반','번호','이름','점수','등급','진도','상태JSON','최근접속']);
  var v = s.getDataRange().getValues(), full = pinOk(p), rows = [];
  for(var i=1;i<v.length;i++){ var pts = Number(v[i][4]||0);
    if(pts>=HOF_AT) rows.push({cls:v[i][1], num:v[i][2], name: full ? String(v[i][3]) : maskName(v[i][3]), pts:pts}); }
  rows.sort(function(a,b){ return b.pts-a.pts; });
  return {ok:true, rows:rows.slice(0,60)};
}

/* ───────── 센터장 ───────── */
function checkPin(p){ var pin = conf('PIN'); if(!pin || String(p.pin)!==pin) throw new Error('센터장 코드가 달라요'); }

function tlogin(p){ checkPin(p); return {ok:true, ai: !!conf('GEMINI_KEY')}; }

function roster(p){
  checkPin(p);
  var R=findRoster(), codes={}; if(R){ var rv=R.sh.getRange(R.head+1,1,Math.max(0,R.sh.getLastRow()-R.head),Math.max(R.ki+1,R.sh.getLastColumn())).getValues(); rv.forEach(function(x){ codes[String(x[R.ci]).trim()+'|'+String(x[R.ni]).trim()]=String(x[R.ki]||'').trim(); }); }
  var prev={}; if(S2){ var s1=ss().getSheetByName('학생'); if(s1){ s1.getDataRange().getValues().forEach(function(x,j){ if(j>0) prev[String(x[0])]=Number(x[4]||0); }); } }
  var s = sheet(SHEET_STUDENTS, ['key','반','번호','이름','점수','등급','진도','상태JSON','최근접속']);
  var v = s.getDataRange().getValues(), rows = [];
  for(var i=1;i<v.length;i++){
    if(p.cls && String(v[i][1])!==String(p.cls)) continue;
    rows.push({cls:v[i][1], num:v[i][2], name:v[i][3], pts:v[i][4], rank:v[i][5], prog:v[i][6], state:v[i][7]||'', seen:v[i][8]?new Date(v[i][8]).toISOString():'', code:codes[String(v[i][1]).trim()+'|'+String(v[i][2]).trim()]||'', prev:(S2 && (String(v[i][0]) in prev)) ? prev[String(v[i][0])] : ''});
  }
  return {ok:true, rows:rows};
}

function logs(p){
  checkPin(p);
  var s = sheet(SHEET_LOG, ['시각','반','번호','이름','종류','항목','내용','자동점검','AI수준','AI피드백']);
  var v = s.getDataRange().getValues(), rows = [];
  for(var i=v.length-1;i>=1 && rows.length<800;i--){
    if(p.cls && String(v[i][1])!==String(p.cls)) continue;
    if(p.num && String(v[i][2])!==String(p.num)) continue;
    rows.push({t:v[i][0]?new Date(v[i][0]).toISOString():'', cls:v[i][1], num:v[i][2], name:v[i][3], type:v[i][4], item:v[i][5], text:v[i][6], score:v[i][7], ai:v[i][8], aifb:v[i][9]});
  }
  return {ok:true, rows:rows};
}


/* ───────── 🌙 달 관측 앨범 (시즌 2) ─────────
   학년 전체가 하나의 달력을 채웁니다. '달앨범' 탭에 한 줄씩 쌓이고,
   작은 사진(썸네일)은 시트 칸에, 큰 사진은 내 드라이브 '아라호 달 앨범' 폴더에 저장됩니다.
   처음 한 번: Apps Script 편집기에서 함수 '달앨범준비'를 골라 ▶ 실행 → 드라이브 권한 허용. */
var MOON_H = ['id','올린 시각','찍은 시각','음력일','반','번호','이름','방향','한마디','사진','원본','같아요','달라요'];
function moonSheet(){ return sheet('달앨범', MOON_H); }
function moonFolder(){
  var it = DriveApp.getFoldersByName('아라호 달 앨범');
  return it.hasNext() ? it.next() : DriveApp.createFolder('아라호 달 앨범');
}
function 달앨범준비(){ moonSheet(); moonFolder(); }
function moonUp(p){
  var img = String(p.img||'');
  if(img.indexOf('data:image/jpeg;base64,')!==0 || img.length>49000) throw new Error('사진을 다시 골라 주세요');
  var t = new Date(p.t); if(isNaN(t)) throw new Error('찍은 날짜를 넣어 주세요');
  var id = qid(), link = '';
  try{
    if(p.full && String(p.full).indexOf('data:image/jpeg;base64,')===0){
      var b = Utilities.newBlob(Utilities.base64Decode(String(p.full).split(',')[1]), 'image/jpeg',
        Utilities.formatDate(t,'Asia/Seoul','yyyyMMdd_HHmm')+'_'+(p.cls||'')+'반'+(p.num||'')+'번.jpg');
      link = moonFolder().createFile(b).getUrl();
    }
  }catch(e){ link=''; }
  moonSheet().appendRow([id, new Date(), t, Number(p.lunar)||'', p.cls||'', p.num||'', p.name||'', p.dir||'', String(p.memo||'').slice(0,200), img, link, 0, 0]);
  return {ok:true, id:id};
}
function moonList(p){
  var v = moonSheet().getDataRange().getValues(), full = pinOk(p), L = [];
  for(var i=1;i<v.length;i++){ var r=v[i]; if(!r[0]) continue;
    var t = r[2] instanceof Date ? r[2].toISOString() : String(r[2]);
    L.push({id:String(r[0]), t:t, lunar:Number(r[3])||0, cls:String(r[4]), num:full?String(r[5]):'', name:full?String(r[6]):maskName(r[6]),
      mine: String(r[4])===String(p.cls||'') && String(r[5])===String(p.num||'') && String(r[6])===String(p.name||''),
      dir:String(r[7]), memo:String(r[8]), img:String(r[9]), orig: full?String(r[10]):'', same:Number(r[11])||0, diff:Number(r[12])||0});
  }
  return {ok:true, list:L};
}
function moonVote(p){
  var s = moonSheet(), v = s.getDataRange().getValues(), col = p.v==='diff' ? 13 : 12, d = Number(p.d)===-1 ? -1 : 1;
  for(var i=1;i<v.length;i++){ if(String(v[i][0])===String(p.id)){ s.getRange(i+1,col).setValue(Math.max(0,(Number(v[i][col-1])||0)+d)); return {ok:true}; } }
  return {ok:false, error:'사진을 찾을 수 없어요'};
}
/* 센터장(코드) 또는 올린 본인만 지울 수 있음 */
function moonDel(p){
  var s = moonSheet(), v = s.getDataRange().getValues(), t = pinOk(p);
  for(var i=v.length-1;i>=1;i--){ if(String(v[i][0])===String(p.id)){
    var own = String(v[i][4])===String(p.cls||'') && String(v[i][5])===String(p.num||'') && String(v[i][6])===String(p.name||'');
    if(!t && !own) throw new Error('내가 올린 사진만 지울 수 있어요');
    s.deleteRow(i+1); return {ok:true}; } }
  return {ok:false, error:'이미 지워졌어요'};
}
