const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const script = fs.readFileSync(path.join(__dirname, '../apps-script-atual/12_Speech_Integration_GE.js'), 'utf8');
const marker = script.indexOf('// Recuperação histórica isolada:');
assert.ok(marker >= 0);
const code = script.slice(marker);
const key = d => d.toISOString().slice(0, 10);
function setup(options = {}) {
  const before = Array.from({length:60}, (_,i)=>[0.5+i/100, 1+i/100]);
  const formulas = Array.from({length:60},(_,i)=>['', `=H${i+3}/B${i+3}`]);
  const oldContent = Array.from({length:60},()=>['PALMEIRAS','']);
  const state = {writes:[],released:0,marked:0,content:oldContent};
  const props = new Map();
  const source = {fileName:'transcricao261014.xlsx',programDate:new Date('2026-10-14T12:00:00Z'),lines:[{timestamp:'00:02',text:'jogo'},{timestamp:'00:40',text:'palmeiras'},{timestamp:'01:05',text:'corinthians'}]};
  const range = {getValues:()=>state.content,getA1Notation:()=> 'J3:K62',setValues:v=>{state.writes.push('J3:K62');state.content=v;}};
  const audience = {getValues:()=>before,getFormulas:()=>formulas};
  const sheet = {getRange:(r,c,n,m)=>{assert.equal(r,3);assert.equal(n,60);assert.equal(m,2);if(c===10)return range;if(c===8)return audience;throw Error('Unexpected sheet write/read');}};
  const timeline = Array.from({length:60},(_,i)=>({hora:`18:${String(i).padStart(2,'0')}`,tema:'CORINTHIANS',termos:'jogo, corinthians'}));
  const FixedDate = class extends Date { constructor(...a){super(...(a.length?a:['2026-10-15T19:00:00Z']));} };
  const ctx = {Date:FixedDate, CFG:{TIMEZONE:'America/Sao_Paulo'}, Utilities:{parseDate:text=>{const [d,m,y]=text.split('/').map(Number);return new FixedDate(y,m-1,d,12);},formatDate:(date,tz,fmt)=>fmt==='dd/MM/yyyy'? `${String(date.getDate()).padStart(2,'0')}/${String(date.getMonth()+1).padStart(2,'0')}/${date.getFullYear()}`:key(date)},
    LockService:{getScriptLock:()=>({waitLock:()=>{},releaseLock:()=>state.released++})},
    formatDateKey_:key, localizarBlocoHistoricoGEPorData_:()=>options.noBlock?null:{sheet,temaCol:10,termosCol:11,audienciaCol:8},
    buscarArquivoSpeechDoDiaNoDrive_:()=>options.noSource?null:{...source,programDate:options.wrongDate?new FixedDate('2026-10-13T12:00:00Z'):source.programDate},
    buildSpeechSignature_:()=> '2026-10-14|DRIVE|file|123',PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k),setProperty:(k,v)=>{props.set(k,v);state.marked++;}})},
    processAndMapTimestampsGE_:input=>{assert.equal(input.rows.length,2);assert.equal(input.rows[0].texto,'jogo palmeiras');return {rows:input.rows};},
    getStopwords_:()=>new Set(),computeTopTermsPerMinute_:()=>({}),buildTimelineGE_:()=>options.short?timeline.slice(0,59):options.duplicate?[...timeline.slice(0,59),timeline[0]]:timeline,
    SpreadsheetApp:{flush:()=>{}},logGE_:()=>{},Logger:{log:()=>{}}
  };
  vm.createContext(ctx);vm.runInContext(code,ctx);
  return {ctx,state,before,formulas};
}
{
 const {ctx,state,before,formulas}=setup();const originalAudience=JSON.stringify([before,formulas]);
 const result=ctx.reprocessarSpeechPorDataGE_('14/10/2026');
 assert.equal(result.success,true);assert.equal(result.minutes,60);assert.equal(result.audiencePreserved,true);
 assert.deepEqual(state.writes,['J3:K62']);assert.equal(JSON.stringify([before,formulas]),originalAudience);
 assert.equal(state.released,1);assert.equal(state.marked,1);
 assert.equal(ctx.reprocessarSpeechPorDataGE_('14/10/2026').skipped,true);assert.equal(state.writes.length,1);
}
for (const options of [{noBlock:true},{noSource:true},{wrongDate:true},{short:true},{duplicate:true}]) {
 const {ctx,state}=setup(options);assert.throws(()=>ctx.reprocessarSpeechPorDataGE_('14/10/2026'));assert.equal(state.writes.length,0);assert.equal(state.marked,0);assert.equal(state.released,1);
}
for (const date of ['14-10-2026','31/02/2026','15/10/2026','16/10/2026']) {
 const {ctx,state}=setup();assert.throws(()=>ctx.reprocessarSpeechPorDataGE_(date));assert.equal(state.writes.length,0);
}
// Equivalência com o parser real atualmente publicado para MM:SS e HH:MM:SS.
{
 const {ctx}=setup();const lines=[{timestamp:'00:12',text:'um'},{timestamp:'00:35',text:'dois'},{timestamp:'01:02:30',text:'tres'}];
 ctx.SpreadsheetApp={getActive:()=>({getSheetByName:()=>({getLastRow:()=>3,getLastColumn:()=>1,getRange:()=>({getValues:()=>lines.map(l=>[`(${l.timestamp}) ${l.text}`])})})})};
 ctx.normalizeText_=x=>x.toLowerCase();vm.runInContext(fs.readFileSync(path.join(__dirname, '../apps-script-atual/10_Input_GE.js'),'utf8'),ctx);
 assert.equal(JSON.stringify(ctx.montarInputSpeechHistoricoGE_({lines})),JSON.stringify(ctx.loadGEInputFromSheet_()));
}
// A busca só abre os XLSX da data pedida, incluindo duplicatas com sufixo.
{
 const opened=[];
 const names=['transcricao261013.xlsx','transcricao261013_2.xlsx','transcricao261014.xlsx','transcricao261014_2.xlsx'];
 const files=names.map((name,i)=>({getName:()=>name,getMimeType:()=> 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',getDateCreated:()=>new Date('2026-10-15T12:00:00Z'),getLastUpdated:()=>new Date(1792065600000+i*1000),getId:()=>String(i),getBlob:()=>{opened.push(name);return name;}}));
 let index=0;
 const sandbox={Date,CFG:{TIMEZONE:'America/Sao_Paulo',SPEECH_DRIVE_FOLDER_ID:'folder'},Logger:{log:()=>{}},Utilities:{formatDate:key},PropertiesService:{getScriptProperties:()=>({getProperty:()=>null})},DriveApp:{getFolderById:()=>({getFiles:()=>({hasNext:()=>index<files.length,next:()=>files[index++]})})}};
 vm.createContext(sandbox);vm.runInContext(script,sandbox);
 sandbox.parseSpeechXlsxBlob_=()=>({youtubeTitle:'Programa',lines:[{timestamp:'00:01',text:'jogo'}]});
 const chosen=sandbox.buscarArquivoSpeechDoDiaNoDrive_({expectedDate:new Date('2026-10-14T12:00:00Z')});
 assert.deepEqual(opened,['transcricao261014.xlsx','transcricao261014_2.xlsx']);
 assert.equal(chosen.fileName,'transcricao261014_2.xlsx');assert.equal(key(chosen.programDate),'2026-10-14');
 const otherFormat=sandbox.extrairDatasCandidatasDoNomeArquivoSpeech_('transcricao141026.xlsx');
 assert.ok(otherFormat.some(d=>key(d)==='2026-10-14'));
 const invalid=sandbox.extrairDatasCandidatasDoNomeArquivoSpeech_('transcricao310226.xlsx');
 assert.ok(!invalid.some(d=>key(d)==='2026-03-03'));
}
console.log('Recovery checks passed: write scope, audience/formulas, idempotence, validation, parser equivalence, date filtering and numbered duplicates.');

