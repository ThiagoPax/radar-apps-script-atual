const assert=require('node:assert/strict'), fs=require('node:fs'), vm=require('node:vm');
const source=fs.readFileSync(process.argv[2]||'monthly-history-after.js','utf8');
const months=['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
const refs=[.56,.63,.71,.70,.65,.66,.69,.72,.60,.58,.65,.50];
class Sheet {
 constructor(){this.cells=new Map;this.maxRows=106;this.insertions=0;this.calls=[];this.rules=[];}
 autoResizeRows(start,count){this.resized={start,count};}
 get(r,c){return this.cells.get(`${r}:${c}`)??'';}
 put(r,c,v){this.cells.set(`${r}:${c}`,v);}
 getMaxRows(){return this.maxRows;}
 getLastColumn(){return 47;}
 getLastRow(){return Math.max(...[...this.cells.keys()].map(k=>+k.split(':')[0]));}
 getSheetId(){return 1117157422;}
 getParent(){return{getId:()=> 'radar-test'};}
 getRange(r,c,n=1,m=1){const s=this;return{
  getValues:()=>Array.from({length:n},(_,i)=>Array.from({length:m},(_,j)=>s.get(r+i,c+j))),
  getDisplayValues:()=>Array.from({length:n},(_,i)=>Array.from({length:m},(_,j)=>String(s.get(r+i,c+j)))),
  getFormula:()=>String(s.get(r,c)).startsWith('=')?s.get(r,c):'',
  getValue:()=>s.get(r,c)
 };}
 insert(row){const next=new Map;for(const[k,v]of this.cells){const[r,c]=k.split(':').map(Number);next.set(`${r>=row?r+1:r}:${c}`,v);}this.cells=next;this.maxRows++;this.insertions++;}
}
function setup(){
 const sheet=new Sheet;
 sheet.put(1,4,new Date('2026-10-16T12:00:00Z'));
 for(let m=0;m<9;m++)sheet.put(1,8+m*4,new Date(Date.UTC(2026,8-m,5,12)));
 sheet.put(1,44,new Date('2026-03-06T12:00:00Z')); // cabeçalho da linha 2 ausente
 sheet.put(66,1,'SEM. ATUAL (12/10-16/10)');sheet.put(97,4,'MENSAL');
 for(let i=0;i<8;i++){sheet.put(98+i,1,months[9-i]+(i===0?' (mês atual)':''));sheet.put(98+i,4,999);sheet.put(98+i,5,'=(#REF!/D'+(98+i)+')-1');}
 sheet.put(106,1,'JANEIRO');sheet.put(106,3,'-%');sheet.put(106,4,.56);
 sheet.put(63,4,'fórmula diária preservada');sheet.put(63,8,.74);sheet.put(65,4,'MULLER');
 const ctx={Date,CFG:{TIMEZONE:'America/Sao_Paulo',GE_OUT_HISTORICO:'Histórico'},Utilities:{formatDate:d=>new Intl.DateTimeFormat('en-GB',{timeZone:'America/Sao_Paulo'}).format(d)},Logger:{log:()=>{}},SpreadsheetApp:{getActive:()=>({getId:()=> 'radar-test',getSheetByName:()=>sheet}),flush:()=>{}},LockService:{getDocumentLock:()=>({tryLock:()=>true,releaseLock:()=>{}})},PropertiesService:{getScriptProperties:()=>({setProperty:(key,value)=>{ctx.savedProperty={key,value};}})}};
 ctx.Sheets={Spreadsheets:{get:(id,options)=>{
  assert.ok(+options.ranges[0].split(':E')[1]<=sheet.maxRows);
  return{sheets:[{properties:{sheetId:1117157422},conditionalFormats:sheet.rules}]};
 },batchUpdate:(body,id)=>{
  assert.equal(id,'radar-test');sheet.calls.push(body);
  for(const req of body.requests){
   assert.equal(Object.keys(req).length,1);
   if(req.deleteConditionalFormatRule)sheet.rules.splice(req.deleteConditionalFormatRule.index,1);
   if(req.addConditionalFormatRule)sheet.rules.splice(req.addConditionalFormatRule.index,0,req.addConditionalFormatRule.rule);
   if(req.insertDimension){const r=req.insertDimension.range;assert.equal(r.dimension,'ROWS');assert.equal(r.endIndex-r.startIndex,1);sheet.insert(r.startIndex+1);}
   if(req.updateCells){const u=req.updateCells,r=u.range;assert.equal(u.rows.length,r.endRowIndex-r.startRowIndex);assert.ok(r.endRowIndex<=sheet.maxRows);
    u.rows.forEach((row,i)=>{assert.equal(row.values.length,r.endColumnIndex-r.startColumnIndex);row.values.forEach((cell,j)=>{const v=cell.userEnteredValue;sheet.put(r.startRowIndex+i+1,r.startColumnIndex+j+1,v?.formulaValue??v?.stringValue??v?.numberValue??'');});});}
  }
 }}};
 ctx.columnToLetter_=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
 vm.createContext(ctx);vm.runInContext(source,ctx);return{sheet,ctx};
}
// Repara exatamente o caso real: outubro a março, fevereiro ausente e janeiro.
{
 const {sheet,ctx}=setup(),beforeDaily=sheet.get(63,4),result=ctx.repararPainelMensalHistoricoGE_(sheet);
 assert.equal(result.inseridos,1);assert.equal(result.fim,107);assert.equal(sheet.get(106,1),'FEVEREIRO');assert.equal(sheet.get(107,1),'JANEIRO');
 assert.equal(sheet.calls.length,1);assert.equal(sheet.get(63,4),beforeDaily);assert.equal(sheet.get(65,4),'MULLER');
 for(let i=0;i<10;i++){
  const r=98+i;assert.equal(sheet.get(r,4),refs[9-i]);assert.ok(sheet.get(r,5).includes(`B${r}/D${r}-1`));assert.ok(!sheet.get(r,5).includes('#REF!'));
  assert.ok(!/TODAY|ROW\(/.test(sheet.get(r,1)+sheet.get(r,2)));
 }
 assert.equal(sheet.get(107,3),'-%');
 assert.ok(sheet.get(105,2).includes('AR63'),'06/03 incluído apesar do cabeçalho ausente');
 const once=JSON.stringify([...sheet.cells]);ctx.repararPainelMensalHistoricoGE_(sheet);
 assert.equal(sheet.insertions,1);assert.equal(JSON.stringify([...sheet.cells]),once);
}
// Uma linha por novo mês; novembro/dezembro usam os valores guardados de 2025.
{
 const {sheet,ctx}=setup();ctx.repararPainelMensalHistoricoGE_(sheet);
 sheet.put(1,4,new Date('2026-11-02T12:00:00Z'));
 assert.equal(ctx.repararPainelMensalHistoricoGE_(sheet).inseridos,1);
 assert.equal(sheet.get(98,1),'NOVEMBRO (mês atual)');assert.equal(sheet.get(98,4),.65);
 sheet.put(1,4,new Date('2026-12-01T12:00:00Z'));
 assert.equal(ctx.repararPainelMensalHistoricoGE_(sheet).inseridos,1);assert.equal(sheet.get(98,4),.50);
 assert.equal(sheet.get(109,1),'JANEIRO');assert.equal(sheet.get(109,3),'-%');
}
// Pular um mês na execução não elimina o mês intermediário.
{
 const {sheet,ctx}=setup();ctx.repararPainelMensalHistoricoGE_(sheet);sheet.put(1,4,new Date('2026-12-01T12:00:00Z'));
 assert.equal(ctx.repararPainelMensalHistoricoGE_(sheet).inseridos,2);assert.equal(sheet.get(99,1),'NOVEMBRO');
}
// O limite acompanha os rótulos mesmo sem a sentinela; limpeza preserva janeiro.
{
 const {sheet,ctx}=setup();sheet.put(106,3,'');assert.equal(ctx.localizarPainelGE_(sheet).panelEnd,106);
 ctx.repararPainelMensalHistoricoGE_(sheet);assert.equal(ctx.localizarPainelGE_(sheet).panelEnd,107);
}
// A cópia H:I -> D:E termina repondo E com referência explícita a B e D.
{
 const {sheet,ctx}=setup();sheet.put(98,8,.58);sheet.put(98,9,'=B98/H98-1');
 ctx.restaurarBlocoMensalAposInsercaoGE_(sheet);
 const req=sheet.calls[0].requests.at(-1).updateCells;
 assert.equal(req.range.startColumnIndex,4);assert.ok(req.rows[0].values[0].userEnteredValue.formulaValue.includes('B98/D98'));
}
// Os 12 números fornecidos ficam persistidos pelo reparo público.
{
 const {ctx}=setup();ctx.repararPainelMensalHistoricoGE();assert.deepEqual(JSON.parse(ctx.savedProperty.value),refs);
}
console.log('Monthly checks passed: missing February, all 2025 references, guarded YoY, stable month labels, no duplication, future November/December, skipped months, panel bounds and daily-column restore.');
// As regras do Radar não se duplicam, migram de I para E e crescem com a tabela.
{
 const {sheet,ctx}=setup();const unrelated={ranges:[{sheetId:1117157422,startRowIndex:2,endRowIndex:62,startColumnIndex:2,endColumnIndex:3}],gradientRule:{minpoint:{type:'MIN'}}};
 sheet.rules=[unrelated,{ranges:[{sheetId:1117157422,startRowIndex:97,endRowIndex:107,startColumnIndex:8,endColumnIndex:9}],booleanRule:{condition:{type:'CUSTOM_FORMULA',values:[{userEnteredValue:'=N("RADAR_MENSAL_ANUAL_POS")=0'}]}}}];
 ctx.repararPainelMensalHistoricoGE_(sheet);ctx.repararPainelMensalHistoricoGE_(sheet);
 assert.equal(sheet.rules.length,3);assert.equal(sheet.rules[2],unrelated);
 assert.equal(sheet.resized.start,98);assert.equal(sheet.resized.count,10);
 sheet.put(1,4,new Date('2026-11-02T12:00:00Z'));ctx.repararPainelMensalHistoricoGE_(sheet);
 assert.equal(sheet.rules.length,3);assert.equal(sheet.rules[0].ranges[0].endRowIndex,108);assert.equal(sheet.rules[0].ranges[0].startColumnIndex,4);
}
