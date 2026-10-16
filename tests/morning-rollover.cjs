const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(process.argv[2] || 'footer-history-after.js', 'utf8');
class Sheet {
  constructor() { this.cells = new Map(); this.writes = []; this.insertions = 0; }
  get(r,c) { return this.cells.get(`${r}:${c}`) ?? ''; }
  put(r,c,v) { this.cells.set(`${r}:${c}`,v); }
  getLastColumn() { return 15; }
  getMaxRows() { return 106; }
  getSheetId() { return 1117157422; }
  getParent() { return {getId:()=> 'radar-test'}; }
  getRange(r,c,n=1,m=1) {
    const s=this, range={
      getValues:()=>Array.from({length:n},(_,i)=>Array.from({length:m},(_,j)=>{const v=s.get(r+i,c+j);return typeof v==='string' && v.startsWith('=')?(s.effective?.get(`${r+i}:${c+j}`) ?? ''):v;})),
      getFormulas:()=>Array.from({length:n},(_,i)=>Array.from({length:m},(_,j)=>{const v=s.get(r+i,c+j);return typeof v==='string' && v.startsWith('=')?v:'';})),
      getValue:()=>range.getValues()[0][0],getFormula:()=>range.getFormulas()[0][0],getDisplayValue:()=>String(s.get(r,c)),
      setValues:values=>{s.writes.push({r,c,n,m});for(let i=0;i<n;i++)for(let j=0;j<m;j++)s.put(r+i,c+j,values[i][j]);return range;},
      setValue:value=>range.setValues([[value]]),setFormula:value=>range.setValue(value),
      clearContent:()=>range.setValues(Array.from({length:n},()=>Array(m).fill('')))
    };
    for(const method of ['setNumberFormat','setFontWeight','setHorizontalAlignment','setBackground','setFontColor','breakApart','merge','copyTo'])range[method]=()=>range;
    return range;
  }
  setColumnWidth() {}
  insertColumns(start,count) {
    this.insertions++;
    const next=new Map();
    for (const [key,value] of this.cells) {
      const [r,c]=key.split(':').map(Number);
      // Reproduz o deslocamento que transformava referências D em H.
      const shifted=typeof value==='string' && value.startsWith('=') ? value.replace(/\b([D-Z])(\d+)\b/g,(_,letter,row)=>String.fromCharCode(letter.charCodeAt(0)+count)+row):value;
      next.set(`${r}:${c>=start?c+count:c}`,shifted);
    }
    this.cells=next;
  }
}
function setup(existing) {
  const sheet=new Sheet();
  sheet.put(1,4,existing?'16/10/2026':'15/10/2026');
  sheet.put(1,1,'=ROUND(0,3743+0,7387*H3;2)');
  sheet.put(1,2,'="entre "&TEXT(H3;"0.00")');
  for(let row=3;row<=62;row++) {
    sheet.put(row,2,'=AVERAGE(H'+row+';L'+row+')');
    sheet.put(row,3,'=H'+row+'/B'+row+'-1');
    sheet.put(row,4,0.8+row/1000);
    sheet.put(row,5,'=H'+row+'/$B'+row);
    sheet.put(row,8,0.6+row/1000);
    sheet.put(row,10,'PALMEIRAS');
    sheet.put(row,11,'palmeiras, jogo');
  }
  sheet.put(63,4,'=AVERAGE(D3:D62)');sheet.put(64,4,'COMENTARISTA 1');
  sheet.put(70,2,'=D63');sheet.put(86,2,'=D63');
  const ctx={Date,CFG:{GE_OUT_HISTORICO:'Histórico',TIMEZONE:'America/Sao_Paulo'},SpreadsheetApp:{getActive:()=>({getSheetByName:()=>sheet}),flush:()=>{},CopyPasteType:{PASTE_FORMAT:'format'}},LockService:{getDocumentLock:()=>({tryLock:()=>true,releaseLock:()=>{}})},Utilities:{formatDate:()=> '16/10/2026'},PropertiesService:{getScriptProperties:()=>({setProperty:()=>{}})},Logger:{log:()=>{}}};
  vm.createContext(ctx);vm.runInContext(source,ctx);
  Object.assign(ctx,{localizarFormulaGE_:x=>x,localizarColunaDataHistoricoGE_:(s,date)=>s.get(1,4)===date?4:-1,localizarPainelGE_:()=>({rowSemana:70,rowMesAtual:86}),gerarFormulaMedianaOffsetGE_:()=> '=MEDIAN(H63)',garantirCamposHistoricoGE_:()=>{},logGE_:()=>{},restaurarBlocoMensalAposInsercaoGE_:()=>{}});
  return {sheet,ctx};
}
function checkFormulas(sheet) {
 assert.equal(sheet.get(1,1),'=ROUND(0,3743+0,7387*D3;2)');
 assert.ok(sheet.get(1,2).includes('*D3'));assert.ok(!/\bH3\b/.test(sheet.get(1,2)));
 for(let row=3;row<=62;row++) {
   assert.equal(sheet.get(row,3),`=IF(OR(B${row}="";B${row}=0;D${row}="");"";D${row}/B${row}-1)`);
   assert.equal(sheet.get(row,5),`=IF(OR(D${row}="";$B${row}=0);"";D${row}/$B${row})`);
 }
}
// Repetir a rotina no dia já aberto corrige todos os indicadores sem nova
// coluna e sem escrever em B, audiência, temas ou termos.
{
 const {sheet,ctx}=setup(true),before=new Map(sheet.cells);
 const result=ctx.criarOuRepararBlocoDiaAtualGE_(new Date('2026-10-16T12:00:00Z'));
 assert.equal(result.jaEstavaOk,true);assert.equal(sheet.insertions,0);checkFormulas(sheet);
 assert.deepEqual(sheet.writes.map(x=>[x.r,x.c,x.n,x.m]),[[63,4,1,3],[64,4,1,3],[65,4,1,1],[1,1,1,2],[3,3,60,1],[3,5,60,1]]);
 for(let row=3;row<=62;row++)for(const col of [2,4,8,10,11])assert.equal(sheet.get(row,col),before.get(`${row}:${col}`));
 const once=JSON.stringify([...sheet.cells]);ctx.criarOuRepararBlocoDiaAtualGE_(new Date());assert.equal(JSON.stringify([...sheet.cells]),once);
}
// Criar o próximo dia desloca o arquivo anterior e então reancora as
// referências em D, mantendo os valores antigos arquivados em H.
{
 const {sheet,ctx}=setup(false),previous=Array.from({length:60},(_,i)=>sheet.get(i+3,4));
 const result=ctx.criarOuRepararBlocoDiaAtualGE_(new Date('2026-10-16T12:00:00Z'));
 assert.equal(result.criou,true);assert.equal(sheet.insertions,1);checkFormulas(sheet);
 for(let i=0;i<60;i++){assert.equal(sheet.get(i+3,8),previous[i]);assert.equal(sheet.get(i+3,4),'');}
 ctx.criarOuRepararBlocoDiaAtualGE_(new Date());assert.equal(sheet.insertions,1);checkFormulas(sheet);
}
// A API grava os três intervalos em uma chamada, alterando só fórmulas.
{
 const {sheet,ctx}=setup(true);let calls=0;
 ctx.Sheets={Spreadsheets:{batchUpdate:(body,id)=>{
   calls++;assert.equal(id,'radar-test');assert.equal(body.requests.length,3);
   for(const request of body.requests){
     const u=request.updateCells;assert.equal(u.fields,'userEnteredValue');
     assert.equal(u.range.sheetId,1117157422);
     assert.equal(u.rows.length,u.range.endRowIndex-u.range.startRowIndex);
     const values=u.rows.map(row=>row.values.map(cell=>cell.userEnteredValue.formulaValue));
     sheet.getRange(u.range.startRowIndex+1,u.range.startColumnIndex+1,values.length,values[0].length).setValues(values);
   }
 }}};
 const result=ctx.atualizarReferenciasDiaAtualHistoricoGE_(sheet);
 assert.equal(calls,1);assert.equal(result.gravacao,'api_atomica');checkFormulas(sheet);
}
// Ausência de API habilitada mantém a compatibilidade com a gravação nativa.
{
 const {sheet,ctx}=setup(true);ctx.Sheets={Spreadsheets:{batchUpdate:()=>{throw Error('API unavailable');}}};
 const result=ctx.atualizarReferenciasDiaAtualHistoricoGE_(sheet);
 assert.equal(result.gravacao,'nativa');checkFormulas(sheet);
}
console.log('Morning rollover checks passed: new day, fast repeat, 60 rows, exact share formula, idempotence, preserved history and atomic API/fallback.');

// Todos os dias úteis usam a coluna correta do Resumo e a data do bloco,
// mesmo quando o relógio da execução aponta para outro dia.
for (const [day,col,label] of [
 ['19/10/2026','B','2ª'],['20/10/2026','C','3ª'],['21/10/2026','D','4ª'],
 ['22/10/2026','E','5ª'],['23/10/2026','F','6ª']
]) {
 const {sheet,ctx}=setup(true);sheet.put(1,4,day);
 const result=ctx.repararRodapeDiaAtualHistoricoGE_(sheet,new Date('2026-10-16T12:00:00Z'));
 assert.equal(result.gravacao,'nativa');
 assert.equal(sheet.get(63,4),'=IFERROR(AVERAGE(D3:D62);"")');
 assert.equal(sheet.get(63,5),`=IFERROR(MEDIAN('Resumo'!${col}3:${col});"")`);
 assert.equal(sheet.get(63,6),'=IF(OR(D63="";E63="";E63=0);"";D63/E63-1)');
 assert.equal(sheet.get(64,5),'MEDIANA '+label);
 assert.equal(sheet.get(64,4),'COMENTARISTA 1');assert.equal(sheet.get(65,4),'COMENTARISTA 2');
 assert.equal(sheet.get(64,6),'DIFERENÇA');
 // Sem limite fixo 42: linhas adicionadas nas próximas semanas são incluídas.
 assert.ok(!/3:[B-F]\d/.test(sheet.get(63,5)));
}
// Um rodapé parcialmente apagado é restaurado inclusive na saída rápida.
{
 const {sheet,ctx}=setup(true);
 sheet.put(63,4,'');sheet.put(63,5,'=MEDIAN(H63)');sheet.put(63,6,'');
 sheet.put(64,4,'');sheet.put(64,5,'MEDIANA 5ª');sheet.put(65,4,'');
 const result=ctx.criarOuRepararBlocoDiaAtualGE_(new Date('2026-10-16T12:00:00Z'));
 assert.equal(result.jaEstavaOk,true);assert.equal(sheet.insertions,0);
 assert.equal(sheet.get(63,5),'=IFERROR(MEDIAN(\'Resumo\'!F3:F);"")');
 assert.equal(sheet.get(64,5),'MEDIANA 6ª');assert.equal(sheet.get(65,4),'COMENTARISTA 2');
}
// Repetições preservam nomes, observações E65/F65, audiência e o arquivo.
{
 const {sheet,ctx}=setup(true);
 sheet.put(64,4,'GARRAFFA');sheet.put(65,4,'MULLER');sheet.put(65,5,'anotação');sheet.put(65,6,17);
 const before=new Map(sheet.cells);
 ctx.repararRodapeDiaAtualHistoricoGE_(sheet);ctx.repararRodapeDiaAtualHistoricoGE_(sheet);
 for(const [key,value] of before) {
  const [r,c]=key.split(':').map(Number);
  if((r===63&&c>=4&&c<=6)||(r===64&&c>=5&&c<=6))continue;
  assert.equal(sheet.get(r,c),value,'preservar '+key);
 }
}
function atomicApi(sheet,ctx) {
 const calls=[];
 ctx.Sheets={Spreadsheets:{batchUpdate:(body,id)=>{
  assert.equal(id,'radar-test');calls.push(body);
  for(const request of body.requests) {
   assert.equal(Object.keys(request).length,1);
   const u=request.updateCells,range=u.range;
   assert.equal(u.rows.length,range.endRowIndex-range.startRowIndex);
   for(let i=0;i<u.rows.length;i++) {
    assert.equal(u.rows[i].values.length,range.endColumnIndex-range.startColumnIndex);
    u.rows[i].values.forEach((cell,j)=>{
     const v=cell.userEnteredValue;
     sheet.put(range.startRowIndex+i+1,range.startColumnIndex+j+1,v?.formulaValue??v?.stringValue??v?.numberValue??'');
    });
   }
  }
 }}};
 return calls;
}
// O lote atômico usa máscaras precisas e formata apenas a linha de resultados.
{
 const {sheet,ctx}=setup(true),calls=atomicApi(sheet,ctx);
 const result=ctx.repararRodapeDiaAtualHistoricoGE_(sheet);
 assert.equal(result.gravacao,'api_atomica');assert.equal(calls.length,1);
 assert.equal(calls[0].requests.length,3);
 const requests=calls[0].requests.map(x=>x.updateCells);
 assert.equal(requests[0].fields,'userEnteredValue,userEnteredFormat.numberFormat');
 assert.deepEqual(Array.from(requests[0].rows[0].values,x=>x.userEnteredFormat.numberFormat.type),['NUMBER','NUMBER','PERCENT']);
 assert.equal(requests[1].fields,'userEnteredValue');assert.equal(requests[2].range.endColumnIndex,4);
 assert.equal(sheet.get(64,5),'MEDIANA 6ª');
}
// API indisponível ainda deixa o bloco completo pela via nativa.
{
 const {sheet,ctx}=setup(true);
 ctx.Sheets={Spreadsheets:{batchUpdate:()=>{throw Error('API unavailable');}}};
 assert.equal(ctx.repararRodapeDiaAtualHistoricoGE_(sheet).gravacao,'nativa');
 assert.equal(sheet.get(65,4),'COMENTARISTA 2');
}
// Ao abrir o próximo dia, o valor da mediana anterior fica arquivado em I63.
for(const api of [false,true]) {
 const {sheet,ctx}=setup(false);
 sheet.effective=new Map([['63:5',0.6846666666666666]]);
 sheet.put(63,5,'=IFERROR(MEDIAN(\'Resumo\'!E3:E);"")');
 if(api)atomicApi(sheet,ctx);
 ctx.criarOuRepararBlocoDiaAtualGE_(new Date('2026-10-16T12:00:00Z'));
 assert.equal(sheet.get(63,9),0.6846666666666666);
 assert.equal(sheet.get(63,5),'=IFERROR(MEDIAN(\'Resumo\'!F3:F);"")');
 assert.equal(sheet.insertions,1);
}
// Medianas vazias/zero são arquivadas sem fabricar um dado.
for(const value of ['',0]) {
 const {sheet,ctx}=setup(true);sheet.effective=new Map([['63:5',value]]);
 sheet.put(63,5,'=IFERROR(MEDIAN(\'Resumo\'!F3:F);"")');
 ctx.congelarMedianaDiaEncerradoHistoricoGE_(sheet);assert.equal(sheet.get(63,5),value);
}
// A abertura manual usa o mesmo reparo, inclusive se o dia já estiver aberto.
{
 const {sheet,ctx}=setup(true);sheet.put(65,4,'');
 assert.equal(ctx.adicionarNovoDiaCoreGE_(new Date('2026-10-16T12:00:00Z')).motivo,'duplicada');
 assert.equal(sheet.get(65,4),'COMENTARISTA 2');assert.equal(sheet.get(64,5),'MEDIANA 6ª');
}
console.log('Footer checks passed: weekdays B:F, D1 date, dynamic weeks, guarded formulas, full repeat repair, atomic write/fallback, preserved names and frozen archive median.');
// Um cabeçalho Date perto da meia-noite segue o calendário de São Paulo.
{
 const {sheet,ctx}=setup(true);
 ctx.Utilities.formatDate=date=>new Intl.DateTimeFormat('en-GB',{timeZone:'America/Sao_Paulo'}).format(date);
 sheet.put(1,4,new Date('2026-10-17T02:30:00Z'));
 ctx.repararRodapeDiaAtualHistoricoGE_(sheet,new Date('2026-10-19T12:00:00Z'));
 assert.equal(sheet.get(64,5),'MEDIANA 6ª');
 assert.equal(sheet.get(63,5),'=IFERROR(MEDIAN(\'Resumo\'!F3:F);"")');
}
