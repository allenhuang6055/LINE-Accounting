'use strict';
// V2 is a single atomic system-sheet record per month, split into bounded cells.
// No transaction rows, historical month tabs, or legacy snapshots are migrated.
const {createHash,randomUUID}=require('node:crypto');
const TAB='系統同步快照V2',HEADERS=['月份','格式版本','代次','分片數','SHA256','更新時間UTC'];
const VERSION='snapshot-v2',CHUNK=30000,MAX_PAYLOAD=6000000;
const ID=/^TX-[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const digest=s=>createHash('sha256').update(s).digest('hex');
function col(n){let s='';while(n){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26);}return s;}
function validate(p){
 if(!p||!Number.isSafeInteger(p.gridRows)||p.gridRows<1||!Array.isArray(p.ids))throw Error('V2 快照結構錯誤');
 const ids=new Set(p.ids);
 if(ids.size!==p.ids.length||p.ids.some(id=>!ID.test(id))||!p.positions||typeof p.positions!=='object'||Array.isArray(p.positions)||Object.keys(p.positions).length!==ids.size||p.ids.some(id=>!Number.isSafeInteger(p.positions[id])||p.positions[id]<2))throw Error('V2 快照交易或位置錯誤');
 if(!p.identities||typeof p.identities!=='object'||Array.isArray(p.identities)||Object.entries(p.identities).some(([id,v])=>!ids.has(id)||!Number.isInteger(v?.metadataId)||v.metadataId<1||!Number.isInteger(v.sheetId)||v.sheetId<0))throw Error('V2 快照身分錯誤');
 return p;
}
async function topology(sheets,book){return (await sheets.spreadsheets.get({spreadsheetId:book,fields:'sheets(properties(sheetId,title,gridProperties))'})).data.sheets||[];}
async function read(sheets,book){
 const tab=(await topology(sheets,book)).find(t=>t.properties.title===TAB)?.properties,out=new Map();if(!tab)return out;
 const rows=(await sheets.spreadsheets.values.get({spreadsheetId:book,range:`'${TAB}'!A2:${col(tab.gridProperties?.columnCount||26)}`,valueRenderOption:'UNFORMATTED_VALUE'})).data.values||[];
 rows.forEach((r,i)=>{
  if(!r.length||r.every(x=>x===''||x==null))return;
  const [month,version,generation,count,hash,time]=r;
  if(!/^\d{6}月$/.test(month)||month<'202610月'||version!==VERSION||typeof generation!=='string'||!/^[0-9a-f-]{36}$/.test(generation)||!Number.isSafeInteger(count)||count<1||count>Math.ceil(MAX_PAYLOAD/CHUNK)||r.length<6+count||r.slice(6+count).some(x=>x!==''&&x!=null)||out.has(month))throw Error('V2 快照記錄不完整或重複，停止對帳');
  const chunks=r.slice(6,6+count);if(chunks.some(c=>typeof c!=='string'||!c.length||c.length>CHUNK))throw Error('V2 快照分片無效');
  const raw=chunks.join('');if(raw.length>MAX_PAYLOAD||digest(raw)!==hash)throw Error('V2 快照校驗失敗，禁止退回舊證據');
  let parsed;try{parsed=JSON.parse(raw);}catch{throw Error('V2 快照 JSON 損壞');}
  out.set(month,{...validate(parsed),storage:'v2',v2Row:i+2,generation,updatedAt:time,storedCells:r.length});
 });return out;
}
async function save({sheets,book,month,gridRows,ids,positions,identities,ensure}){
 if(!/^\d{6}月$/.test(month)||month<'202610月')throw Error('禁止遷移歷史月份快照');
 const payload=validate({gridRows,ids:[...ids].sort(),positions:Object.fromEntries(positions),identities});
 const raw=JSON.stringify(payload);if(raw.length>MAX_PAYLOAD)throw Error('V2 快照超出明確容量，保留待辦並停止，未截斷');
 const chunks=[];for(let i=0;i<raw.length;i+=CHUNK)chunks.push(raw.slice(i,i+CHUNK));
 await ensure(book,TAB,HEADERS);
 // Re-read rather than trust an in-memory row number after a lost response.
 const previous=(await read(sheets,book)).get(month),tab=(await topology(sheets,book)).find(t=>t.properties.title===TAB).properties;
 const generation=randomUUID(),values=[month,VERSION,generation,chunks.length,digest(raw),new Date().toISOString(),...chunks];
 const cells=values.map(v=>({userEnteredValue:typeof v==='number'?{numberValue:v}:{stringValue:v}}));
 if(previous)while(cells.length<previous.storedCells)cells.push({});
 const requests=[];if((tab.gridProperties?.columnCount||0)<cells.length)requests.push({updateSheetProperties:{properties:{sheetId:tab.sheetId,gridProperties:{columnCount:cells.length}},fields:'gridProperties.columnCount'}});
 if(previous)requests.push({updateCells:{range:{sheetId:tab.sheetId,startRowIndex:previous.v2Row-1,endRowIndex:previous.v2Row,startColumnIndex:0,endColumnIndex:cells.length},rows:[{values:cells}],fields:'userEnteredValue'}});
 else requests.push({appendCells:{sheetId:tab.sheetId,rows:[{values:cells}],fields:'userEnteredValue'}});
 await sheets.spreadsheets.batchUpdate({spreadsheetId:book,requestBody:{requests}});
 return {...payload,storage:'v2',v2Row:previous?.v2Row,generation,storedCells:cells.length};
}
module.exports={TAB,HEADERS,CHUNK,MAX_PAYLOAD,validate,read,save};
