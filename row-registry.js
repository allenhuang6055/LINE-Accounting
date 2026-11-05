'use strict';
// Full identities live in protected cells, not Google's limited metadata text.
// Explicit metadataId + registry append commit in ONE spreadsheet batch.
const {randomInt}=require('node:crypto');
const TAB='系統交易身分',HEADERS=['交易身分JSON','metadataId'];
function key(txId,sourceSheetId,monthSheet,sheetId){return JSON.stringify({txId,sourceSheetId,monthSheet,sheetId});}
async function load(sheets,book,secure){
 let tabs=(await sheets.spreadsheets.get({spreadsheetId:book,fields:'sheets.properties'})).data.sheets||[],tab=tabs.find(t=>t.properties.title===TAB)?.properties;
 if(!tab){await sheets.spreadsheets.batchUpdate({spreadsheetId:book,requestBody:{requests:[{addSheet:{properties:{title:TAB}}}]}});tabs=(await sheets.spreadsheets.get({spreadsheetId:book,fields:'sheets.properties'})).data.sheets;tab=tabs.find(t=>t.properties.title===TAB).properties;
 }
 await secure(TAB);
 const header=(await sheets.spreadsheets.values.get({spreadsheetId:book,range:`'${TAB}'!A1:B1`,valueRenderOption:'UNFORMATTED_VALUE'})).data.values?.[0]||[];
 if(header.join('|')!==HEADERS.join('|')){
  if(header.some(v=>v!==''&&v!=null))throw Error('身分登記表表頭不符，停止寫入');
  await sheets.spreadsheets.values.update({spreadsheetId:book,range:`'${TAB}'!A1:B1`,valueInputOption:'RAW',requestBody:{values:[HEADERS]}});
 }
 const rows=(await sheets.spreadsheets.values.get({spreadsheetId:book,range:`'${TAB}'!A2:B`,valueRenderOption:'UNFORMATTED_VALUE'})).data.values||[],records=new Map(),used=new Set();
 for(const row of rows){if(!row.length||!row[0]){if(row[1]!=null&&row[1]!=='')throw Error('身分登記表缺少交易');continue;}let p;try{p=JSON.parse(row[0]);}catch{throw Error('身分登記表 JSON 損壞');}
 if(!p||!/^TX-[0-9a-f-]{36}$/i.test(p.txId)||typeof p.sourceSheetId!=='string'||!p.sourceSheetId||!/^\d{6}月$/.test(p.monthSheet)||p.monthSheet<'202610月'||!Number.isInteger(p.sheetId)||p.sheetId<0||!Number.isInteger(row[1])||row[1]<1||records.has(row[0])||used.has(row[1]))throw Error('身分登記重複或無效');
 if(key(p.txId,p.sourceSheetId,p.monthSheet,p.sheetId)!==row[0])throw Error('身分登記鍵不標準');records.set(row[0],row[1]);used.add(row[1]);}
 return {tab,records,used};
}
function freshId(data,registry){const used=new Set([...registry.used,...data.developerMetadata.map(m=>m.metadataId)]);let id;do{id=randomInt(1,2147483647);}while(used.has(id));return id;}
function appendRequest(registry,identity,id){return {appendCells:{sheetId:registry.tab.sheetId,rows:[{values:[{userEnteredValue:{stringValue:identity}},{userEnteredValue:{numberValue:id}}]}],fields:'userEnteredValue'}};}
module.exports={TAB,HEADERS,key,load,freshId,appendRequest};
