'use strict';
const {AsyncLocalStorage}=require('node:async_hooks');
const {randomUUID}=require('node:crypto');
const lockContext=new AsyncLocalStorage(),activeLocks=new Set();
const LEGACY_TEMP='LINE accounting temporary mutation lock';
function isTemporary(p,actor){
 if(p.description===LEGACY_TEMP)return true;
 try{const d=JSON.parse(p.description);return d.app==='line-accounting'&&d.kind==='temporary-mutation'&&d.v===2&&d.actor===actor&&typeof d.token==='string'&&/^[0-9a-f-]{36}$/.test(d.token);}catch{return false;}
}
const KEY='line-accounting.tx.v1';
const COMPACT_KEY='l',COMPACT_VALUE='2';
const registryStore=require('./row-registry');
const MIN_MONTH='202610月';
function assertMonth(month) {
 if (!/^\d{6}月$/.test(month)||month<MIN_MONTH) throw Error('歷史或無效月份禁止寫入');
}
function identityValue(txId,sourceSheetId,monthSheet) {
 assertMonth(monthSheet);
 return JSON.stringify({txId,sourceSheetId,monthSheet});
}
async function sheetInfo(sheets,book) {
 const response=await sheets.spreadsheets.get({spreadsheetId:book,fields:'sheets(properties(sheetId,title,gridProperties),protectedRanges)'});
 return response.data;
}
async function metadata(sheets,book) {
 // Search returns row-scoped metadata explicitly; do not rely on the spreadsheet's
 // top-level developerMetadata field (which may omit metadata on child locations).
 const response=await sheets.spreadsheets.developerMetadata.search({spreadsheetId:book,
  requestBody:{dataFilters:[KEY,COMPACT_KEY].map(metadataKey=>({developerMetadataLookup:{metadataKey,visibility:'DOCUMENT'}}))}});
 return {developerMetadata:(response.data.matchedDeveloperMetadata||[]).map(m=>m.developerMetadata)};
}
function matches(data,value,sheetId) {
 return (data.developerMetadata||[]).filter(m=>m.metadataKey===KEY && m.metadataValue===value &&
  m.location?.dimensionRange?.sheetId===sheetId && m.location.dimensionRange.dimension==='ROWS' &&
  m.location.dimensionRange.endIndex===m.location.dimensionRange.startIndex+1);
}
function fullSheetProtection(p,sheetId) {
 return p?.range?.sheetId===sheetId && (p.range.startRowIndex==null||p.range.startRowIndex===0) && p.range.endRowIndex==null &&
  (p.range.startColumnIndex==null||p.range.startColumnIndex===0) && p.range.endColumnIndex==null;
}
function validProtection(p,sheetId,actor,owner) {
 const users=p?.editors?.users||[];
 const normalize=email=>String(email).trim().toLowerCase();
 const allowed=new Set([actor,owner].filter(Boolean).map(normalize));
 return fullSheetProtection(p,sheetId) && Number.isInteger(p.protectedRangeId) && !p.warningOnly && !p.unprotectedRanges?.length &&
  !p.editors?.groups?.length && !p.editors?.domainUsersCanEdit && users.length>0 && users.length<=allowed.size &&
  new Set(users.map(normalize)).size===users.length && users.map(normalize).includes(normalize(actor)) && users.every(u=>allowed.has(normalize(u)));
}
// Ordinary editors cannot sort/delete the protected sheet while the final ID check and
// mutation run. Spreadsheet owners/admins must not override this protection.
async function releaseTemporary(sheets,book,tab,p,actor,owner){
 const latest=await sheetInfo(sheets,book);
 const current=latest.sheets?.find(s=>s.properties.sheetId===tab.sheetId)?.protectedRanges?.find(q=>q.protectedRangeId===p.protectedRangeId);
 if(!current)return;
 if(current.description!==p.description||!isTemporary(current,actor)||!validProtection(current,tab.sheetId,actor,owner))throw Error('臨時保護已改變，禁止清除');
 await sheets.spreadsheets.batchUpdate({spreadsheetId:book,requestBody:{requests:[{deleteProtectedRange:{protectedRangeId:p.protectedRangeId}}]}});
}
async function protect(sheets,book,title,actor,work,owners=new Map()) {
 if(!actor)throw Error('缺少服務帳號身分，禁止結構寫入');
 const data=await sheetInfo(sheets,book),sheet=data.sheets?.find(t=>t.properties.title===title),tab=sheet?.properties;
 if(!tab)throw Error('找不到待保護分頁');
 const full=(sheet.protectedRanges||[]).filter(p=>fullSheetProtection(p,tab.sheetId)),key=JSON.stringify([book,tab.sheetId]);
 if(full.some(p=>!validProtection(p,tab.sheetId,actor,owners.get(book))))throw Error('既有整頁保護無法確認，停止寫入');
 if(activeLocks.has(key)){
  if(!lockContext.getStore()?.has(key))throw Error('分頁由其他工作使用');
  return work(tab);
 }
 let lock=full.find(p=>isTemporary(p,actor));
 if(full.length&&!lock)return work(tab); // Existing permanent/user protection is never removed.
 if(lock&&full.length!==1)throw Error('臨時保護不唯一，停止寫入');
 if(!lock){
  const description=JSON.stringify({app:'line-accounting',kind:'temporary-mutation',v:2,actor,token:randomUUID()});
  const response=await sheets.spreadsheets.batchUpdate({spreadsheetId:book,requestBody:{requests:[{addProtectedRange:{protectedRange:{range:{sheetId:tab.sheetId},description,warningOnly:false,editors:{users:[actor],groups:[],domainUsersCanEdit:false}}}}]}});
  lock=response.data.replies?.[0]?.addProtectedRange?.protectedRange;
  if(!Number.isInteger(lock?.protectedRangeId)||!validProtection(lock,tab.sheetId,actor,owners.get(book)))throw Error('分頁保護未確認，禁止寫入');
 }
 // Adopt a verified orphan without unlocking first. Nested work retains the lock.
 activeLocks.add(key);const context=new Set(lockContext.getStore()||[]);context.add(key);
 try{return await lockContext.run(context,()=>work(tab));}
 finally{try{await releaseTemporary(sheets,book,tab,lock,actor,owners.get(book));}finally{activeLocks.delete(key);}}
}
async function recoverTemporaryProtections(sheets,book,actor,owners=new Map()){
 if(!actor)throw Error('缺少服務帳號身分，禁止恢復保護');
 const data=await sheetInfo(sheets,book);let recovered=0;
 for(const sheet of data.sheets||[]){
  const tab=sheet.properties;
  if(!/^\d{6}月$/.test(tab.title)||tab.title<MIN_MONTH)continue;
  const full=(sheet.protectedRanges||[]).filter(p=>fullSheetProtection(p,tab.sheetId));
  const temporary=full.filter(p=>isTemporary(p,actor));
  if(!temporary.length||activeLocks.has(JSON.stringify([book,tab.sheetId])))continue;
  if(full.length!==1||!validProtection(temporary[0],tab.sheetId,actor,owners.get(book)))throw Error('遺留保護無法安全確認');
  await releaseTemporary(sheets,book,tab,temporary[0],actor,owners.get(book));recovered++;
 }
 return recovered; // Only protection IDs change; never cells or transaction metadata.
}
async function secureSystem(sheets,book,title,actor,owners=new Map()) {
 if(!actor)throw Error('缺少服務帳號身分，禁止系統分頁寫入');
 const data=await sheetInfo(sheets,book);
 const sheet=data.sheets?.find(t=>t.properties.title===title);
 if(!sheet)throw Error('找不到系統分頁');
 const description='LINE accounting managed system sheet';
 const valid=(sheet.protectedRanges||[]).some(p=>p.description===description && validProtection(p,sheet.properties.sheetId,actor,owners.get(book)));
 if(valid)return;
 if((sheet.protectedRanges||[]).some(p=>fullSheetProtection(p,sheet.properties.sheetId)))throw Error('系統分頁既有整頁保護不符，停止寫入');
 const result=await sheets.spreadsheets.batchUpdate({spreadsheetId:book,requestBody:{requests:[{addProtectedRange:{protectedRange:{
  range:{sheetId:sheet.properties.sheetId},description,warningOnly:false,editors:{users:[actor],groups:[],domainUsersCanEdit:false}
 }}}]}});
 const applied=result.data.replies?.[0]?.addProtectedRange?.protectedRange;
 if(!Number.isInteger(applied?.protectedRangeId)||!validProtection(applied,sheet.properties.sheetId,actor,owners.get(book)))throw Error('系統分頁保護未確認');
}

async function bind(sheets,book,tab,txId,sourceBook,month,row,cached,owners=new Map(),actor=process.env.GOOGLE_CLIENT_EMAIL) {
 const value=identityValue(txId,sourceBook,month),data=cached||await metadata(sheets,book);
 const secure=title=>secureSystem(sheets,book,title,actor,owners);
 if(!actor)throw Error('缺少身分登記表服務帳號');
 const registry=data.registry||(data.registry=await registryStore.load(sheets,book,secure)),key=registryStore.key(txId,sourceBook,month,tab.sheetId),known=registry.records.get(key);
 if(known!=null){
  if(data.developerMetadata.filter(m=>[KEY,COMPACT_KEY].includes(m.metadataKey)&&m.location?.dimensionRange?.sheetId===tab.sheetId&&m.location.dimensionRange.startIndex===row-1).length!==1)throw Error('列級交易身分重複或缺失');
  const hit=data.developerMetadata.filter(m=>m.metadataId===known);
  if(hit.length!==1||hit[0].metadataKey!==COMPACT_KEY||hit[0].metadataValue!==COMPACT_VALUE||hit[0].location?.dimensionRange?.sheetId!==tab.sheetId||hit[0].location.dimensionRange.dimension!=='ROWS'||hit[0].location.dimensionRange.startIndex!==row-1||hit[0].location.dimensionRange.endIndex!==row)throw Error('登記身分與來源列不符，停止寫入');
  return known;
 }
 const legacy=matches(data,value,tab.sheetId);if(legacy.length>1)throw Error('列級交易身分重複');
 if(legacy.length&&legacy[0].location.dimensionRange.startIndex!==row-1)throw Error('交易 ID 與列級身分位置不符');
 if(!legacy.length&&data.developerMetadata.some(m=>[KEY,COMPACT_KEY].includes(m.metadataKey)&&m.location?.dimensionRange?.sheetId===tab.sheetId&&m.location.dimensionRange.startIndex===row-1))throw Error('資料列已有其他或未登記交易身分');
 const id=legacy[0]?.metadataId||registryStore.freshId(data,registry),marker={metadataId:id,metadataKey:COMPACT_KEY,metadataValue:COMPACT_VALUE,visibility:'DOCUMENT',location:{dimensionRange:{sheetId:tab.sheetId,dimension:'ROWS',startIndex:row-1,endIndex:row}}};
 if(legacy.length&&registry.used.has(id))throw Error('舊列身分已被其他登記使用');
 const mutation=legacy.length?{updateDeveloperMetadata:{dataFilters:[{developerMetadataLookup:{metadataId:id}}],developerMetadata:{metadataKey:COMPACT_KEY,metadataValue:COMPACT_VALUE},fields:'metadataKey,metadataValue'}}:{createDeveloperMetadata:{developerMetadata:marker}};
 await sheets.spreadsheets.batchUpdate({spreadsheetId:book,requestBody:{requests:[mutation,registryStore.appendRequest(registry,key,id)]}});
 registry.records.set(key,id);registry.used.add(id);if(legacy.length)Object.assign(legacy[0],marker);else data.developerMetadata.push(marker);
 return id;
}
function rowForIdentity(data,value,tab,id) {
 const hits=(data.developerMetadata||[]).filter(m=>m.metadataId===id&&((m.metadataKey===KEY&&m.metadataValue===value)||(m.metadataKey===COMPACT_KEY&&m.metadataValue===COMPACT_VALUE))&&m.location?.dimensionRange?.sheetId===tab.sheetId&&m.location.dimensionRange.dimension==='ROWS'&&m.location.dimensionRange.endIndex===m.location.dimensionRange.startIndex+1);
 if(hits.length!==1)throw Error('無法唯一確認列級交易身分');return hits[0].location.dimensionRange.startIndex+1;
}
function deletionProven(data,tab,previous,txId,currentIds) {
 const proof=previous?.[txId];
 if(!proof || proof.sheetId!==tab.sheetId || !Number.isInteger(proof.metadataId)||currentIds.has(txId))return false;
 // A cleared ID leaves row metadata intact. Only disappearance of the known metadata
 // with the matching ID's absence is accepted. Never infer deletion from grid size.
 return !(data.developerMetadata||[]).some(m=>m.metadataId===proof.metadataId);
}
module.exports={COMPACT_KEY,COMPACT_VALUE,registryStore,recoverTemporaryProtections,KEY,assertMonth,identityValue,sheetInfo,metadata,matches,protect,secureSystem,bind,rowForIdentity,deletionProven};
// Explicit operator declaration, scoped per spreadsheet. This does not discover
// or certify ownership using Drive, and never trusts an arbitrary second editor.
function withDeclaredOwners(declarations={},declaredActor=process.env.GOOGLE_CLIENT_EMAIL){
 if(!declarations||typeof declarations!=='object'||Array.isArray(declarations))throw Error('帳本擁有者設定無效');
 const owners=new Map();
 for(const [book,email] of Object.entries(declarations)){
  if(!book||typeof email!=='string'||!/^\S+@\S+\.\S+$/.test(email.trim()))throw Error('帳本擁有者設定無效');
  owners.set(book,email.trim().toLowerCase());
 }
 return {...module.exports,bind:(sheets,book,tab,txId,sourceBook,month,row,cached)=>bind(sheets,book,tab,txId,sourceBook,month,row,cached,owners,declaredActor),protect:(sheets,book,title,actor,work)=>protect(sheets,book,title,actor,work,owners),
  secureSystem:(sheets,book,title,actor)=>secureSystem(sheets,book,title,actor,owners),
  recoverTemporaryProtections:(sheets,book,actor)=>recoverTemporaryProtections(sheets,book,actor,owners)};
}
module.exports.withDeclaredOwners=withDeclaredOwners;
