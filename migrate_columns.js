/**
 * TESTING monthly ledger one-time migration:
 * A:O old layout -> rename A as 帳務歸屬, insert B as 付款場別,
 * old B:O become C:P; old P/Q become Q/R (read-only historical reference).
 * Defaults DRY RUN. Only operates on explicitly whitelisted TESTING sheet ID.
 * Run: node migrate_columns.js
 */
const LEGACY = [
  '場別','收支類型','付款來源','科目代號','科目分類','日期',
  '品項','用途說明','廠商名稱','數量','單價','金額',
  '發票或憑證','代墊對象','填表人'
];
const NEW = [
  '帳務歸屬','付款場別','收支類型','付款來源','科目代號','科目分類',
  '日期','品項','用途說明','廠商名稱','數量','單價','金額',
  '發票或憑證','代墊對象','填表人'
];

function planMigration(rows) {
  const header = rows[0] || [];
  const norm = x => String(x ?? '').trim();
  if (NEW.every((h,i) => norm(header[i]) === h)) return {status:'already',payers:[],count:0};
  if (!LEGACY.every((h,i) => norm(header[i]) === h)) {
    throw new Error('欄位與舊版 A:O 不符；已停止，請先核對實際標題');
  }
  for (const [index,expected] of [[15,'帳務歸屬'],[16,'付款場別']]) {
    if (norm(header[index]) && norm(header[index]) !== expected) {
      throw new Error(`原 ${index===15?'P':'Q'} 欄有非預期標題；不覆蓋`);
    }
  }
  const payers=[];
  let inherited=0, fromOldQ=0, changed=0, nonempty=0;
  for (let i=1; i<rows.length; i++) {
    const r=rows[i] || [];
    if (!r.some(x=>norm(x))) {payers.push('');continue;}
    nonempty++;
    const owner=norm(r[0]);
    if(!owner) throw new Error(`第 ${i+1} 列 A 欄沒有歸屬場別；已停止`);
    const formerPayer=norm(r[16]);
    const payer=formerPayer || owner;
    payers.push(payer);
    if (formerPayer) {fromOldQ++; if(formerPayer!==owner) changed++;}
    else inherited++;
  }
  return {status:'ready',payers,count:nonempty,inherited,fromOldQ,changed};
}

function migrationRequests(tabId,plan) {
  if (plan.status!=='ready') throw new Error('沒有待搬移欄位');
  const cell = x => ({userEnteredValue:{stringValue:x}});
  const requests=[
    {insertDimension:{range:{sheetId:tabId,dimension:'COLUMNS',startIndex:1,endIndex:2},inheritFromBefore:true}},
    {updateCells:{range:{sheetId:tabId,startRowIndex:0,endRowIndex:1,startColumnIndex:0,endColumnIndex:2},
      rows:[{values:[cell('帳務歸屬'),cell('付款場別')]}],fields:'userEnteredValue'}},
    // Old P/Q remain for inspection; distinguish historical columns from new A/B.
    {updateCells:{range:{sheetId:tabId,startRowIndex:0,endRowIndex:1,startColumnIndex:16,endColumnIndex:18},
      rows:[{values:[cell('舊帳務歸屬(核對用)'),cell('舊付款場別(核對用)')]}],fields:'userEnteredValue'}},
  ];
  if (plan.payers.length) {
    requests.push({updateCells:{
      range:{sheetId:tabId,startRowIndex:1,endRowIndex:plan.payers.length+1,startColumnIndex:1,endColumnIndex:2},
      rows:plan.payers.map(x=>({values:[cell(x)]})), fields:'userEnteredValue'
    }});
  }
  return requests;
}

async function main() {
  require('dotenv').config();
  const {google}=require('googleapis');
  const id=String(process.env.MIGRATE_SHEET_ID||'').trim();
  const month=String(process.env.MIGRATE_MONTH||'').trim();
  const apply=process.env.MIGRATE_APPLY==='YES';
  if(process.env.APP_ENV!=='TESTING') throw new Error('僅允許 APP_ENV=TESTING');
  if(!id || !/^\d{6}月$/.test(month)) throw new Error('請設定 MIGRATE_SHEET_ID 與 MIGRATE_MONTH，例如 202610月');
  const allowed=(process.env.TEST_ALLOWED_SHEET_IDS||'').split(',').map(x=>x.trim()).filter(Boolean);
  if(!allowed.includes(id)) throw new Error('帳本不在 TEST_ALLOWED_SHEET_IDS，已停止');
  for(const special of ['CENTRAL_SHEET_ID','ACCOUNT_MASTER_SHEET_ID','TEST_USER_MASTER_SHEET_ID']){
    if (process.env[special] && id===process.env[special]) throw new Error('目標不能是中央或主檔試算表');
  }
  if(apply && process.env.MIGRATE_CONFIRM_SHEET_ID!==id)
    throw new Error('正式執行前須設定 MIGRATE_CONFIRM_SHEET_ID 與 MIGRATE_SHEET_ID 完全相同');
  const auth=new google.auth.JWT({
    email:process.env.GOOGLE_CLIENT_EMAIL,
    key:(process.env.GOOGLE_PRIVATE_KEY||'').replace(/\\n/g,'\n'),
    scopes:['https://www.googleapis.com/auth/spreadsheets']
  });
  const sheets=google.sheets({version:'v4',auth});
  const meta=await sheets.spreadsheets.get({spreadsheetId:id,fields:'sheets.properties(sheetId,title)'});
  const tab=(meta.data.sheets||[]).find(s=>s.properties.title===month);
  if(!tab) throw new Error('找不到月份分頁 '+month);
  const response=await sheets.spreadsheets.values.get({
    spreadsheetId:id,range:`'${month}'!A1:Q`,valueRenderOption:'FORMATTED_VALUE'
  });
  const plan=planMigration(response.data.values||[]);
  if(plan.status==='already') { console.log('此月份 A/B 已完成搬移，不會重複插入');return;}
  console.log(`月份=${month}；有效資料=${plan.count}筆；使用舊Q付款場別=${plan.fromOldQ}筆；按A補付款場別=${plan.inherited}筆`);
  console.log(`舊Q與A不同=${plan.changed}筆（只保留來源紀錄，不另行判讀舊測試代付指令）`);
  console.log('搬移後：A帳務歸屬、B付款場別、C收支類型、G日期、M金額、P填表人；舊P/Q在Q/R供核對。');
  if(!apply) { console.log('DRY RUN：只檢查未修改；確認備份後才可設定 MIGRATE_APPLY=YES。'); return; }
  // One atomic Google Sheets batch update: insert B + headers + payer values.
  await sheets.spreadsheets.batchUpdate({spreadsheetId:id,requestBody:{requests:migrationRequests(tab.properties.sheetId,plan)}});
  const check=await sheets.spreadsheets.values.get({
    spreadsheetId:id,range:`'${month}'!A1:R2`,valueRenderOption:'FORMATTED_VALUE'
  });
  const first=check.data.values?.[0]||[];
  if(!NEW.every((h,i)=>String(first[i]||'').trim()===h)) throw new Error('已執行搬移但回讀檢查失敗，請停止寫入並人工檢查');
  console.log('搬移成功：已回讀檢查 A:P 標題。接下來核對 C/G/M/P 資料，再部署 TESTING server.js。');
}

if(require.main===module) main().catch(e=>{console.error('停止：',e.message);process.exitCode=1;});
module.exports={LEGACY,NEW,planMigration,migrationRequests};
