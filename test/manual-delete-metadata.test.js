'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const source=fs.readFileSync(path.join(__dirname,'../server.js'),'utf8');
const start=source.indexOf('const MANUAL_DELETE_METADATA_KEY =');
const end=source.indexOf('// V2.9 離線候選',start);
assert.ok(start>=0&&end>start);
const context={};vm.runInNewContext(source.slice(start,end)+'\nthis.track=manualDeleteMarkers;',context);
const track=context.track;
function mock(existing=[]){
 const calls=[];
 const sheets={spreadsheets:{
  developerMetadata:{search:async()=>({data:{matchedDeveloperMetadata:existing.map(developerMetadata=>({developerMetadata}))}})},
  batchUpdate:async({requestBody})=>{
   calls.push(requestBody.requests);
   return {data:{replies:requestBody.requests.map((_,i)=>({createDeveloperMetadata:{developerMetadata:{metadataId:101+i}}}))}};
  }
 }};
 return {sheets,calls};
}
const marker=(id,metadataId,row=2)=>({metadataId,metadataKey:'line-accounting.manual-delete.v1',metadataValue:id,location:{dimensionRange:{sheetId:7,dimension:'ROWS',startIndex:row-1,endIndex:row}}});
test('first baseline binds metadata in one batch, not proof of deletion',async()=>{
 const m=mock(),r=await track(m.sheets,'book',7,new Map([['a',[2]],['b',[3]]]),null);
 assert.equal(r.markers.a,101);assert.equal(r.markers.b,102);assert.equal(r.missing.size,0);
 assert.equal(m.calls.length,1);assert.equal(m.calls[0].length,2);
});
test('deleted row metadata disappears and is proven missing',async()=>{
 const m=mock([marker('a',21)]),r=await track(m.sheets,'book',7,new Map([['a',[2]]]),{markers:{a:21,b:22}});
 assert.equal(r.missing.has('b'),true);assert.equal(r.missing.has('a'),false);
 assert.equal(m.calls.length,0);
});
test('clearing ID does not remove row metadata',async()=>{
 const m=mock([marker('a',21),marker('b',22,3)]),r=await track(m.sheets,'book',7,new Map([['a',[2]]]),{markers:{a:21,b:22}});
 assert.equal(r.missing.has('b'),false);
});
test('moved marker is not evidence of deletion',async()=>{
 const m=mock([marker('a',21),marker('b',22,8)]),r=await track(m.sheets,'book',7,new Map([['a',[2]]]),{markers:{a:21,b:22}});
 assert.equal(r.missing.has('b'),false);
});
test('duplicate ID fails closed',async()=>{
 const m=mock();await assert.rejects(track(m.sheets,'book',7,new Map([['a',[2,3]]]),null),/交易ID重複/);
});
