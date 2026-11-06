'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
const {withDeclaredOwners}=require('../sheet-identity');
const actor='bot@example.com';
function mockSheets({failAdd=false,foreignProtection=false}={}){
  const protections=[];
  if(foreignProtection) protections.push({protectedRangeId:9,range:{sheetId:7},warningOnly:false,editors:{users:['stranger@example.com']}});
  const calls=[];
  const sheets={spreadsheets:{
    get:async()=>({data:{sheets:[{properties:{sheetId:7,title:'202611月'},protectedRanges:[...protections]}]}}),
    batchUpdate:async({requestBody})=>{
      const request=requestBody.requests[0];
      if(request.addProtectedRange){
        calls.push('add');
        if(failAdd)throw Error('insufficient permission');
        const protectedRange={...request.addProtectedRange.protectedRange,protectedRangeId:11};
        protections.push(protectedRange);
        return {data:{replies:[{addProtectedRange:{protectedRange}}]}};
      }
      if(request.deleteProtectedRange){
        calls.push('delete');
        const i=protections.findIndex(p=>p.protectedRangeId===request.deleteProtectedRange.protectedRangeId);
        if(i<0)throw Error('unknown protection');
        protections.splice(i,1);
        return {data:{replies:[{}]}};
      }
      throw Error('unexpected batch request');
    }
  }};
  return {sheets,calls,protections};
}
test('employee-owned sheet works without owner declaration if bot may protect',async()=>{
  const m=mockSheets();
  const identity=withDeclaredOwners({},actor);
  const result=await identity.protect(m.sheets,'employee-owned-book','202611月',actor,async tab=>tab.sheetId);
  assert.equal(result,7);
  assert.deepEqual(m.calls,['add','delete']);
  assert.equal(m.protections.length,0);
});
test('permission denied stops safely before touching transaction rows',async()=>{
  const m=mockSheets({failAdd:true});
  const identity=withDeclaredOwners({},actor);
  let ran=false;
  await assert.rejects(identity.protect(m.sheets,'employee-owned-book','202611月',actor,async()=>{ran=true}),/insufficient permission/);
  assert.equal(ran,false);
});
test('unknown protection editor is not treated as owner',async()=>{
  const m=mockSheets({foreignProtection:true});
  const identity=withDeclaredOwners({},actor);
  await assert.rejects(identity.protect(m.sheets,'employee-owned-book','202611月',actor,async()=>{}),/既有整頁保護無法確認/);
  assert.deepEqual(m.calls,[]);
});
test('owner declaration cannot bypass unrelated protection editor',async()=>{
  const m=mockSheets({foreignProtection:true});
  const identity=withDeclaredOwners({'employee-owned-book':'owner@example.com'},actor);
  await assert.rejects(identity.protect(m.sheets,'employee-owned-book','202611月',actor,async()=>{}),/既有整頁保護無法確認/);
});
