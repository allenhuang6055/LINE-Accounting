'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { deletionProven } = require('../sheet-identity');

const tab = {sheetId: 7};
const prior = {T1:{sheetId:7,metadataId:101},T2:{sheetId:7,metadataId:102}};
const rowMarker = (id, sheetId=7) => ({
  metadataId:id,metadataKey:'l',metadataValue:'2',
  location:{dimensionRange:{sheetId,dimension:'ROWS',startIndex:2,endIndex:3}}
});
test('deleted last row with vanished metadata is eligible',()=>{
  assert.equal(deletionProven({developerMetadata:[rowMarker(102)]},tab,prior,'T1',new Set(['T2'])),true);
});
test('clearing transaction ID without deleting row is not eligible',()=>{
  assert.equal(deletionProven({developerMetadata:[rowMarker(101),rowMarker(102)]},tab,prior,'T1',new Set(['T2'])),false);
});
test('moving a row or moving metadata to another tab is not deletion',()=>{
  assert.equal(deletionProven({developerMetadata:[rowMarker(101,8)]},tab,prior,'T1',new Set()),false);
});
test('still-present ID is not deletion even when metadata is missing',()=>{
  assert.equal(deletionProven({developerMetadata:[]},tab,prior,'T1',new Set(['T1'])),false);
});
test('missing prior identity cannot authorize deletion',()=>{
  assert.equal(deletionProven({developerMetadata:[]},tab,{},'T1',new Set()),false);
});
test('wrong sheet identity cannot authorize deletion',()=>{
  assert.equal(deletionProven({developerMetadata:[]},{sheetId:8},prior,'T1',new Set()),false);
});
