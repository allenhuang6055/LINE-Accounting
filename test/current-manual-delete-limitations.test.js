'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
// Extract the production function without starting LINE, Google, or Render.
const fs=require('node:fs');
const vm=require('node:vm');
const source=fs.readFileSync(require('node:path').join(__dirname,'../server.js'),'utf8');
const begin=source.indexOf('function v29HasRowDeletionProof(');
const end=source.indexOf('\n// 中央安全對帳',begin);
assert.ok(begin>=0&&end>begin,'production deletion proof must exist');
const context={};vm.runInNewContext(source.slice(begin,end)+'\nthis.proof=v29HasRowDeletionProof;',context);
const proof=context.proof;
const base={snapshot:{ids:['A','B','C'],positions:{A:2,B:3,C:4}},byId:new Map([['A',[2]],['B',[3]]]),currentIds:new Set(['A','B']),newlyMissing:[{id:'C'}]};
test('last-row deletion is recognized only when grid row count decreases',()=>assert.equal(proof({...base,gridRows:99,snapshot:{...base.snapshot,gridRows:100}}),true));
test('last-row cleared ID without grid shrink stays blocked',()=>assert.equal(proof({...base,gridRows:100,snapshot:{...base.snapshot,gridRows:100}}),false));
test('unrelated grid shrink with missing ID stays blocked',()=>assert.equal(proof({...base,gridRows:98,snapshot:{...base.snapshot,gridRows:100}}),false));
test('current production recognizes shifted middle row deletion',()=>assert.equal(proof({...base,byId:new Map([['A',[2]],['C',[3]]]),currentIds:new Set(['A','C']),newlyMissing:[{id:'B'}]}),true));
test('cleared S cell must not be mistaken for deletion',()=>assert.equal(proof(base),false));
