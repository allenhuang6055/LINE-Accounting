'use strict';
const {AsyncLocalStorage}=require('node:async_hooks');
// Serializes webhook and background writes in one process. It is not a
// cross-service lock; this version requires a single writer deployment.
let sharedQueue;
function createWriteQueue(){
 if(sharedQueue)return sharedQueue;
 const context=new AsyncLocalStorage();let tail=Promise.resolve();
 return sharedQueue={exclusive(work){
  if(context.getStore())return work();
  const task=tail.then(()=>context.run(true,work));tail=task.catch(()=>{});return task;
 }};
}
module.exports={createWriteQueue};
