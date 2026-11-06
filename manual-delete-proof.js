'use strict';
// A mixed batch must never delete only the subset with proof.
// Verify membership, not merely matching counts.
function allMissingProven(missing, verified) {
  if (!Array.isArray(missing) || !Array.isArray(verified) || missing.length === 0) return false;
  const ids = missing.map(e => e?.id);
  const proven = verified.map(e => e?.id);
  return ids.every(id => typeof id === 'string' && id.length > 0) &&
    new Set(ids).size === ids.length &&
    new Set(proven).size === proven.length &&
    ids.length === proven.length &&
    proven.every(id => ids.includes(id));
}
module.exports = {allMissingProven};
