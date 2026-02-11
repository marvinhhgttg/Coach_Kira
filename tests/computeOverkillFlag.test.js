const test = require('node:test');
const assert = require('node:assert/strict');
const { computeOverkillFlag } = require('../computeOverkillFlag');

test('computeOverkillFlag boundary checks', () => {
  assert.equal(computeOverkillFlag(1.31, 30.0), true);
  assert.equal(computeOverkillFlag(1.3, 30.01), true);
  assert.equal(computeOverkillFlag(1.3, 30.0), false);
  assert.equal(computeOverkillFlag(null, 30.0), false);
  assert.equal(computeOverkillFlag(NaN, NaN), false);
});
