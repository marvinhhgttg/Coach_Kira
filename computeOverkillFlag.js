function computeOverkillFlag(acwr, kei) {
  const acwrNum = Number(acwr);
  const keiNum = Number(kei);
  const acwrOver = Number.isFinite(acwrNum) && acwrNum > 1.3;
  const keiOver = Number.isFinite(keiNum) && keiNum > 30;
  return acwrOver || keiOver;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { computeOverkillFlag };
}
