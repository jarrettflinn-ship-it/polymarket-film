'use strict';
const https = require('https');

// ── HTTP helper ────────────────────────────────────────────────────────────
function kalshiGet(kPath) {
  return new Promise((resolve) => {
    const opts = {
      hostname: 'external-api.kalshi.com',
      path:     kPath,
      headers:  { Accept: 'application/json', 'User-Agent': 'PolyFilm/1.0' },
    };
    const req = https.get(opts, (res) => {
      let raw = '';
      res.on('data',  c  => (raw += c));
      res.on('end',   () => { try { resolve(JSON.parse(raw)); } catch { resolve({}); } });
    });
    req.on('error', () => resolve({}));
    req.setTimeout(15000, () => { req.destroy(); resolve({}); });
  });
}

// ── Parser ─────────────────────────────────────────────────────────────────
function parseKalshiEvent(ev) {
  const markets = (ev.markets || []).map(m => {
    const yesBid   = parseFloat(m.yes_bid_dollars)   || 0;
    const yesAsk   = parseFloat(m.yes_ask_dollars)   || 0;
    const last     = parseFloat(m.last_price_dollars) || 0;
    const yesPrice = yesBid > 0 && yesAsk > 0 ? (yesBid + yesAsk) / 2 : (last || 0.5);
    return {
      ticker:         m.ticker,
      question:       m.title || m.ticker,
      yesPrice:       +yesPrice.toFixed(4),
      noPrice:        +(1 - yesPrice).toFixed(4),
      volume:         parseFloat(m.volume_fp)        || 0,
      openInterest:   parseFloat(m.open_interest_fp) || 0,
      lastTradePrice: last,
      yesBid, yesAsk,
      closed: m.status === 'finalized' || m.status === 'settled',
    };
  }).filter(m => !m.closed);

  const mk0 = (ev.markets || [])[0] || {};
  return {
    id:       ev.event_ticker,
    ticker:   ev.event_ticker,
    title:    ev.title || ev.event_ticker,
    category: ev.category || '',
    volume:   markets.reduce((s, m) => s + m.volume, 0),
    endDate:  mk0.expected_expiration_time || mk0.close_time || null,
    source:   'kalshi',
    markets:  markets.sort((a, b) => b.yesPrice - a.yesPrice),
  };
}

module.exports = { kalshiGet, parseKalshiEvent };
