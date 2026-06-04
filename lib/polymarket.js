'use strict';
const https = require('https');

// ── HTTP helper ────────────────────────────────────────────────────────────
function get(path) {
  return new Promise((resolve) => {
    const opts = {
      hostname: 'gamma-api.polymarket.com',
      path,
      headers: { Accept: 'application/json', 'User-Agent': 'PolyFilm/1.0' },
    };
    const req = https.get(opts, (res) => {
      let raw = '';
      res.on('data', (c) => (raw += c));
      res.on('end', () => { try { resolve(JSON.parse(raw)); } catch { resolve([]); } });
    });
    req.on('error', () => resolve([]));
    req.setTimeout(10000, () => { req.destroy(); resolve([]); });
  });
}

// ── Tag config ─────────────────────────────────────────────────────────────
const LIVE_TAGS = ['movies', 'box-office', 'pop-culture', 'netflix', 'top-netflix'];

// ── Keyword gates ──────────────────────────────────────────────────────────
const FILM_RE = /\b(movie|film|box office|oscar|emmy|grammy|golden globe|actor|actress|director|marvel|pixar|disney|hollywood|cinema|netflix|hbo|amazon prime|prime video|prime original|disney plus|apple tv|apple original|peacock original|paramount plus|hulu original|fx original|amc original|showtime|starz|streaming|gross|opening weekend|james bond|sequel|remake|blockbuster|best picture|animated|superhero|studio|rotten tomatoes|imdb|box.?office|season \d|series \d|episode \d|season finale|series finale|showrunner|season premiere|series premiere|midseason)\b/i;
const NON_FILM_RE = /\b(poker|chess|nba|nfl|mlb|nhl|golf|tennis|soccer|football|basketball|baseball|election|president|senate|congress|cryptocurrency|crypto|bitcoin|formula.?1|f1\b|ufc|mma|wrestling)\b/i;

const FILM_TAGS = new Set(['movies', 'box-office', 'netflix', 'top-netflix']);
const SHOW_TAGS = new Set(['euphoria', 'the-boys', 'stranger-things', 'squid-game', 'wednesday', 'bridgerton']);

// ── Parsers ────────────────────────────────────────────────────────────────
function parseStrArr(val, fallback) {
  if (Array.isArray(val))      return val;
  if (typeof val === 'string') { try { return JSON.parse(val); } catch {} }
  return fallback;
}

function parseEvent(ev, isHistorical = false) {
  const markets = (ev.markets || [])
    .map((m) => {
      const prices = parseStrArr(m.outcomePrices, ['0.5', '0.5']);
      return {
        id:                m.id,
        question:          m.question || '',
        yesPrice:          parseFloat(prices[0])           || 0,
        noPrice:           parseFloat(prices[1])           || 0,
        volume:            parseFloat(m.volume)            || 0,
        closed:            !!m.closed,
        outcomes:          parseStrArr(m.outcomes, ['Yes', 'No']),
        oneDayPriceChange: parseFloat(m.oneDayPriceChange) || 0,
        spread:            parseFloat(m.spread)            || 0,
        lastTradePrice:    parseFloat(m.lastTradePrice)    || 0,
        volume1wk:         parseFloat(m.volume1wk)         || 0,
        volume1mo:         parseFloat(m.volume1mo)         || 0,
      };
    })
    .filter((m) => m.volume > 0 || !m.closed);

  return {
    id:           ev.id,
    title:        ev.title     || 'Unknown',
    slug:         ev.slug      || '',
    volume:       parseFloat(ev.volume)       || 0,
    volume24hr:   parseFloat(ev.volume24hr)   || 0,
    volume1wk:    parseFloat(ev.volume1wk)    || 0,
    openInterest: parseFloat(ev.openInterest) || 0,
    liquidity:    parseFloat(ev.liquidity)    || 0,
    competitive:  parseFloat(ev.competitive)  || 0,
    commentCount: parseInt(ev.commentCount, 10) || 0,
    endDate:      ev.endDate   || null,
    startDate:    ev.startDate || null,
    closed:       !!ev.closed,
    archived:     !!ev.archived,
    restricted:   !!ev.restricted,
    historical:   isHistorical,
    tags:         (ev.tags || []).map((t) => t.slug || t.label || '').filter(Boolean),
    markets:      markets.sort((a, b) => b.yesPrice - a.yesPrice),
    image:        ev.image || null,
  };
}

// ── Dedup + filter ─────────────────────────────────────────────────────────
function isFilmEvent(ev) {
  const title = ev.title || '';
  if (NON_FILM_RE.test(title)) return false;
  const tags = (ev.tags || []).map(t => t.slug || '');
  if (tags.some(s => FILM_TAGS.has(s))) return true;
  if (tags.some(s => SHOW_TAGS.has(s))) return true;
  return FILM_RE.test(title);
}

function addBatch(batch, seen, events, { historical = false } = {}) {
  if (!Array.isArray(batch)) return;
  for (const ev of batch) {
    if (seen.has(ev.id)) continue;
    if (!isFilmEvent(ev))  continue;
    seen.add(ev.id);
    events.push(parseEvent(ev, historical));
  }
}

// ── Fetchers ───────────────────────────────────────────────────────────────
function fetchLive(tag) {
  return get(`/events?tag_slug=${tag}&closed=false&limit=200&order=volume&ascending=false`);
}

// Single-page closed fetch — fast enough for serverless (no pagination needed)
function fetchRecentClosed(tag, limit = 100) {
  return get(`/events?tag_slug=${tag}&closed=true&limit=${limit}&order=end_date_iso&ascending=false`);
}

module.exports = { get, LIVE_TAGS, fetchLive, fetchRecentClosed, parseEvent, isFilmEvent, addBatch };
