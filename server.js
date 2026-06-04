'use strict';
const express = require('express');
const path    = require('path');
const https   = require('https');

const app  = express();
const PORT = 3001;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Tag config ─────────────────────────────────────────────────────────────
const LIVE_TAGS = ['movies', 'box-office', 'pop-culture', 'netflix', 'top-netflix'];
const HIST_TAGS = ['movies', 'box-office', 'pop-culture', 'netflix', 'top-netflix'];

// Keyword gates for pop-culture (too noisy without filtering)
// Covers: film awards, streaming platforms, box-office terms, TV series signals
const FILM_RE = /\b(movie|film|box office|oscar|emmy|grammy|golden globe|actor|actress|director|marvel|pixar|disney|hollywood|cinema|netflix|hbo|amazon prime|prime video|prime original|disney plus|apple tv|apple original|peacock original|paramount plus|hulu original|fx original|amc original|showtime|starz|streaming|gross|opening weekend|james bond|sequel|remake|blockbuster|best picture|animated|superhero|studio|rotten tomatoes|imdb|box.?office|season \d|series \d|episode \d|season finale|series finale|showrunner|season premiere|series premiere|midseason)\b/i;
const NON_FILM_RE = /\b(poker|chess|nba|nfl|mlb|nhl|golf|tennis|soccer|football|basketball|baseball|election|president|senate|congress|cryptocurrency|crypto|bitcoin|formula.?1|f1\b|ufc|mma|wrestling)\b/i;

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

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ── Fetchers ───────────────────────────────────────────────────────────────

// Live markets: closed=false avoids US geo-filter that active=true triggers
function fetchLive(tag) {
  return get(`/events?tag_slug=${tag}&closed=false&limit=200&order=volume&ascending=false`);
}

// Paginate through ALL closed (historical) markets for a tag
async function fetchAllClosed(tag) {
  const results = [];
  const LIMIT   = 100;
  const MAX     = 2000; // safety cap per tag

  for (let offset = 0; offset < MAX; offset += LIMIT) {
    const batch = await get(
      `/events?tag_slug=${tag}&closed=true&limit=${LIMIT}&offset=${offset}&order=end_date_iso&ascending=false`
    );
    if (!Array.isArray(batch) || batch.length === 0) break;
    results.push(...batch);
    if (batch.length < LIMIT) break; // last page
    await sleep(80);                 // polite delay between pages
  }

  return results;
}

// ── Parsers ────────────────────────────────────────────────────────────────

// Gamma API returns outcomePrices / outcomes as JSON strings in bulk queries
function parseStrArr(val, fallback) {
  if (Array.isArray(val))   return val;
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
    .filter((m) => m.volume > 0 || !m.closed); // keep new zero-volume markets if still open

  return {
    id:           ev.id,
    title:        ev.title     || 'Unknown',
    slug:         ev.slug      || '',
    volume:       parseFloat(ev.volume)      || 0,
    volume24hr:   parseFloat(ev.volume24hr)  || 0,
    volume1wk:    parseFloat(ev.volume1wk)   || 0,
    openInterest: parseFloat(ev.openInterest)|| 0,
    liquidity:    parseFloat(ev.liquidity)   || 0,
    competitive:  parseFloat(ev.competitive) || 0,
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

// ── Dedup + filter helper ──────────────────────────────────────────────────
// Tags whose events are always considered film/entertainment (no keyword gate needed)
const FILM_TAGS = new Set(['movies', 'box-office', 'netflix', 'top-netflix']);

// Show-specific tags Polymarket uses — automatically whitelisted
const SHOW_TAGS = new Set(['euphoria', 'the-boys', 'stranger-things', 'squid-game', 'wednesday', 'bridgerton']);

function isFilmEvent(ev) {
  const title = ev.title || '';
  const allText = title + ' ' + (ev.markets || []).map(m => m.question || '').join(' ');

  if (NON_FILM_RE.test(title)) return false;

  const tags = (ev.tags || []).map(t => t.slug || '');

  // Hard-whitelisted tags → always entertainment
  if (tags.some(s => FILM_TAGS.has(s))) return true;

  // Known show-specific tags → always entertainment
  if (tags.some(s => SHOW_TAGS.has(s))) return true;

  // For pop-culture / tv tags: match against the title only — checking sub-market questions
  // causes false positives when non-entertainment markets mention actor/film/gross colloquially.
  // The expanded FILM_RE (with season \d, platform names, etc.) covers show titles robustly.
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

// ── Caches ─────────────────────────────────────────────────────────────────
let _liveCache    = null, _liveCacheTs   = 0;
let _histCache    = null, _histCacheTs   = 0;
let _histBuilding = false;

const LIVE_TTL = 5  * 60 * 1000;  // 5 min — live markets change often
const HIST_TTL = 60 * 60 * 1000;  // 1 hr  — historical rarely changes

// Build the historical database (runs once per hour, returns cached after)
async function buildHistory(force) {
  const now = Date.now();
  if (!force && _histCache && now - _histCacheTs < HIST_TTL) return _histCache;
  if (_histBuilding) return _histCache || [];

  _histBuilding = true;
  console.log('[history] Building full database — this may take ~15s on first run…');
  const t0 = Date.now();

  try {
    const batches = await Promise.all(HIST_TAGS.map(fetchAllClosed));
    const seen    = new Set();
    const events  = [];

    for (const batch of batches) addBatch(batch, seen, events, { historical: true });

    events.sort((a, b) => {
      // Sort by endDate desc (most recent first)
      const da = a.endDate ? new Date(a.endDate).getTime() : 0;
      const db = b.endDate ? new Date(b.endDate).getTime() : 0;
      return db - da;
    });

    _histCache   = events;
    _histCacheTs = now;
    console.log(`[history] Done — ${events.length} historical events in ${((Date.now()-t0)/1000).toFixed(1)}s`);
  } catch (err) {
    console.error('[history] Error:', err.message);
  } finally {
    _histBuilding = false;
  }

  return _histCache || [];
}

// ── API endpoint ───────────────────────────────────────────────────────────
app.get('/api/markets', async (req, res) => {
  const force = req.query.force === '1';
  const now   = Date.now();

  // Serve full combined cache if fresh
  if (!force && _liveCache && now - _liveCacheTs < LIVE_TTL) {
    return res.json(_liveCache);
  }

  try {
    // Live fetch (fast, ~1s) — process all LIVE_TAGS, not just 3
    const liveBatches = await Promise.all(LIVE_TAGS.map(fetchLive));

    const seen   = new Set();
    const events = [];
    for (const batch of liveBatches) addBatch(batch, seen, events);

    // Historical (from cache if available, else triggers background build)
    const hist = await buildHistory(force);
    for (const ev of hist) {
      if (!seen.has(ev.id)) { seen.add(ev.id); events.push(ev); }
    }

    events.sort((a, b) => b.volume - a.volume);

    const liveCount = events.filter(e => !e.historical).length;
    const histCount = events.filter(e =>  e.historical).length;

    const payload = {
      events,
      fetchedAt:  new Date().toISOString(),
      count:      events.length,
      liveCount,
      histCount,
      histReady:  !!_histCache,
    };

    _liveCache   = payload;
    _liveCacheTs = now;
    res.json(payload);
  } catch (err) {
    console.error('[api/markets]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── Kalshi HTTP helper ─────────────────────────────────────────────────────
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

// ── Kalshi live cache ──────────────────────────────────────────────────────
let _kalshiCache = null, _kalshiCacheTs = 0;
const KALSHI_TTL  = 5 * 60 * 1000;

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

// GET /api/kalshi/live — open Kalshi entertainment markets
app.get('/api/kalshi/live', async (req, res) => {
  const now = Date.now();
  if (_kalshiCache && now - _kalshiCacheTs < KALSHI_TTL) return res.json(_kalshiCache);
  try {
    const data   = await kalshiGet('/trade-api/v2/events?status=open&with_nested_markets=true&limit=200');
    const events = (data.events || [])
      .filter(ev => (ev.category || '').toLowerCase() === 'entertainment')
      .map(parseKalshiEvent)
      .filter(ev => ev.markets.length > 0)
      .sort((a, b) => b.volume - a.volume);

    const payload = { events, fetchedAt: new Date().toISOString(), count: events.length };
    _kalshiCache   = payload;
    _kalshiCacheTs = now;
    res.json(payload);
  } catch (err) {
    console.error('[api/kalshi/live]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/history — paginated Polymarket historical entertainment markets
app.get('/api/history', async (req, res) => {
  const { page = '1', perPage = '100', search = '', sort = 'date', dir = 'desc' } = req.query;
  const pg    = Math.max(1, parseInt(page,    10) || 1);
  const limit = Math.min(500, parseInt(perPage, 10) || 100);
  try {
    const hist = await buildHistory(false); // already entertainment-filtered by isFilmEvent()

    const rows = [];
    for (const ev of hist) {
      for (const m of ev.markets) {
        rows.push({
          source: 'polymarket', eventTitle: ev.title, question: m.question,
          slug: ev.slug, yesPrice: m.yesPrice, noPrice: m.noPrice,
          volume: m.volume, endDate: ev.endDate,
          outcomes: m.outcomes || ['Yes', 'No'],
          volume1wk: m.volume1wk || 0, volume1mo: m.volume1mo || 0,
        });
      }
    }

    let filtered = rows;
    if (search) {
      const q = search.toLowerCase();
      filtered = rows.filter(r => (r.eventTitle + ' ' + r.question).toLowerCase().includes(q));
    }

    if (sort === 'volume') {
      filtered.sort((a, b) => dir === 'asc' ? a.volume - b.volume : b.volume - a.volume);
    } else {
      filtered.sort((a, b) => {
        const da = a.endDate ? new Date(a.endDate).getTime() : 0;
        const db = b.endDate ? new Date(b.endDate).getTime() : 0;
        return dir === 'asc' ? da - db : db - da;
      });
    }

    const total = filtered.length;
    res.json({ rows: filtered.slice((pg - 1) * limit, pg * limit), total, page: pg, perPage: limit });
  } catch (err) {
    console.error('[api/history]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/price-history?market=<id> — proxy Polymarket Gamma prices-history
app.get('/api/price-history', async (req, res) => {
  const { market, fidelity = '60' } = req.query;
  if (!market) return res.status(400).json({ error: 'market required' });
  const data = await get(`/prices-history?market=${encodeURIComponent(market)}&interval=max&fidelity=${fidelity}`);
  res.json(Array.isArray(data) ? data : []);
});

// GET /api/kalshi/price-history/:ticker — proxy Kalshi market history
app.get('/api/kalshi/price-history/:ticker', async (req, res) => {
  const data = await kalshiGet(`/trade-api/v2/markets/${encodeURIComponent(req.params.ticker)}/history?limit=100`);
  res.json(data.history || []);
});

// Kick off the history build in the background immediately on startup
// so the first real request doesn't wait for it
buildHistory(false).catch(() => {});

app.listen(PORT, () => {
  console.log(`\n  PolyFilm — Box Office Intelligence\n  → http://localhost:${PORT}\n`);
});
