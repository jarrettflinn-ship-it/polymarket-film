'use strict';
const { LIVE_TAGS, fetchRecentClosed, addBatch } = require('../lib/polymarket');

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  const {
    page    = '1',
    perPage = '100',
    search  = '',
    sort    = 'date',
    dir     = 'desc',
  } = req.query;

  const pg    = Math.max(1, parseInt(page,    10) || 1);
  const limit = Math.min(500, parseInt(perPage, 10) || 100);

  try {
    // One page per tag in parallel — fast, no pagination timeout risk
    const histBatches = await Promise.all(LIVE_TAGS.map(tag => fetchRecentClosed(tag, 100)));

    const seen   = new Set();
    const events = [];
    for (const batch of histBatches) addBatch(batch, seen, events, { historical: true });

    // Flatten events → individual market rows
    const rows = [];
    for (const ev of events) {
      for (const m of ev.markets) {
        rows.push({
          source:     'polymarket',
          eventTitle: ev.title,
          question:   m.question,
          slug:       ev.slug,
          yesPrice:   m.yesPrice,
          noPrice:    m.noPrice,
          volume:     m.volume,
          endDate:    ev.endDate,
          outcomes:   m.outcomes || ['Yes', 'No'],
          volume1wk:  m.volume1wk || 0,
          volume1mo:  m.volume1mo || 0,
        });
      }
    }

    // Search filter
    let filtered = rows;
    if (search) {
      const q = search.toLowerCase();
      filtered = rows.filter(r => (r.eventTitle + ' ' + r.question).toLowerCase().includes(q));
    }

    // Sort
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
    res.json({
      rows:    filtered.slice((pg - 1) * limit, pg * limit),
      total,
      page:    pg,
      perPage: limit,
    });
  } catch (err) {
    console.error('[api/history]', err.message);
    res.status(500).json({ error: err.message });
  }
};
