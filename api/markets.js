'use strict';
const { LIVE_TAGS, fetchLive, fetchRecentClosed, addBatch } = require('../lib/polymarket');

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  try {
    // Fetch live + one page of recent historical in parallel (fast, ~2s total)
    const [liveBatches, histBatches] = await Promise.all([
      Promise.all(LIVE_TAGS.map(fetchLive)),
      Promise.all(LIVE_TAGS.map(tag => fetchRecentClosed(tag, 100))),
    ]);

    const seen   = new Set();
    const events = [];
    for (const batch of liveBatches) addBatch(batch, seen, events);
    for (const batch of histBatches) addBatch(batch, seen, events, { historical: true });

    events.sort((a, b) => b.volume - a.volume);

    const liveCount = events.filter(e => !e.historical).length;
    const histCount = events.filter(e =>  e.historical).length;

    res.json({
      events,
      fetchedAt: new Date().toISOString(),
      count:     events.length,
      liveCount,
      histCount,
      histReady: true,
    });
  } catch (err) {
    console.error('[api/markets]', err.message);
    res.status(500).json({ error: err.message });
  }
};
