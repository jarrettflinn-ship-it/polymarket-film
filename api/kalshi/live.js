'use strict';
const { kalshiGet, parseKalshiEvent } = require('../../lib/kalshi');

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const data   = await kalshiGet(
      '/trade-api/v2/events?status=open&with_nested_markets=true&limit=200'
    );
    const events = (data.events || [])
      .filter(ev => (ev.category || '').toLowerCase() === 'entertainment')
      .map(parseKalshiEvent)
      .filter(ev => ev.markets.length > 0)
      .sort((a, b) => b.volume - a.volume);

    res.json({ events, fetchedAt: new Date().toISOString(), count: events.length });
  } catch (err) {
    console.error('[api/kalshi/live]', err.message);
    res.status(500).json({ error: err.message });
  }
};
