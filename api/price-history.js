'use strict';
const { get } = require('../lib/polymarket');

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  const { market, fidelity = '60' } = req.query;
  if (!market) return res.status(400).json({ error: 'market query param required' });

  const data = await get(
    `/prices-history?market=${encodeURIComponent(market)}&interval=max&fidelity=${fidelity}`
  );
  res.json(Array.isArray(data) ? data : []);
};
