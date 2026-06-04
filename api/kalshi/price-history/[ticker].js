'use strict';
const { kalshiGet } = require('../../../lib/kalshi');

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  const { ticker } = req.query;
  if (!ticker) return res.status(400).json({ error: 'ticker required' });

  const data = await kalshiGet(
    `/trade-api/v2/markets/${encodeURIComponent(ticker)}/history?limit=100`
  );
  res.json(data.history || []);
};
