'use strict';

const express = require('express');
const DEFAULTS = Object.freeze({
  bodyBytes: 1024 * 1024, bodyMs: 15000, responseMs: 60000,
  burst: 120, perSecond: 2, maxConnections: 32, maxJobs: 16,
});

function install(app, overrides = {}) {
  const limits = { ...DEFAULTS, ...overrides };
  app.set('query parser', false);
  app.disable('x-powered-by');
  // Global, bounded state: do not trust client-supplied forwarding headers.
  let tokens = limits.burst, last = performance.now(), connections = 0, jobs = 0;
  app.use((req, res, next) => {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('Cache-Control', 'no-store');
    const now = performance.now();
    tokens = Math.min(limits.burst, tokens + (now - last) * limits.perSecond / 1000);
    last = now;
    const webhook = req.method === 'POST' && /^\/webhook\/?$/i.test(req.path);
    if (connections >= limits.maxConnections || tokens < 1) {
      res.setHeader('Retry-After', '1');
      return res.status(429).end();
    }
    tokens--;
    connections++;
    let released = false;
    const release = () => { if (!released) { released = true; connections--; clearTimeout(timer); } };
    // A response deadline never cancels or releases an accepted accounting job.
    const timer = setTimeout(() => { if (!res.writableEnded) res.status(503).end(); }, limits.responseMs);
    timer.unref?.();
    res.once('finish', release); res.once('close', release);
    if (!webhook && !['GET', 'HEAD'].includes(req.method)) return res.status(404).end();
    next();
  });
  const raw = express.raw({ type: () => true, limit: limits.bodyBytes, inflate: false });
  function body(req, res, next) {
    if (req.headers['content-encoding'] && req.headers['content-encoding'] !== 'identity') return res.status(415).end();
    if (!/^application\/json(?:\s*;|$)/i.test(req.headers['content-type'] || '')) return res.status(415).end();
    const length = Number(req.headers['content-length']);
    if (length > limits.bodyBytes) { res.setHeader('Connection', 'close'); return res.status(413).end(); }
    const timer = setTimeout(() => {
      if (!res.writableEnded) { res.setHeader('Connection', 'close'); res.status(408).end(); }
      req.destroy();
    }, limits.bodyMs);
    timer.unref?.();
    raw(req, res, err => {
      clearTimeout(timer);
      if (res.writableEnded || res.destroyed) return;
      if (err) return res.status(err.type === 'entity.too.large' ? 413 : 400).end();
      if (!Buffer.isBuffer(req.body)) return res.status(400).end();
      next();
    });
    req.once('aborted', () => clearTimeout(timer));
  }
  function admit(req, res, next) {
    if (!req.body || !Array.isArray(req.body.events) || req.body.events.length > 100 ||
        req.body.events.some(e => !e || typeof e !== 'object' || Array.isArray(e))) return res.status(400).end();
    if (res.writableEnded || res.destroyed) return;
    if (jobs >= limits.maxJobs) { res.setHeader('Retry-After', '1'); return res.status(503).end(); }
    jobs++;
    let done = false;
    req.releaseHttpJob = () => { if (!done) { done = true; jobs--; } };
    next();
  }
  function errors(err, req, res, next) {
    // SDK errors can contain raw JSON and signatures. Never log or return them.
    if (res.headersSent) { if (!res.destroyed) res.destroy(); return; }
    res.status(err?.name === 'SignatureValidationFailed' ? 401 : 400).end();
  }
  return { body, admit, errors, limits };
}
function configureServer(server) {
  server.maxConnections = 128;
  server.headersTimeout = 10000;
  server.requestTimeout = 20000;
  server.keepAliveTimeout = 5000;
  server.maxRequestsPerSocket = 100;
}
module.exports = { install, configureServer, DEFAULTS };
