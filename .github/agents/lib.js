#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

// Central agent utilities
const repoRoot = path.resolve(__dirname, '..', '..');
const cwd = repoRoot;

function listFiles(dir) {
  const out = [];
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const e of entries) {
    const full = path.join(dir, e.name);
    if (e.isDirectory()) {
      if (['node_modules', '.git'].includes(e.name)) continue;
      out.push(...listFiles(full));
    } else {
      if (full.endsWith('.gs') || full.endsWith('.js') || full.endsWith('.html')) out.push(full);
    }
  }
  return out;
}

module.exports = {
  repoRoot,
  cwd,
  listFiles,
  spawnSync
};
