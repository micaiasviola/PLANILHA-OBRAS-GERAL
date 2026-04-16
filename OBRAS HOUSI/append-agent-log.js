// tools: scripts/append-agent-log.js
// Appends a single line to AGENT-LOGS.md and commits/pushes it.
// Usage: node scripts/append-agent-log.js "ACTION" "descrição curta" [PR_NUMBER]

const { execSync } = require('child_process');
const fs = require('fs');

const ACTION = process.argv[2] || 'ACTION';
const DESC = process.argv[3] || '';
const PR = process.argv[4] ? `| pr: ${process.argv[4]}` : '';
const AGENT_NAME = process.env.AGENT_NAME || 'LocalAgent';
const PATH = 'AGENT-LOGS.md';

function shortSha() {
  try {
    return execSync('git rev-parse --short HEAD').toString().trim();
  } catch (e) { return ''; }
}

const ts = new Date().toISOString();
const sha = shortSha();
const line = `- ${ts} | ${AGENT_NAME} | ${ACTION} | ${DESC} | commit: ${sha} ${PR}\n`;

try {
  fs.appendFileSync(PATH, line, 'utf8');
  console.log('Appended to', PATH, ':', line.trim());
} catch (e) {
  console.error('Failed to append to', PATH, e.message);
  process.exit(1);
}

try {
  execSync('git add "' + PATH + '"', { stdio: 'inherit' });
  execSync('git commit -m "agent-log: ' + ts + ' ' + ACTION + '"', { stdio: 'inherit' });
  execSync('git push', { stdio: 'inherit' });
  console.log('Committed and pushed agent log.');
} catch (e) {
  console.error('Git commit/push failed:', e.message);
  // Don't fail the entire process; the agent can retry.
}
