/**
 * Runs every *.test.js in this directory, in a child process each, and exits
 * non-zero if any fail. No dependencies: `npm test` works on a clean clone.
 */
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const files = fs.readdirSync(__dirname)
  .filter(f => f.endsWith('.test.js'))
  .sort();

if (!files.length) {
  console.error('No test files found in ' + __dirname);
  process.exit(1);
}

let failed = 0;
files.forEach(file => {
  console.log('\n' + '='.repeat(72) + '\n' + file + '\n' + '='.repeat(72));
  const result = spawnSync(process.execPath, [path.join(__dirname, file)], { stdio: 'inherit' });
  if (result.status !== 0) failed++;
});

console.log('\n' + '='.repeat(72));
console.log(failed === 0
  ? `${files.length} test file(s) passed.`
  : `${failed} of ${files.length} test file(s) FAILED.`);
process.exit(failed === 0 ? 0 : 1);
