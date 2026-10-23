import test from 'node:test';
import { execFileSync } from 'node:child_process';
test('Python CLI fake HTTP suite (11 cases, no live service)', () => {
  execFileSync('python3', ['-B', 'tests/bridge-client.test.py'], { stdio: 'pipe' });
});
