/**
 * Makes a release of the package:
 *
 *   npm run release -- 0.2.0          a version
 *   npm run release -- patch          or the next patch, minor or major
 *                                     version after the last release
 *
 * From a clean working tree, it runs the checks CI runs short of the
 * browser tests, sets the version, moves the changelog's Unreleased notes
 * under it, writes the third-party notices again, commits, tags vX.Y.Z and
 * packs the package into releases/. Nothing is pushed or published.
 */
import { execFileSync, execSync } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
/** The package released: the editor, @ruled/sheet. */
const pkg = join(root, 'packages', 'spreadsheet');
const git = (...args) => execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();
const run = (command) => execSync(command, { cwd: root, stdio: 'inherit' });
const fail = (message) => {
  console.error(message);
  process.exit(1);
};

const SEMVER = /^(\d+)\.(\d+)\.(\d+)(-[0-9A-Za-z.-]+)?$/;
const asked = process.argv[2];
if (!asked) fail('Say which version: npm run release -- <version | patch | minor | major>');

// Released versions are the v-tags; before the first, the last is 0.0.0.
const last = git('tag', '--list', 'v*', '--sort=-v:refname').split('\n')[0]?.replace(/^v/, '') || '0.0.0';
const [major, minor, patch] = SEMVER.exec(last).slice(1, 4).map(Number);
const next = { major: `${major + 1}.0.0`, minor: `${major}.${minor + 1}.0`, patch: `${major}.${minor}.${patch + 1}` }[asked] ?? asked;
if (!SEMVER.test(next)) fail(`"${asked}" isn't a version (such as 1.2.3) or patch, minor or major.`);
const order = (v) => SEMVER.exec(v).slice(1, 4).map(Number);
const later = (a, b) => {
  const [x, y] = [order(a), order(b)];
  const i = x.findIndex((n, k) => n !== y[k]);
  return i < 0 ? a !== b && !a.includes('-') && b.includes('-') : x[i] > y[i];
};
if (last !== '0.0.0' && !later(next, last)) fail(`${next} doesn't come after the last release, ${last}.`);

// Untracked files can't reach the package, which ships only what package.json lists.
if (git('status', '--porcelain', '--untracked-files=no')) fail('Commit or put aside your changes first: a release is made from a clean working tree.');

const changelogPath = join(pkg, 'CHANGELOG.md');
const changelog = readFileSync(changelogPath, 'utf8');
const unreleased = /^## Unreleased\n([\s\S]*?)(?=^## |(?![\s\S]))/m.exec(changelog);
if (!unreleased || !unreleased[1].trim()) fail('CHANGELOG.md has no notes under "## Unreleased" to release.');

for (const script of ['typecheck', 'lint', 'test', 'build', 'build:lib', 'check:bundle', 'validate:saves']) run(`npm run ${script}`);

run(`npm version ${next} --no-git-tag-version --allow-same-version -w @ruled/sheet`);
const date = new Date().toISOString().slice(0, 10);
writeFileSync(changelogPath, changelog.replace(unreleased[0], () => `## Unreleased\n\n## ${next} (${date})\n${unreleased[1]}`));
run('npm run notices');

git('add', 'packages/spreadsheet/package.json', 'package-lock.json', 'packages/spreadsheet/CHANGELOG.md', 'packages/spreadsheet/THIRD-PARTY-NOTICES.md');
git('commit', '-m', `Release ${next}`);
git('tag', '-a', `v${next}`, '-m', `@ruled/sheet ${next}`);

const releases = join(root, 'releases');
mkdirSync(releases, { recursive: true });
const [{ filename }] = JSON.parse(execFileSync('npm', ['pack', '--json', '--pack-destination', releases], { cwd: pkg, encoding: 'utf8' }));
console.log(`\nReleased ${next}: tagged v${next}, packed releases/${filename}. Push the tag and share or publish the tarball when ready.`);
