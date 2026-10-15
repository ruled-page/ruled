/**
 * The steps of a release that change the repository, for the release
 * workflows in .github/workflows (see README "Releases"):
 *
 *   node scripts/release.mjs prepare <version | patch | minor | major>
 *       Sets the packages, @ruled/formula, @ruled/sheet and @ruled/plugin, to
 *       the version (or the next patch, minor or major one after the last
 *       release), moves the changelog's Unreleased notes under it and writes
 *       the third-party notices again, which needs `npm run build:lib` first.
 *       Commits nothing. Prints the version, and with GITHUB_OUTPUT set
 *       writes it there as `version`, and whether it's a prerelease.
 *
 *   node scripts/release.mjs check <version | patch | minor | major>
 *       What prepare checks first, that the version comes after the last
 *       release and the changelog has notes to release, changing nothing:
 *       for a workflow to fail before it installs and builds.
 *
 *   node scripts/release.mjs notes [version]
 *       Prints a version's notes from the changelog: by default the version
 *       the packages are at.
 */
import { execFileSync, execSync } from 'node:child_process';
import { appendFileSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const sheet = join(root, 'packages', 'spreadsheet');
const changelogPath = join(sheet, 'CHANGELOG.md');
const git = (...args) => execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();
const run = (command) => execSync(command, { cwd: root, stdio: 'inherit' });
const fail = (message) => {
  console.error(message);
  process.exit(1);
};

const SEMVER = /^(\d+)\.(\d+)\.(\d+)(-[0-9A-Za-z.-]+)?$/;
const order = (v) => SEMVER.exec(v).slice(1, 4).map(Number);
/** Whether version `a` comes after `b`: a higher number, or the release of the prerelease `b`. */
const later = (a, b) => {
  const [x, y] = [order(a), order(b)];
  const i = x.findIndex((n, k) => n !== y[k]);
  return i < 0 ? a !== b && !a.includes('-') && b.includes('-') : x[i] > y[i];
};

/** A version's notes: what the changelog has under its heading, up to the next. */
function notes(version) {
  const changelog = readFileSync(changelogPath, 'utf8');
  const escaped = version.replace(/[.+-]/g, '\\$&');
  const found = new RegExp(`^## ${escaped} \\(.*\\)\\n([\\s\\S]*?)(?=^## |(?![\\s\\S]))`, 'm').exec(changelog);
  if (!found) fail(`CHANGELOG.md has no notes under "## ${version}".`);
  return found[1].trim();
}

/** The version a release of `asked` makes, and the changelog with its Unreleased notes; fails if either won't do. */
function checked(asked) {
  if (!asked) fail('Say which version: node scripts/release.mjs prepare <version | patch | minor | major>');
  // Released versions are the v-tags; before the first, the last is 0.0.0.
  const last = git('tag', '--list', 'v*', '--sort=-v:refname').split('\n')[0]?.replace(/^v/, '') || '0.0.0';
  const [major, minor, patch] = order(last);
  const next = { major: `${major + 1}.0.0`, minor: `${major}.${minor + 1}.0`, patch: `${major}.${minor}.${patch + 1}` }[asked] ?? asked;
  if (!SEMVER.test(next)) fail(`"${asked}" isn't a version (such as 1.2.3 or 1.2.3-beta.1) or patch, minor or major.`);
  if (last !== '0.0.0' && !later(next, last)) fail(`${next} doesn't come after the last release, ${last}.`);

  const changelog = readFileSync(changelogPath, 'utf8');
  const unreleased = /^## Unreleased\n([\s\S]*?)(?=^## |(?![\s\S]))/m.exec(changelog);
  if (!unreleased || !unreleased[1].trim()) fail('CHANGELOG.md has no notes under "## Unreleased" to release.');
  return { next, changelog, unreleased };
}

function prepare(asked) {
  const { next, changelog, unreleased } = checked(asked);

  // The packages at the version, and every workspace that names one of them, so each still takes the workspace's copy.
  // A range keeps its operator: @ruled/plugin's peer range for @ruled/sheet moves to ^version.
  for (const workspace of JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')).workspaces) {
    const path = join(root, workspace, 'package.json');
    const manifest = readFileSync(path, 'utf8');
    const named = /"name": "(@ruled\/(?:formula|sheet|plugin))"/.exec(manifest);
    writeFileSync(path, manifest.replace(/("@ruled\/(?:formula|sheet|plugin)": "(?:[~^]|>=)?)[^"]*"/g, `$1${next}"`).replace(/^(  "version": ")[^"]*"/m, (line, start) => (named ? `${start}${next}"` : line)));
  }
  run('npm install --package-lock-only --ignore-scripts --no-audit --no-fund');

  const date = new Date().toISOString().slice(0, 10);
  writeFileSync(changelogPath, changelog.replace(unreleased[0], () => `## Unreleased\n\n## ${next} (${date})\n${unreleased[1]}`));
  run('npm run notices');

  console.log(next);
  if (process.env.GITHUB_OUTPUT) appendFileSync(process.env.GITHUB_OUTPUT, `version=${next}\nprerelease=${next.includes('-')}\n`);
}

const [command, arg] = process.argv.slice(2);
if (command === 'prepare') prepare(arg);
else if (command === 'check') console.log(checked(arg).next);
else if (command === 'notes') console.log(notes(arg ?? JSON.parse(readFileSync(join(sheet, 'package.json'), 'utf8')).version));
else fail('Usage: node scripts/release.mjs prepare <version | patch | minor | major>, check <the same>, or notes [version]');
