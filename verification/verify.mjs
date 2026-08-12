import crypto from 'node:crypto';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { spawn } from 'node:child_process';
import { inflateRawSync } from 'node:zlib';

const repoRoot = path.resolve(import.meta.dirname, '..');
const artifactRoot = path.join(repoRoot, 'artifacts');
const evidenceRoot = path.join(repoRoot, 'verification', 'evidence');
const npmCli = process.env.npm_execpath;
const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm';
const attachments = ['输入数据包.zip', 'reference.zip', '关键标准答案.xlsx', '任务规格转化.xlsx'];
const expectedReference = [
  'output/playwright.config.ts',
  'output/reports/keyboard_focus_report.csv',
  'output/reports/save_rollback_report.csv',
  'output/reports/settings_state_report.csv',
  'output/tests/subscription_settings.spec.ts',
].sort();
const reportKeys = {
  'output/reports/keyboard_focus_report.csv': ['order_index'],
  'output/reports/save_rollback_report.csv': ['scenario_id'],
  'output/reports/settings_state_report.csv': ['scenario_id'],
};

const sha256 = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex');
const sha256File = (file) => sha256(fs.readFileSync(file));
const assert = (value, message) => { if (!value) throw new Error(message); };

function parseZipBytes(data) {
  const files = new Map();
  let offset = 0;
  while (offset + 30 <= data.length) {
    if (data.readUInt32LE(offset) !== 0x04034b50) break;
    const flags = data.readUInt16LE(offset + 6);
    const method = data.readUInt16LE(offset + 8);
    const compressedSize = data.readUInt32LE(offset + 18);
    const uncompressedSize = data.readUInt32LE(offset + 22);
    const nameLength = data.readUInt16LE(offset + 26);
    const extraLength = data.readUInt16LE(offset + 28);
    assert(!(flags & 0x08), 'ZIP数据描述符不受支持');
    const name = data.subarray(offset + 30, offset + 30 + nameLength).toString('utf8').replaceAll('\\', '/');
    const start = offset + 30 + nameLength + extraLength;
    const compressed = data.subarray(start, start + compressedSize);
    if (!name.endsWith('/')) {
      const body = method === 0 ? compressed : method === 8 ? inflateRawSync(compressed) : null;
      assert(body && body.length === uncompressedSize, `无法解压${name}`);
      files.set(name, body);
    }
    offset = start + compressedSize;
  }
  return files;
}

const parseZip = (file) => parseZipBytes(fs.readFileSync(file));
async function extractZip(file, destination) {
  for (const [name, bytes] of parseZip(file)) {
    const target = path.resolve(destination, name);
    assert(target.startsWith(path.resolve(destination) + path.sep), `非法ZIP路径${name}`);
    await fsp.mkdir(path.dirname(target), { recursive: true });
    await fsp.writeFile(target, bytes);
  }
}

function workbookSheets(file) {
  const workbook = parseZipBytes(fs.readFileSync(file)).get('xl/workbook.xml')?.toString('utf8') ?? '';
  return [...workbook.matchAll(/<(?:[A-Za-z]+:)?sheet[^>]+name="([^"]+)"/gu)].map((match) => match[1]);
}

async function run(command, args, cwd) {
  const started = Date.now();
  return await new Promise((resolve) => {
    let child;
    try { child = spawn(command, args, { cwd, env: process.env, windowsHide: true }); }
    catch (error) { resolve({ code: 1, stdout: '', stderr: error.stack ?? error.message, elapsed_ms: Date.now() - started }); return; }
    let stdout = '';
    let stderr = '';
    let settled = false;
    child.stdout.on('data', (chunk) => { stdout += chunk; });
    child.stderr.on('data', (chunk) => { stderr += chunk; });
    child.on('error', (error) => { if (!settled) { settled = true; resolve({ code: 1, stdout, stderr: `${stderr}${error.stack ?? error.message}`, elapsed_ms: Date.now() - started }); } });
    child.on('exit', (code) => { if (!settled) { settled = true; resolve({ code: code ?? 1, stdout, stderr, elapsed_ms: Date.now() - started }); } });
  });
}

async function runNpm(args, cwd) {
  return npmCli ? await run(process.execPath, [npmCli, ...args], cwd) : await run(npmCommand, args, cwd);
}

function treeDigest(root, ignored = new Set()) {
  const lines = [];
  function visit(current, prefix = '') {
    for (const entry of fs.readdirSync(current, { withFileTypes: true }).toSorted((left, right) => left.name.localeCompare(right.name))) {
      const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
      if (ignored.has(relative.split('/')[0])) continue;
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) visit(full, relative);
      else lines.push(`${relative}\0${sha256File(full)}`);
    }
  }
  visit(root);
  return sha256(Buffer.from(lines.join('\n')));
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = '';
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (quoted) {
      if (char === '"' && text[index + 1] === '"') { cell += '"'; index += 1; }
      else if (char === '"') quoted = false;
      else cell += char;
    } else if (char === '"') quoted = true;
    else if (char === ',') { row.push(cell); cell = ''; }
    else if (char === '\n') { row.push(cell.replace(/\r$/u, '')); rows.push(row); row = []; cell = ''; }
    else cell += char;
  }
  if (cell || row.length) { row.push(cell.replace(/\r$/u, '')); rows.push(row); }
  const headers = rows.shift() ?? [];
  return rows.filter((values) => values.some((value) => value !== '')).map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ''])));
}

function csvCell(value) {
  const text = String(value ?? '');
  return /[",\r\n]/u.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function toCsv(headers, rows, eol = '\n') {
  return `${headers.join(',')}${eol}${rows.map((row) => headers.map((header) => csvCell(row[header])).join(',')).join(eol)}${eol}`;
}

function normalizedRows(file, text) {
  const keys = reportKeys[file];
  return parseCsv(text).toSorted((left, right) => keys.map((key) => String(left[key]).localeCompare(String(right[key]))).find((value) => value !== 0) ?? 0);
}

function classifyExecutable(name, bytes) {
  const lower = name.toLowerCase();
  if (bytes.length >= 4 && bytes[0] === 0x7f && bytes.subarray(1, 4).toString('ascii') === 'ELF') return 'linux_elf';
  if (bytes.length >= 4 && [0xfeedface, 0xfeedfacf, 0xcafebabe].includes(bytes.readUInt32BE(0))) return 'macos_macho';
  if (/\.(?:sh|bash|so)(?:\.|$)/u.test(lower)) return 'posix_member';
  if (/^#!.*(?:ba|z|k)?sh/mu.test(bytes.subarray(0, 128).toString('utf8'))) return 'posix_shebang';
  return null;
}

async function prepare(label, mutate) {
  const root = path.join(os.tmpdir(), label);
  await fsp.rm(root, { recursive: true, force: true });
  await fsp.mkdir(root, { recursive: true });
  await extractZip(path.join(artifactRoot, '输入数据包.zip'), root);
  const inputRoot = path.join(root, 'input_data');
  const reference = parseZip(path.join(artifactRoot, 'reference.zip'));
  await fsp.mkdir(path.join(inputRoot, 'output', 'tests'), { recursive: true });
  await fsp.writeFile(path.join(inputRoot, 'output', 'playwright.config.ts'), reference.get('output/playwright.config.ts'));
  await fsp.writeFile(path.join(inputRoot, 'output', 'tests', 'subscription_settings.spec.ts'), reference.get('output/tests/subscription_settings.spec.ts'));
  const dependencyRoot = path.join(repoRoot, 'node_modules');
  await fsp.symlink(dependencyRoot, path.join(root, 'node_modules'), process.platform === 'win32' ? 'junction' : 'dir');
  if (mutate) await mutate(inputRoot);
  return { root, inputRoot, outputRoot: path.join(inputRoot, 'output'), reference };
}

function outputPaths(root) {
  const paths = [];
  function walk(current, prefix = '') {
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
      if (entry.isDirectory()) walk(path.join(current, entry.name), relative);
      else paths.push(`output/${relative}`);
    }
  }
  walk(root);
  return paths.sort();
}

function compareReference(outputRoot, reference) {
  assert(JSON.stringify(outputPaths(outputRoot)) === JSON.stringify(expectedReference), '输出成员与Reference不一致');
  const semantic = crypto.createHash('sha256');
  for (const file of expectedReference) {
    const relative = file.slice('output/'.length);
    const actual = fs.readFileSync(path.join(outputRoot, relative));
    const expected = reference.get(file);
    if (file.endsWith('.csv')) {
      const actualRows = normalizedRows(file, actual.toString('utf8'));
      const expectedRows = normalizedRows(file, expected.toString('utf8'));
      assert(JSON.stringify(actualRows) === JSON.stringify(expectedRows), `${file}与Reference不一致`);
      semantic.update(JSON.stringify(actualRows));
    } else {
      const actualText = actual.toString('utf8').replaceAll('\r\n', '\n');
      const expectedText = expected.toString('utf8').replaceAll('\r\n', '\n');
      assert(actualText === expectedText, `${file}与Reference不一致`);
      semantic.update(actualText);
    }
  }
  return semantic.digest('hex');
}

await fsp.rm(evidenceRoot, { recursive: true, force: true });
await fsp.mkdir(evidenceRoot, { recursive: true });
assert(process.platform === 'win32' && process.env.GITHUB_ACTIONS === 'true', '该验证器只接受GitHub托管Windows运行');

const attachmentSha256 = Object.fromEntries(attachments.map((name) => [name, sha256File(path.join(artifactRoot, name))]));
const inputMembers = parseZip(path.join(artifactRoot, '输入数据包.zip'));
const executableScan = [...inputMembers].map(([name, bytes]) => ({ name, classification: classifyExecutable(name, bytes) })).filter((item) => item.classification);
assert(executableScan.length === 0, `输入包含平台专用可执行成员：${JSON.stringify(executableScan)}`);
const referenceMembers = [...parseZip(path.join(artifactRoot, 'reference.zip')).keys()].sort();
assert(JSON.stringify(referenceMembers) === JSON.stringify(expectedReference), 'Reference成员错误');
assert(JSON.stringify(workbookSheets(path.join(artifactRoot, '关键标准答案.xlsx'))) === JSON.stringify(['交付物答案清单', '固定字段答案', '固定集合答案', '固定数值答案', '允许变体答案']), '关键标准答案Sheet错误');
assert(JSON.stringify(workbookSheets(path.join(artifactRoot, '任务规格转化.xlsx'))) === JSON.stringify(['任务规格转化']), '任务规格Sheet错误');
const solutionText = parseZip(path.join(artifactRoot, 'reference.zip')).get('output/tests/subscription_settings.spec.ts').toString('utf8');
assert(!/S-EMAIL-PUSH|S-SMS-BLOCK|S-EU-CONSENT|S-PUSH-ONLY|S-STALE-ROLLBACK|node:http|node:https|fetch\s*\(/u.test(solutionText), '完成版测试含固定旅程ID硬编码或外部网络实现');

const cleanRuns = [];
for (const label of ['Q10399 第一次 空目录', 'Q10399 第二次 中文 空格目录']) {
  const prepared = await prepare(label);
  const before = treeDigest(prepared.inputRoot, new Set(['output']));
  const result = await runNpm(['run', 'process'], prepared.inputRoot);
  assert(result.code === 0, `${label}执行失败\n${result.stdout}\n${result.stderr}`);
  const after = treeDigest(prepared.inputRoot, new Set(['output']));
  assert(before === after, `${label}修改了输入`);
  const semantic = compareReference(prepared.outputRoot, prepared.reference);
  cleanRuns.push({ directory_label: label, exit_code: result.code, input_digest_before: before, input_digest_after: after, semantic_digest: semantic, elapsed_ms: result.elapsed_ms, chromium_executed: true });
}
assert(cleanRuns[0].semantic_digest === cleanRuns[1].semantic_digest, '两次结构化结果不一致');

const crlf = await prepare('Q10399 CRLF 输入', async (inputRoot) => {
  const file = path.join(inputRoot, 'fixtures', 'save_scenarios.csv');
  const text = await fsp.readFile(file, 'utf8');
  await fsp.writeFile(file, text.replace(/\r?\n/gu, '\r\n'));
});
let result = await runNpm(['run', 'process'], crlf.inputRoot);
assert(result.code === 0, `CRLF输入执行失败\n${result.stdout}\n${result.stderr}`);
const crlfDigest = compareReference(crlf.outputRoot, crlf.reference);
assert(crlfDigest === cleanRuns[0].semantic_digest, 'CRLF输入改变业务结果');

const mutation = await prepare('Q10399 同意地区变化', async (inputRoot) => {
  const file = path.join(inputRoot, 'fixtures', 'consent_policy.json');
  const policy = JSON.parse(await fsp.readFile(file, 'utf8'));
  policy.rollback_rules.explicit_consent_regions = [];
  await fsp.writeFile(file, `${JSON.stringify(policy, null, 2)}\n`);
});
result = await runNpm(['run', 'process'], mutation.inputRoot);
assert(result.code === 0, `同意地区变化执行失败\n${result.stdout}\n${result.stderr}`);
const actualStates = normalizedRows('output/reports/settings_state_report.csv', fs.readFileSync(path.join(mutation.outputRoot, 'reports', 'settings_state_report.csv'), 'utf8'));
const baselineStates = normalizedRows('output/reports/settings_state_report.csv', mutation.reference.get('output/reports/settings_state_report.csv').toString('utf8'));
const changed = actualStates.find((row) => row.scenario_id === 'S-EU-CONSENT');
const beforeChanged = baselineStates.find((row) => row.scenario_id === 'S-EU-CONSENT');
assert(beforeChanged?.result === 'blocked_before_save' && changed?.result === 'saved' && changed?.request_sent === 'true' && changed?.request_revision === 'rev-42', '同意地区变化没有放行目标旅程');
assert(actualStates.filter((row) => row.scenario_id !== 'S-EU-CONSENT').every((row) => JSON.stringify(row) === JSON.stringify(baselineStates.find((item) => item.scenario_id === row.scenario_id))), '同意地区变化影响无关旅程');
const mutatedRollbacks = normalizedRows('output/reports/save_rollback_report.csv', fs.readFileSync(path.join(mutation.outputRoot, 'reports', 'save_rollback_report.csv'), 'utf8'));
assert(!mutatedRollbacks.some((row) => row.scenario_id === 'S-EU-CONSENT'), '放行旅程仍出现在回滚报告');

const negative = await prepare('Q10399 无效输入', async (inputRoot) => {
  await fsp.rm(path.join(inputRoot, 'fixtures', 'consent_policy.json'));
});
result = await runNpm(['run', 'process'], negative.inputRoot);
const deliverablesAbsent = !fs.existsSync(negative.outputRoot) || outputPaths(negative.outputRoot).length === 0;
assert(result.code !== 0 && deliverablesAbsent, '无效输入没有失败关闭');

const evidence = {
  schema_version: 1,
  task_asset_id: 'playwright_subscription_consent_rollback',
  result: 'PASS',
  generated_at_utc: new Date().toISOString(),
  git_commit_sha: process.env.GITHUB_SHA,
  workflow_run_id: process.env.GITHUB_RUN_ID,
  runner: { os: process.env.RUNNER_OS, arch: process.env.RUNNER_ARCH, image_os: process.env.ImageOS, image_version: process.env.ImageVersion, node: process.version, powershell_hosted_workflow: true },
  software: { name: 'Playwright', version: '1.62.0', browser: 'Chromium', executed: true },
  attachment_sha256: attachmentSha256,
  workbook_checks: { answer_sheet_names: workbookSheets(path.join(artifactRoot, '关键标准答案.xlsx')), specification_sheet_names: ['任务规格转化'] },
  platform_audit: { linux_executables: executableScan, linux_executables_executed: false, no_wsl_required: true, no_linux_container_required: true, no_posix_shell_required: true, no_unix_only_api_required: true, cross_platform_paths: true },
  clean_runs: cleanRuns,
  crlf_input: { file: 'fixtures/save_scenarios.csv', exit_code: 0, semantic_digest: crlfDigest, reference_match: true },
  positive_mutation: { changed_input: 'explicit_consent_regions清空', exit_code: 0, affected_scenario: 'S-EU-CONSENT', prior_result: beforeChanged.result, changed_result: changed.result, request_sent: changed.request_sent, unrelated_scenarios_unchanged: true },
  invalid_input: { removed_input: 'fixtures/consent_policy.json', exit_code: result.code, deliverables_absent: deliverablesAbsent },
  network: { installation_network_access: 'npm与Chromium安装阶段', formal_run_network_access: 'loopback only, asserted by Playwright request observer' },
};
await fsp.writeFile(path.join(evidenceRoot, 'windows-verification.json'), `${JSON.stringify(evidence, null, 2)}\n`);
console.log(JSON.stringify(evidence, null, 2));
