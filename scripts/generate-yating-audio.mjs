import { execFileSync } from 'node:child_process';
import { mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';

const manifestPath = path.resolve('assets/audio/manifest.json');
const outputDir = path.resolve('assets/audio/questions');
const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
const key = execFileSync(
  'security',
  ['find-generic-password', '-s', 'TOCFL Yating Trial Key', '-w'],
  { encoding: 'utf8', stdio: ['ignore', 'pipe', 'inherit'] },
).trim();

if (!key) throw new Error('找不到 TOCFL Yating Trial Key。');
await mkdir(outputDir, { recursive: true });

const models = {
  yating: 'zh_en_female_1',
  jiahao: 'zh_en_male_1',
};

async function alreadyGenerated(filePath) {
  try {
    return (await stat(filePath)).size > 44;
  } catch {
    return false;
  }
}

async function generate(item) {
  const outputPath = path.join(outputDir, item.file);
  if (await alreadyGenerated(outputPath)) return { id: item.id, skipped: true };

  let lastError;
  for (let attempt = 1; attempt <= 4; attempt += 1) {
    try {
      const response = await fetch('https://tts.api.yating.tw/v2/speeches/short', {
        method: 'POST',
        headers: { key, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          input: { text: item.text, type: 'text' },
          voice: { model: models[item.voice], speed: 1.1, pitch: 1, energy: 1 },
          audioConfig: { encoding: 'LINEAR16', sampleRate: '22K' },
        }),
      });
      const result = await response.json();
      if (!response.ok || !result.audioContent) {
        throw new Error(`${response.status} ${JSON.stringify(result)}`);
      }
      await writeFile(outputPath, Buffer.from(result.audioContent, 'base64'));
      return { id: item.id, skipped: false };
    } catch (error) {
      lastError = error;
      if (attempt < 4) await new Promise(resolve => setTimeout(resolve, 1000 * attempt));
    }
  }
  throw new Error(`${item.id} 產生失敗：${lastError?.message || lastError}`);
}

let nextIndex = 0;
let completed = 0;
let skipped = 0;
const failures = [];

async function worker() {
  while (true) {
    const index = nextIndex++;
    if (index >= manifest.items.length) return;
    const item = manifest.items[index];
    try {
      const result = await generate(item);
      completed += 1;
      if (result.skipped) skipped += 1;
      if (completed % 10 === 0 || completed === manifest.items.length) {
        console.log(`進度 ${completed}/${manifest.items.length}（沿用 ${skipped}）`);
      }
    } catch (error) {
      failures.push({ id: item.id, error: error.message });
      console.error(error.message);
    }
  }
}

await Promise.all([worker(), worker(), worker()]);

if (failures.length) {
  console.error(JSON.stringify(failures, null, 2));
  process.exitCode = 1;
} else {
  console.log(`完成：${completed} 題，新增 ${completed - skipped} 題，沿用 ${skipped} 題。`);
}
