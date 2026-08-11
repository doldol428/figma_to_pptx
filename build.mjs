import { build, context } from 'esbuild';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(fileURLToPath(import.meta.url));
const dist = resolve(root, 'dist');
const watch = process.argv.includes('--watch');

const common = {
  bundle: true,
  platform: 'browser',
  target: 'es2017',
  format: 'iife',
  logLevel: 'info',
  minify: !watch,
};

/** UI 번들은 파일이 아니라 문자열로 받아 ui.html 안에 인라인한다 (Figma 는 단일 HTML 만 받는다). */
async function buildUi() {
  const result = await build({
    ...common,
    entryPoints: [resolve(root, 'src/ui/ui.ts')],
    write: false,
    outfile: 'ui.js',
  });

  const js = result.outputFiles[0].text;
  const template = await readFile(resolve(root, 'src/ui/ui.html'), 'utf8');
  // 번들 안에 </script> 리터럴이 있으면 HTML 파서가 스크립트를 조기 종료시킨다.
  const safe = js.replaceAll('</script', '<\\/script');
  await writeFile(resolve(dist, 'ui.html'), template.replace('__SCRIPT__', () => safe), 'utf8');
  console.log(`  dist/ui.html  ${(safe.length / 1024).toFixed(0)}kb inline`);
}

async function buildMain(opts) {
  await build({
    ...common,
    ...opts,
    entryPoints: [resolve(root, 'src/main/code.ts')],
    outfile: resolve(dist, 'code.js'),
  });
}

await mkdir(dist, { recursive: true });

if (watch) {
  const ctx = await context({
    ...common,
    entryPoints: [resolve(root, 'src/main/code.ts')],
    outfile: resolve(dist, 'code.js'),
    plugins: [{
      name: 'ui-rebuild',
      setup(b) {
        b.onEnd(() => buildUi());
      },
    }],
  });
  await ctx.watch();
  console.log('watching…');
} else {
  await buildMain({});
  await buildUi();
}
