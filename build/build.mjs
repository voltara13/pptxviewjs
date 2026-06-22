import esbuild from 'esbuild';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');

const shim = {
  name: 'global-externals',
  setup(b) {
    b.onResolve({ filter: /^chart\.js\/auto$/ }, () => ({ path: path.join(__dirname, 'shim-chart.js') }));
    b.onResolve({ filter: /^jszip$/ }, () => ({ path: path.join(__dirname, 'shim-jszip.js') }));
  },
};

const minify = process.argv.includes('--min');
await esbuild.build({
  entryPoints: [path.join(root, 'src/index.js')],
  bundle: true,
  format: 'iife',
  globalName: 'PptxViewJS',
  platform: 'browser',
  target: 'es2019',
  minify,
  legalComments: 'none',
  plugins: [shim],
  outfile: path.join(root, 'dist', minify ? 'PptxViewJS.min.js' : 'PptxViewJS.js'),
});
console.log('built', minify ? 'PptxViewJS.min.js' : 'PptxViewJS.js');
