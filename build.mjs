import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'fs';
await build({ entryPoints: ['src/main.js'], bundle: true, format: 'iife', platform: 'browser', minify: true, outfile: 'dist/bundle.js', logLevel: 'warning' });
const shell = readFileSync('src/shell.html', 'utf8');
const js = readFileSync('dist/bundle.js', 'utf8').replace(/<\/script/g, '<\\/script');
writeFileSync('index.html', shell.replace('/*BUNDLE*/', () => js));
console.log('built index.html');
