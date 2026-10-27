import {build} from 'esbuild';
import fs from 'node:fs/promises';
await fs.mkdir('panel',{recursive:true});
await build({entryPoints:['src/main.ts'],bundle:true,outfile:'panel/main.js',format:'iife',platform:'browser',target:['chrome120'],minify:true,legalComments:'eof',define:{'process.env.NODE_ENV':'"production"'}});
await fs.copyFile('src/style.css','panel/style.css');
console.log('Office Preview built');
