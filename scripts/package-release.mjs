import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
import assert from 'node:assert/strict';
import {zipSync} from 'fflate';

export async function packageRelease(root, version) {
  const p=JSON.parse(await fs.readFile(path.join(root,'package.json'),'utf8'));
  assert.equal(p.version,version,'Prepared package version must match semantic-release');
  assert.match(version,/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/);
  const entries={};
  async function collect(relative) {
    const source=path.join(root,relative),stat=await fs.lstat(source);
    assert(!stat.isSymbolicLink(),'Release cannot contain symlinks');
    if(stat.isDirectory()){
      for(const child of (await fs.readdir(source)).sort())await collect(`${relative}/${child}`);
    }else entries[`office-preview/${relative}`]=new Uint8Array(await fs.readFile(source));
  }
  for(const item of ['package.json','README.md','LICENSE','THIRD_PARTY_NOTICES.md','panel','licenses'])await collect(item);
  const dist=path.join(root,'dist');await fs.mkdir(dist,{recursive:true});
  const zipName=`openchamber-office-preview-v${version}.zip`;
  const tgzName=`openchamber-office-preview-${version}.tgz`;
  await fs.writeFile(path.join(dist,zipName),zipSync(entries,{level:6}));
  const sums=[];
  for(const name of [tgzName,zipName]){
    const bytes=await fs.readFile(path.join(dist,name));
    sums.push(`${createHash('sha256').update(bytes).digest('hex')}  ${name}`);
  }
  await fs.writeFile(path.join(dist,'SHA256SUMS.txt'),`${sums.join('\n')}\n`);
  return {zipName,tgzName,entries:Object.keys(entries)};
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
  await packageRelease(process.cwd(),process.argv[2]);
  console.log(`Release archives prepared for ${process.argv[2]}`);
}
