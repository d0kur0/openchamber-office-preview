import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {unzipSync,strFromU8} from 'fflate';
import {packageRelease} from '../scripts/package-release.mjs';

test('release ZIP has matching version, only installable files, and correct checksums',async()=>{
  const root=await fs.mkdtemp(path.join(os.tmpdir(),'office-release-test-'));
  try{
    for(const dir of ['panel','licenses','dist'])await fs.mkdir(path.join(root,dir));
    await fs.writeFile(path.join(root,'package.json'),JSON.stringify({name:'openchamber-office-preview',version:'1.2.3'}));
    for(const name of ['README.md','LICENSE','THIRD_PARTY_NOTICES.md','panel/index.html','panel/main.js','panel/style.css','licenses/test.txt'])await fs.writeFile(path.join(root,name),'fixture');
    await fs.writeFile(path.join(root,'.npmrc'),'must-not-ship');
    await fs.writeFile(path.join(root,'dist/openchamber-office-preview-1.2.3.tgz'),'tgz fixture');
    const result=await packageRelease(root,'1.2.3');
    const archive=unzipSync(await fs.readFile(path.join(root,'dist',result.zipName)));
    assert.equal(JSON.parse(strFromU8(archive['office-preview/package.json'])).version,'1.2.3');
    assert(!Object.keys(archive).some(x=>x.includes('.npmrc')||x.includes('dist/')));
    assert(archive['office-preview/panel/index.html']);
    for(const line of (await fs.readFile(path.join(root,'dist/SHA256SUMS.txt'),'utf8')).trim().split('\n')){
      const [hash,name]=line.split('  ');
      assert.equal(createHash('sha256').update(await fs.readFile(path.join(root,'dist',name))).digest('hex'),hash);
    }
    await assert.rejects(packageRelease(root,'1.2.4'),/Prepared package version/);
  }finally{
    assert.equal(path.dirname(path.resolve(root)),path.resolve(os.tmpdir()));
    assert(path.basename(root).startsWith('office-release-test-'));
    await fs.rm(root,{recursive:true,force:true});
  }
});
