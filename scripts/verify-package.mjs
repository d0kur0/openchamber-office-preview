import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
const p=JSON.parse(await fs.readFile('package.json','utf8'));
assert.equal(p.name,'openchamber-office-preview');
assert.notEqual(p.private,true);
assert.equal(p.openchamber.apiVersion,1);
assert.equal(p.openchamber.contributes.panel.id,'office-preview');
assert.equal(p.openchamber.contributes.fileEditors[0].content,'binary');
assert.equal(p.dependencies,undefined,'The viewer must ship without runtime npm dependencies');
for(const file of ['panel/index.html','panel/main.js','panel/style.css','LICENSE','THIRD_PARTY_NOTICES.md']){
  assert((await fs.stat(file)).size>0,`${file} must be built`);
}
if(process.env.GITHUB_REF_TYPE==='tag'){
  assert.equal(process.env.GITHUB_REF_NAME,`v${p.version}`,'Tag and package version must match');
}
console.log(`Package verified: ${p.name}@${p.version}`);
