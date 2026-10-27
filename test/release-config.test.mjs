import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeCommits} from '@semantic-release/commit-analyzer';
import {generateNotes} from '@semantic-release/release-notes-generator';
import config from '../release.config.mjs';

test('release policy maps conventional commits to expected versions',async()=>{
  const analyzer=config.plugins.find(x=>Array.isArray(x)&&x[0]==='@semantic-release/commit-analyzer')[1];
  for(const [message,expected] of [
    ['fix: correct spreadsheet rendering','patch'],
    ['feat: support another document format','minor'],
    ['feat!: change extension API','major'],
    ['ci(release): update publishing pipeline','patch'],
    ['docs: explain installation',null],
  ]){
    const result=await analyzeCommits(analyzer,{cwd:process.cwd(),env:process.env,commits:[{hash:'test',message}],logger:{log(){}}});
    assert.equal(result,expected,message);
  }
});

test('release notes preset renders with the installed changelog writer',async()=>{
  const preset=config.plugins.find(x=>Array.isArray(x)&&x[0]==='@semantic-release/release-notes-generator')[1];
  const notes=await generateNotes(preset,{
    cwd:process.cwd(),env:process.env,
    commits:[{hash:'a'.repeat(40),message:'fix: correct spreadsheet rendering'}],
    lastRelease:{version:'0.1.0',gitTag:'v0.1.0'},
    nextRelease:{version:'0.1.1',gitTag:'v0.1.1'},
    options:{repositoryUrl:'https://github.com/d0kur0/openchamber-office-preview.git'},
    logger:{log(){}},
  });
  assert.match(notes,/0\.1\.1/);
  assert.match(notes,/correct spreadsheet rendering/);
});
