import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeCommits} from '@semantic-release/commit-analyzer';
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
