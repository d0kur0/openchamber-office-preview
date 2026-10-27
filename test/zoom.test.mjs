import test from 'node:test';
import assert from 'node:assert/strict';
import {previewScale} from '../src/core.mjs';

test('automatic spreadsheet preview stays readable for wide sheets and narrow panes',()=>{
 for(const width of [320,640,1920])assert.equal(previewScale('auto','excel',width,3200),1.6);
 assert.equal(previewScale('fit','excel',800,1600),.5);
});
test('pages and slides have a readability floor, with literal fit available separately',()=>{
 assert.equal(previewScale('auto','word',320,800),.85);
 assert.equal(previewScale('auto','slides',320,1200),.6);
 assert.equal(previewScale('fit','word',320,800),.4);
 assert.equal(previewScale('auto','word',1600,800),1);
});
test('manual zoom is bounded and invalid values cannot break scroll geometry',()=>{
 assert.equal(previewScale('4','excel',320,1000),4);
 assert.equal(previewScale('100','excel',320,1000),4);
 assert.equal(previewScale('.01','excel',320,1000),.25);
 assert.equal(previewScale('broken','excel',320,1000),1);
 assert.ok(Number.isFinite(previewScale('fit','excel',0,0)));
});
