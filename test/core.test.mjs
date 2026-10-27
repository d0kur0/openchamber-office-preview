import test from 'node:test';
import assert from 'node:assert/strict';
import {zipSync,strToU8} from 'fflate';
import {checkArchive,columnName,cellValue,resolvePart,LIMITS} from '../src/core.mjs';
test('rejects non-Office bytes and oversized expanded ZIP before unzip',()=>{
 assert.throws(()=>checkArchive(new Uint8Array(22)),/Office Open XML/);
 const zip=zipSync({'document.xml':strToU8('<test/>')});
 assert.equal(checkArchive(zip).entries,1);
 const bad=zip.slice(),v=new DataView(bad.buffer);let found=false;
 for(let i=0;i<bad.length-46;i++)if(v.getUint32(i,true)===0x02014b50){v.setUint32(i+24,LIMITS.expanded+1,true);found=true;break;}
 assert.ok(found);assert.throws(()=>checkArchive(bad),/100 МБ/);
});
test('external relationships are not resolved as document parts',()=>{
 assert.equal(resolvePart('xl/drawings/drawing1.xml','../charts/chart1.xml'),'xl/charts/chart1.xml');
 assert.equal(resolvePart('xl/workbook.xml','https://outside.invalid/a.xml'),null);
 assert.equal(resolvePart('xl/workbook.xml','../../outside.xml'),null);
});
test('formula preview uses cache and distinguishes missing cache from zero',()=>{
 assert.equal(cellValue({formula:'SUM(A1:A2)',result:0}),0);
 assert.equal(cellValue({formula:'SUM(A1:A2)'}),null);
 assert.equal(cellValue({richText:[{text:'Русский '},{text:'English'}]}),'Русский English');
 assert.equal(columnName(27),'AA');
});
