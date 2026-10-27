export const LIMITS = { file:20*1024*1024, expanded:100*1024*1024, entries:4000 };

export function checkArchive(bytes) {
  if(bytes.length>LIMITS.file) throw new Error('Файл больше 20 МБ. Откройте его в офисном приложении.');
  const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
  if(bytes.length<22||view.getUint32(0,true)!==0x04034b50) throw new Error('Это не файл Office Open XML. Старые DOC, XLS и PPT не поддерживаются.');
  let end=-1;
  for(let p=bytes.length-22;p>=Math.max(0,bytes.length-65557);p--)if(view.getUint32(p,true)===0x06054b50){end=p;break;}
  if(end<0)throw new Error('Повреждённый архив Office: не найден каталог.');
  const count=view.getUint16(end+10,true),offset=view.getUint32(end+16,true);
  if(count===65535||offset===0xffffffff)throw new Error('Архив ZIP64 не поддерживается.');
  if(count>LIMITS.entries)throw new Error('В документе слишком много вложенных объектов.');
  let total=0,p=offset;
  for(let i=0;i<count;i++){
    if(p+46>bytes.length||view.getUint32(p,true)!==0x02014b50)throw new Error('Повреждённый каталог документа.');
    if(view.getUint16(p+8,true)&1)throw new Error('Файл защищён паролем. Откройте незашифрованную копию.');
    total+=view.getUint32(p+24,true);
    if(total>LIMITS.expanded)throw new Error('Слишком большой объём распакованных данных (больше 100 МБ).');
    p+=46+view.getUint16(p+28,true)+view.getUint16(p+30,true)+view.getUint16(p+32,true);
  }
  return {entries:count,expanded:total};
}

export function columnName(n){let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;}
export function resolvePart(base,target){
  if(!target||/^[a-z]+:/i.test(target)||target.startsWith('//'))return null;
  const parts=target.startsWith('/')?[]:base.split('/').slice(0,-1);
  for(const part of target.split('/')){if(part==='..'){if(!parts.length)return null;parts.pop();}else if(part&&part!=='.')parts.push(part);}
  return parts.join('/');
}
export function cellValue(value){
  if(value===undefined||value===null)return '';
  if(value instanceof Date)return value;
  if(typeof value==='object'){
    if('formula'in value||'sharedFormula'in value)return 'result'in value?value.result:null;
    if(value.richText)return value.richText.map(x=>x.text).join('');
    if(value.text!==undefined)return value.text;
    if(value.error)return value.error;
  }
  return value;
}
export function safeFilename(name){return String(name||'Документ').split(/[\\/]/).pop();}

// Fit is explicit; automatic spreadsheet previews prioritize readable text.
export function previewScale(mode,kind,available,natural){
  const fit=Math.max(.1,Math.min(1,Math.max(1,available)/Math.max(1,natural)));
  if(mode==='fit')return fit;
  if(mode==='auto')return kind==='excel'?1.6:Math.max(kind==='slides'?.6:.85,fit);
  const value=Number(mode);
  return Number.isFinite(value)&&value>0?Math.max(.25,Math.min(4,value)):1;
}
