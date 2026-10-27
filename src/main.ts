import {connectHost} from '@openchamber/sdk';
import {applyHostReady} from '@openchamber/sdk/ui';
import {renderAsync} from 'docx-preview';
import {loadPresentation,getSlides} from '@office-kit/pptx';
import {renderSlideToSvg} from '@office-kit/pptx-preview';
import DOMPurify from 'dompurify';
import {renderExcel} from './excel';
import {checkArchive,safeFilename} from './core.mjs';

const host=connectHost();
const root=document.querySelector('#app')!;
root.innerHTML=`<main class="app"><header class="toolbar"><div id="kind" class="file-icon" aria-hidden="true">OF</div><div class="file-meta"><div class="filename" id="filename">Office Preview</div><div class="file-subtitle" id="subtitle">Excel · Word · PowerPoint</div></div><div class="slide-nav" id="slide-nav" hidden><button id="previous" aria-label="Предыдущий слайд">←</button><span id="slide-count" class="slide-count"></span><button id="next" aria-label="Следующий слайд">→</button></div><div class="zoom"><button id="zoom-out" aria-label="Уменьшить масштаб">−</button><select id="zoom" aria-label="Масштаб"><option value="fit">По ширине</option><option value="0.5">50%</option><option value="0.75">75%</option><option value="1">100%</option><option value="1.25">125%</option><option value="1.5">150%</option><option value="1.6">160%</option></select><button id="zoom-in" aria-label="Увеличить масштаб">+</button></div><button id="choose" class="open">Открыть файл</button><input id="file-input" class="sr-only" type="file" accept=".xlsx,.xlsm,.docx,.pptx" aria-label="Выбрать офисный файл"></header><div id="formula" class="formula-bar" hidden><span id="address" class="address"></span><span class="formula-mark">fx</span><span id="formula-value" class="formula-value"></span></div><section id="viewport" class="viewport" aria-label="Просмотр документа"><div id="content" class="content"></div></section><nav id="tabs" class="sheet-tabs" role="tablist" aria-label="Листы книги" hidden></nav><footer class="footer"><span id="status" role="status">Готов к открытию файла</span><span id="mode">Только просмотр · локально</span></footer></main>`;
const $=<T extends HTMLElement>(id:string)=>document.getElementById(id) as T;
const content=$('content'),viewport=$('viewport'),tabs=$('tabs');
let original:Uint8Array|null=null,version='',generation=0,dispose:(()=>void)|null=null,scale=1;
let slideIndex=0,slides:any[]=[],presentation:any=null;
let filename='';
const cleanSvg=(s:string)=>DOMPurify.sanitize(s,{USE_PROFILES:{svg:true,svgFilters:true},ADD_TAGS:['image'],ADD_ATTR:['href','xlink:href']});
function message(title:string,body:string,error=false){
 content.replaceChildren();const box=document.createElement('div');box.className='empty';
 const h=document.createElement('h2');h.textContent=title;if(error)h.className='error-title';
 const p=document.createElement('p');p.textContent=body;box.append(h,p);content.append(box);
}
function empty(){
 message('Документ рядом с чатом','Откройте XLSX, DOCX или PPTX в Files, выберите файл кнопкой сверху или перетащите его сюда.');
 const icons=document.createElement('div');icons.className='format-icons';
 for(const name of ['Excel','Word','PowerPoint']){const e=document.createElement('span');e.className='format-pill';e.textContent=name;icons.append(e);}content.firstElementChild?.prepend(icons);
}
function zoom(){
 const setting=($<HTMLSelectElement>('zoom')).value;
 // CSS zoom keeps scroll geometry consistent for large pages and spreadsheets.
 content.style.zoom='1';content.style.width='max-content';content.style.minWidth='100%';
 const natural=Math.max(content.scrollWidth,content.getBoundingClientRect().width);
 scale=setting==='fit'?Math.min(1,Math.max(.25,(viewport.clientWidth-36)/Math.max(natural,1))):Number(setting);
 content.style.zoom=String(scale);content.style.minWidth=`${100/scale}%`;
}
let resizeTimer:any;new ResizeObserver(()=>{clearTimeout(resizeTimer);resizeTimer=setTimeout(zoom,60);}).observe(viewport);
$<HTMLSelectElement>('zoom').onchange=zoom;
function changeScale(delta:number){
 const select=$<HTMLSelectElement>('zoom');const next=Math.min(1.6,Math.max(.5,Math.round((scale+delta)*100)/100));
 let option=Array.from(select.options).find(x=>Number(x.value)===next);
 if(!option){option=new Option(`${Math.round(next*100)}%`,String(next));select.add(option);}
 select.value=String(next);zoom();
}
$('zoom-in').onclick=()=>changeScale(.1);$('zoom-out').onclick=()=>changeScale(-.1);
const picker=$<HTMLInputElement>('file-input');$('choose').onclick=()=>picker.click();
picker.onchange=async()=>{const file=picker.files?.[0];if(file)await open(file.name,new Uint8Array(await file.arrayBuffer()),false);picker.value='';};
viewport.ondragover=e=>{e.preventDefault();};viewport.ondrop=async e=>{e.preventDefault();const file=e.dataTransfer?.files?.[0];if(file)await open(file.name,new Uint8Array(await file.arrayBuffer()),false);};

function paintSlide(){
 const stage=content.querySelector('.slide-stage');if(!stage||!presentation)return;
 stage.innerHTML=cleanSvg(renderSlideToSvg(presentation,slides[slideIndex],{textLayout:'svg'}));
 const svg=stage.querySelector('svg');const box=svg?.viewBox.baseVal;if(svg&&box?.width&&box?.height){svg.setAttribute('width',String(box.width));svg.setAttribute('height',String(box.height));}
 for(const [i,button] of Array.from(content.querySelectorAll('.slide-thumb')).entries())button.setAttribute('aria-current',String(i===slideIndex));
 $('slide-count').textContent=`${slideIndex+1} / ${slides.length}`;
 $<HTMLButtonElement>('previous').disabled=slideIndex===0;$<HTMLButtonElement>('next').disabled=slideIndex===slides.length-1;
 $('status').textContent=`Слайд ${slideIndex+1} из ${slides.length}`;zoom();
}
$('previous').onclick=()=>{if(slideIndex>0){slideIndex--;paintSlide();}};
$('next').onclick=()=>{if(slideIndex<slides.length-1){slideIndex++;paintSlide();}};
document.addEventListener('keydown',e=>{if(presentation&&e.target===document.body){if(e.key==='ArrowRight')$('next').click();if(e.key==='ArrowLeft')$('previous').click();}});

async function open(name:string,bytes:Uint8Array,fromHost:boolean){
 const token=++generation;dispose?.();dispose=null;original=bytes;filename=safeFilename(name);version=`${bytes.length}-${Date.now()}`;
 presentation=null;slides=[];slideIndex=0;tabs.replaceChildren();tabs.hidden=true;$('formula').hidden=true;$('slide-nav').hidden=true;
 $('filename').textContent=filename;$('status').textContent='Открываем документ…';viewport.className='viewport';
 $('mode').textContent='Только просмотр · локально';$('choose').hidden=fromHost;
 content.style.zoom='1';content.replaceChildren();
 const loading=document.createElement('div');loading.className='empty';loading.innerHTML='<div class="spinner" aria-hidden="true"></div><p>Подготавливаем страницы…</p>';content.append(loading);
 try{
  checkArchive(bytes);
  const ext=filename.split('.').pop()?.toLowerCase();
  const kinds:any={xlsx:['XL','Excel',''],xlsm:['XL','Excel',''],docx:['W','Word','word'],pptx:['P','PowerPoint','slides']};
  if(!kinds[ext!])throw new Error('Поддерживаются XLSX, XLSM, DOCX и PPTX. PDF открывается штатным просмотрщиком OpenChamber.');
  const [symbol,label,cls]=kinds[ext!];$('kind').textContent=symbol;$('kind').className=`file-icon ${cls}`;
  $('subtitle').textContent=`${label} · ${(bytes.length/1024).toFixed(0)} КБ · только чтение`;
  if(ext==='xlsx'||ext==='xlsm'){
   viewport.classList.add('excel');tabs.hidden=false;$('formula').hidden=false;$('mode').textContent='Сохранённые значения формул';
   const page=document.createElement('div');
   const cleanup=await renderExcel(bytes,page,tabs,(address,value)=>{$('address').textContent=address;$('formula-value').textContent=value;},text=>{$('status').textContent=text;queueMicrotask(zoom);});
   if(token!==generation){cleanup();return;}content.replaceChildren(page);dispose=cleanup;
  }else if(ext==='docx'){
   const page=document.createElement('div');page.className='docx-content';
   await renderAsync(bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength),page,undefined,{inWrapper:true,breakPages:true,ignoreWidth:false,ignoreHeight:false,renderHeaders:true,renderFooters:true,renderFootnotes:true,useBase64URL:true,renderAltChunks:false});
   if(token!==generation)return;
   DOMPurify.sanitize(page,{IN_PLACE:true,ADD_TAGS:['style'],FORBID_TAGS:['script','iframe','object','embed']});
   // Documents are previews, so hyperlinks do not navigate the host or run actions.
   for(const a of Array.from(page.querySelectorAll('a'))){a.removeAttribute('href');a.removeAttribute('target');}
   for(const el of Array.from(page.querySelectorAll('[src]'))){const src=el.getAttribute('src')||'';if(!/^(data:|blob:)/.test(src))el.removeAttribute('src');}
   content.replaceChildren(page);$('status').textContent=`${page.querySelectorAll('section.docx').length||1} стр. · Word`;
  }else{
   const pres=await loadPresentation(bytes);if(token!==generation)return;presentation=pres;slides=getSlides(pres);
   if(!slides.length)throw new Error('В презентации нет слайдов.');
   const layout=document.createElement('div');layout.className='slide-layout';const rail=document.createElement('nav');rail.className='slide-rail';rail.setAttribute('aria-label','Слайды');
   for(let i=0;i<Math.min(slides.length,100);i++){
    const button=document.createElement('button');button.className='slide-thumb';button.setAttribute('aria-label',`Слайд ${i+1}`);
    if(i<20)button.innerHTML=cleanSvg(renderSlideToSvg(pres,slides[i],{textLayout:'svg'}));const text=document.createElement('span');text.textContent=`Слайд ${i+1}`;button.append(text);button.onclick=()=>{slideIndex=i;paintSlide();};rail.append(button);
   }
   const stage=document.createElement('div');stage.className='slide-stage';layout.append(rail,stage);content.replaceChildren(layout);$('slide-nav').hidden=false;paintSlide();
  }
  if(fromHost)host.reportFileChange({dirty:false,edited:false});
  zoom();setTimeout(zoom,100);
 }catch(error){
  if(token!==generation)return;tabs.hidden=true;$('formula').hidden=true;$('slide-nav').hidden=true;
  const reason=error instanceof Error?error.message:String(error);
  message('Не удалось открыть документ',reason,false);content.querySelector('h2')?.classList.add('error-title');
  $('status').textContent='Исходный файл не изменён';$('choose').hidden=false;
 }
}
host.onReady(ctx=>{applyHostReady(ctx,document.documentElement);document.body.dataset.theme=ctx.theme.mode;});
host.onFileOpen(file=>{if(file.encoding==='binary')void open(file.name,new Uint8Array(file.bytes),true);else message('Нужен офисный файл','Откройте XLSX, DOCX или PPTX.');});
host.onFileSnapshot(()=>{if(!original)throw new Error('No document open');return{bytes:original,version};});
window.addEventListener('pagehide',()=>{dispose?.();host.dispose();});
empty();

