import ExcelJS from 'exceljs';
import {Chart,registerables} from 'chart.js';
import {unzipSync,strFromU8} from 'fflate';
import {columnName,cellValue,resolvePart} from './core.mjs';
Chart.register(...registerables);
const palette=['#3b77bd','#60a378','#dd9b52','#9b80bc','#cb6d6d'];
const number=new Intl.NumberFormat('ru-RU',{maximumFractionDigits:6});
const rub=new Intl.NumberFormat('ru-RU',{style:'currency',currency:'RUB',maximumFractionDigits:0});
function display(cell:any){
 const v=cellValue(cell.value),fmt=cell.numFmt||'';
 if(v===null)return 'Нет значения';
 if(v instanceof Date)return /yyyy-mm(?!-)/.test(fmt)?v.toISOString().slice(0,7):new Intl.DateTimeFormat('ru-RU').format(v);
 if(typeof v==='number'){
  if(fmt.includes('%'))return new Intl.NumberFormat('ru-RU',{style:'percent',maximumFractionDigits:(fmt.match(/\.(0+)/)?.[1]?.length??0)}).format(v);
  if(/[₽]|руб|RUB/i.test(fmt))return rub.format(v);
  const decimals=fmt.match(/\.(0+)/)?.[1]?.length;
  if(decimals!==undefined)return new Intl.NumberFormat('ru-RU',{minimumFractionDigits:decimals,maximumFractionDigits:decimals}).format(v);
  return number.format(v);
 }
 return String(v??'');
}
function color(c:any){const argb=c?.argb;return typeof argb==='string'?`#${argb.slice(-6)}`:undefined;}
function doc(archive:any,path:string){const bytes=archive[path];return bytes?new DOMParser().parseFromString(strFromU8(bytes),'application/xml'):null;}
function all(node:any,local:string):Element[]{return node?Array.from(node.getElementsByTagName('*')).filter((e:any)=>e.localName===local) as Element[]:[];}
function child(node:any,local:string){return Array.from(node?.children||[]).find((x:any)=>x.localName===local) as Element|undefined;}
function text(node:any,local:string){return all(node,local)[0]?.textContent||'';}
function rels(archive:any,path:string){
 const slash=path.lastIndexOf('/'),p=`${path.slice(0,slash+1)}_rels/${path.slice(slash+1)}.rels`;
 return new Map(all(doc(archive,p),'Relationship').filter(x=>x.getAttribute('TargetMode')!=='External').map(x=>[x.getAttribute('Id'),resolvePart(path,x.getAttribute('Target'))]));
}
function cache(node:any){
 const points=all(node,'pt').map(pt=>({i:Number(pt.getAttribute('idx')),v:text(pt,'v')})).sort((a,b)=>a.i-b.i);
 return points.map(p=>p.v);
}
function sheetCharts(archive:any,sheetName:string){
 const wb=doc(archive,'xl/workbook.xml'),sheet=all(wb,'sheet').find(x=>x.getAttribute('name')===sheetName);
 if(!sheet)return [];
 const relation=sheet.getAttribute('r:id'),sheetPath=rels(archive,'xl/workbook.xml').get(relation);
 if(!sheetPath)return [];
 const sheetXml=doc(archive,sheetPath),drawing=all(sheetXml,'drawing')[0];
 const drawingPath=drawing&&rels(archive,sheetPath).get(drawing.getAttribute('r:id'));
 if(!drawingPath)return [];
 const drawingXml=doc(archive,drawingPath),drawingRels=rels(archive,drawingPath);
 return all(drawingXml,'chart').map(x=>drawingRels.get(x.getAttribute('r:id'))).filter(Boolean).map(path=>doc(archive,path as string)).filter(Boolean);
}
export async function renderExcel(bytes:Uint8Array,area:HTMLElement,tabs:HTMLElement,onSelect:(address:string,value:string)=>void,onStatus:(text:string)=>void){
 const wb=new ExcelJS.Workbook();await wb.xlsx.load(bytes as any);
 const archive=unzipSync(bytes),charts:Chart[]=[];
 let selected=0,rows=300;
 const draw=()=>{
  for(const c of charts)c.destroy();charts.length=0;area.replaceChildren();tabs.replaceChildren();
  const sheets=wb.worksheets.filter(s=>s.state!=='veryHidden');
  if(!sheets.length){area.textContent='В книге нет листов.';return;}
  selected=Math.min(selected,sheets.length-1);const sheet=sheets[selected];
  sheets.forEach((s,i)=>{const button=document.createElement('button');button.textContent=s.name;button.setAttribute('role','tab');button.setAttribute('aria-selected',String(i===selected));button.onclick=()=>{selected=i;rows=300;draw();};tabs.append(button);});
  const lastRow=Math.min(Math.max(sheet.rowCount,1),rows),lastCol=Math.min(Math.max(sheet.columnCount,1),80);
  const layout=document.createElement('div');layout.className='excel-layout';
  const wrapper=document.createElement('div');wrapper.className='grid-wrap';
  const table=document.createElement('table');table.className='grid';table.setAttribute('aria-label',`Лист ${sheet.name}`);
  const group=document.createElement('colgroup');let col=document.createElement('col');col.style.width='39px';group.append(col);
  for(let i=1;i<=lastCol;i++){col=document.createElement('col');col.style.width=`${Math.min(360,Math.max(70,(sheet.getColumn(i).width||13)*7+12))}px`;group.append(col);}
  table.append(group);
  const head=document.createElement('thead'),header=document.createElement('tr'),corner=document.createElement('th');corner.className='corner';header.append(corner);
  for(let c=1;c<=lastCol;c++){const th=document.createElement('th');th.textContent=columnName(c);header.append(th);}head.append(header);table.append(head);
  const merges=new Map<string,{r1:number,c1:number,r2:number,c2:number}>();
  for(const value of (sheet.model.merges||[])){const match=value.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);if(!match)continue;const cn=(x:string)=>Array.from(x).reduce((a,b)=>a*26+b.charCodeAt(0)-64,0);merges.set(`${match[1]}${match[2]}`,{r1:+match[2],c1:cn(match[1]),r2:+match[4],c2:cn(match[3])});}
  const body=document.createElement('tbody');let missingCache=0;
  for(let r=1;r<=lastRow;r++){
   if(sheet.getRow(r).hidden)continue;
   const tr=document.createElement('tr'),rowhead=document.createElement('th');rowhead.textContent=String(r);rowhead.scope='row';rowhead.className='row-head';tr.append(rowhead);
   tr.style.height=`${Math.max(24,(sheet.getRow(r).height||18)*1.333)}px`;
   for(let c=1;c<=lastCol;c++){
    const cell=sheet.getCell(r,c);if(sheet.getColumn(c).hidden)continue;if(cell.isMerged&&cell.master.address!==cell.address)continue;
    const td=document.createElement('td'),merge=merges.get(cell.address);if(merge){td.rowSpan=Math.min(merge.r2,lastRow)-r+1;td.colSpan=Math.min(merge.c2,lastCol)-c+1;}
    const value=cellValue(cell.value);td.textContent=display(cell);td.title=td.textContent||'';td.dataset.address=cell.address;td.tabIndex=0;
    if(value===null&&cell.type===ExcelJS.ValueType.Formula){missingCache++;td.classList.add('missing-cache');}
    td.style.textAlign=cell.alignment?.horizontal||(typeof value==='number'?'right':'left');
    if(cell.alignment?.wrapText){td.style.whiteSpace='normal';td.style.overflowWrap='anywhere';}
    if(cell.font?.bold)td.style.fontWeight='700';if(cell.font?.italic)td.style.fontStyle='italic';
    if(cell.font?.size)td.style.fontSize=`${Math.min(36,Math.max(9,cell.font.size))*1.2}px`;
    if(cell.font?.name)td.style.fontFamily=`"${cell.font.name.replace(/["<>]/g,'')}",Calibri,Arial,sans-serif`;
    td.style.color=color(cell.font?.color)||'#20252b';
    if(cell.fill?.type==='pattern'&&cell.fill.pattern==='solid')td.style.background=color(cell.fill.fgColor)||'#fff';
    for(const side of ['top','right','bottom','left']){const border=(cell.border as any)?.[side];if(border?.style)(td.style as any)[`border${side[0].toUpperCase()+side.slice(1)}`]=`${/medium|thick/i.test(border.style)?2:1}px solid ${color(border.color)||'#adb5bc'}`;}
    const select=()=>{table.querySelector('.selected')?.classList.remove('selected');td.classList.add('selected');onSelect(cell.address,cell.formula?`=${cell.formula}`:display(cell));};
    td.onclick=select;td.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();select();}};
    tr.append(td);
   }body.append(tr);
  }
  table.append(body);wrapper.append(table);
  if(sheet.rowCount>lastRow){const button=document.createElement('button');button.className='load-more';button.textContent=`Ещё строки (${lastRow} из ${sheet.rowCount})`;button.onclick=()=>{rows=Math.min(rows+300,2000);draw();};button.disabled=rows>=2000;wrapper.append(button);}
  layout.append(wrapper);
  const column=document.createElement('div');column.className='charts';
  for(const chartDoc of sheetCharts(archive,sheet.name)){
   const plot=all(chartDoc,'plotArea')[0];if(!plot)continue;
   const plotType=Array.from(plot.children).find((x:any)=>/^(barChart|lineChart|pieChart|doughnutChart|areaChart)$/.test(x.localName));
   if(!plotType)continue;
   const type=plotType.localName,series=all(plotType,'ser');
   const labels=cache(child(series[0],'cat'));
   const datasets=series.map((s,i)=>({label:text(child(s,'tx'),'v')||`Ряд ${i+1}`,data:cache(child(s,'val')).map(Number),backgroundColor:/pie|doughnut/.test(type)?palette:palette[i%palette.length],borderColor:palette[i%palette.length],borderWidth:1.5,fill:type==='areaChart'}));
   if(!datasets.length)continue;
   const card=document.createElement('section');card.className='chart-card';const title=document.createElement('h3');title.textContent=all(all(chartDoc,'title')[0],'t').map(x=>x.textContent).join('')||'Диаграмма';
   const holder=document.createElement('div');holder.className='chart-canvas';const canvas=document.createElement('canvas');canvas.setAttribute('aria-label',title.textContent);holder.append(canvas);card.append(title,holder);column.append(card);
   const kind=type==='barChart'?'bar':type==='pieChart'?'pie':type==='doughnutChart'?'doughnut':'line';
   const chart=new Chart(canvas,{type:kind,data:{labels:labels.length?labels:datasets[0].data.map((_,i)=>String(i+1)),datasets},options:{responsive:true,maintainAspectRatio:false,animation:false,plugins:{legend:{position:'bottom',labels:{boxWidth:11,font:{family:'Segoe UI',size:11}}}},...(!/pie|doughnut/.test(type)?{scales:{y:{beginAtZero:true,ticks:{callback:(v:any)=>number.format(Number(v))},grid:{color:'#edf0ec'}},x:{grid:{display:false}}}}:{})}} as any);
   charts.push(chart);
  }
  if(column.children.length)layout.append(column);
  area.append(layout);onSelect('','Выберите ячейку, чтобы увидеть значение или формулу');
  onStatus(`${sheet.name} · ${sheet.rowCount} строк × ${sheet.columnCount} столбцов${missingCache?' · Есть формулы без сохранённых значений':''}`);
 };
 draw();return()=>charts.forEach(c=>c.destroy());
}
