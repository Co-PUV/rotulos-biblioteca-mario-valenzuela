
const workerSource = `
importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
const clean=v=>String(v??'').trim().toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').replace(/[^a-z0-9]/g,'');
const digits=v=>String(v??'').replace(/[^0-9]/g,'');
const codeHeader=v=>{const x=clean(v);return x==='column16357'||x.includes('codigodebarra')||x.includes('codigobarra')||x.includes('barcode')||x.includes('itembarcode')};
const classHeader=v=>{const x=clean(v);return x.includes('clasificacion')||x.includes('clasificaicon')||x.includes('signatura')||x.includes('callnumber')};
const excluded=n=>{const x=clean(n);return x.includes('encuadernacion')||x.includes('inventarionumerico')||x.includes('cuantitavio')||x.includes('cuantitativo')};
function value(ws,r,c){const z=ws[XLSX.utils.encode_cell({r,c})];return z?String(z.w??z.v??'').trim():''}
function sparse(ws,maxRow){const rows=new Map(),cols=new Set();for(const k in ws){if(k[0]==='!')continue;const p=XLSX.utils.decode_cell(k);if(p.r<=maxRow){cols.add(p.c);if(!rows.has(p.r))rows.set(p.r,[]);rows.get(p.r).push(p.c)}}return{rows,cols:[...cols]}}
function locate(ws){
 if(!ws||!ws['!ref'])return null;
 const rg=XLSX.utils.decode_range(ws['!ref']), max=Math.min(rg.e.r,rg.s.r+300), sp=sparse(ws,max);
 for(let r=rg.s.r;r<=max;r++){let b=-1,c=-1;for(const col of sp.rows.get(r)||[]){const v=value(ws,r,col);if(b<0&&codeHeader(v))b=col;if(c<0&&classHeader(v))c=col}if(b>=0&&c>=0)return{r,b,c,last:rg.e.r}}
 return null;
}
function infer(ws){
 if(!ws||!ws['!ref'])return null;
 const rg=XLSX.utils.decode_range(ws['!ref']),max=Math.min(rg.e.r,rg.s.r+300),sp=sparse(ws,max);
 let c=-1,rh=-1;
 for(let r=rg.s.r;r<=max&&c<0;r++)for(const col of sp.rows.get(r)||[])if(classHeader(value(ws,r,col))){c=col;rh=r;break}
 if(c<0)return null;
 let best=-1,score=0;
 for(const col of sp.cols){if(col===c)continue;let n=0;for(let r=rh+1;r<=Math.min(rg.e.r,rh+600);r++){const d=digits(value(ws,r,col));if(d.length>=8&&d.length<=18)n++}if(n>score){score=n;best=col}}
 return best>=0&&score>0?{r:rh,b:best,c,last:rg.e.r}:null;
}
function extract(ws,x){const out=[];for(let r=x.r+1;r<=x.last;r++){const code=value(ws,r,x.b),key=digits(code),cl=value(ws,r,x.c);if(key.length>=8&&cl)out.push([key,{code,classification:cl}])}return out}
self.onmessage=e=>{try{
 const wb=XLSX.read(e.data.buffer,{type:'array',raw:false,cellText:true}),mode=e.data.mode,out=[];
 if(mode==='nuevos'){
  let n=wb.SheetNames.find(x=>clean(x)==='nuevosinventario2026')||wb.SheetNames.find(x=>{const y=clean(x);return y.includes('nuevos')&&y.includes('inventario')&&y.includes('2026')});
  if(!n&&wb.SheetNames.length===1)n=wb.SheetNames[0]; if(!n)throw Error('No se encontró la hoja de Nuevos.');
  const ws=wb.Sheets[n],x=locate(ws)||infer(ws); if(!x)throw Error('No se encontraron las columnas de código y clasificación en Nuevos.');
  for(const [k,v] of extract(ws,x))out.push([k,{...v,sheet:n}]);
 }else{
  for(const n of wb.SheetNames){if(excluded(n))continue;const ws=wb.Sheets[n],x=locate(ws)||infer(ws);if(!x)continue;for(const [k,v] of extract(ws,x))out.push([k,{...v,sheet:n}])}
  if(!out.length)throw Error('No se encontraron registros válidos de código de barras y clasificación en el Inventario.');
 }
 self.postMessage({ok:true,out});
}catch(err){self.postMessage({ok:false,error:err.message||String(err)})}}`;
const workerURL=URL.createObjectURL(new Blob([workerSource],{type:'text/javascript'}));

const states={nuevos:{db:new Map(),labels:[]},inventario:{db:new Map(),labels:[]}};

function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]))}
function status(mode,msg,type=''){
 const e=document.getElementById('status-'+mode);e.textContent=msg;e.className='status show '+type;
}
function render(mode){
 const st=states[mode], labels=document.getElementById('labels-'+mode), tbody=document.getElementById('table-'+mode);
 document.getElementById('count-'+mode).textContent=st.labels.length;
 labels.innerHTML=st.labels.length?st.labels.map(x=>'<div class="label">'+esc(x.classification).replace(/\s+/g,'<br>')+'</div>').join(''):'<div class="empty">Aún no hay rótulos seleccionados.</div>';
 tbody.innerHTML=st.labels.map((x,i)=>`<tr><td>${i+1}</td><td>${esc(x.code)}</td><td>${esc(x.classification)}</td><td><button type="button" class="secondary" data-remove="${mode}:${i}">Eliminar</button></td></tr>`).join('');
 tbody.querySelectorAll('[data-remove]').forEach(b=>b.onclick=()=>{const [m,i]=b.dataset.remove.split(':');states[m].labels.splice(+i,1);render(m)});
}
async function load(mode,file){
 if(!file)return;
 const p=document.getElementById('progress-'+mode),bar=p.querySelector('span');p.style.display='block';bar.style.width='15%';
 status(mode,'Leyendo el archivo...');
 try{
  const buffer=await file.arrayBuffer();bar.style.width='45%';
  const records=await new Promise((resolve,reject)=>{
   const w=new Worker(workerURL),t=setTimeout(()=>{w.terminate();reject(Error('La lectura tardó demasiado.'))},120000);
   w.onmessage=e=>{clearTimeout(t);w.terminate();e.data.ok?resolve(e.data.out):reject(Error(e.data.error))};
   w.onerror=e=>{clearTimeout(t);w.terminate();reject(Error('No fue posible procesar el archivo.'))};
   w.postMessage({buffer,mode},[buffer]);
  });
  bar.style.width='85%';
  const db=new Map();records.forEach(([k,v])=>db.set(k,v));
  states[mode].db=db;bar.style.width='100%';
  status(mode,`Archivo cargado correctamente. Se encontraron ${db.size} códigos disponibles.`,'ok');
 }catch(err){status(mode,'No se pudo cargar el archivo: '+err.message,'error')}
 finally{setTimeout(()=>{p.style.display='none';bar.style.width='0'},700)}
}
function search(mode){
 const input=document.getElementById('scan-'+mode),code=String(input.value).replace(/\D/g,''),st=states[mode];
 if(!code)return status(mode,'Ingrese o escanee un código de barras.','warn');
 if(!st.db.size)return status(mode,'Primero seleccione y cargue el archivo de esta pestaña.','warn');
 const item=st.db.get(code);
 if(!item){document.getElementById('result-'+mode).innerHTML='';return status(mode,'El código no fue encontrado en el archivo cargado en esta pestaña.','error')}
 if(!st.labels.some(x=>x.code===item.code&&x.classification===item.classification))st.labels.push(item);
 document.getElementById('result-'+mode).innerHTML=`<div class="status show ok"><b>✓ Código encontrado</b><br>${esc(item.code)}<br><b>CLASIFICACIÓN: ${esc(item.classification)}</b></div>`;
 input.value='';render(mode);status(mode,'Rótulo agregado a la lista de esta pestaña.','ok');input.focus();
}
async function makeWord(mode){
 const labels=states[mode].labels;
 if(!labels.length)return status(mode,'No hay rótulos para generar.','warn');
 try{
  const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,AlignmentType,VerticalAlign,BorderStyle}=window.docx;
  if(!Document||!Packer)throw Error('El generador de Word no está disponible.');
  const items=[...labels];while(items.length%6)items.push(null);
  const rows=[];
  for(let i=0;i<items.length;i+=6)rows.push(new TableRow({children:items.slice(i,i+6).map(item=>{
   const border={style:BorderStyle.DASHED,size:3,color:'888888'};
   const content=item?String(item.classification).split(/\s+/).map(t=>new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new TextRun({text:t,bold:true,size:24,font:'Arial'})]})):[new Paragraph('')];
   return new TableCell({width:{size:16.66,type:WidthType.PERCENTAGE},verticalAlign:VerticalAlign.CENTER,borders:{top:border,bottom:border,left:border,right:border},children:content});
  })}));
  const doc=new Document({sections:[{properties:{page:{margin:{top:500,bottom:500,left:500,right:500}}},children:[new Table({width:{size:100,type:WidthType.PERCENTAGE},rows})]}]});
  const blob=await Packer.toBlob(doc),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=`Rotulos_${mode==='nuevos'?'Nuevos':'Inventario'}_Mario_Valenzuela.docx`;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1500);
  status(mode,'Documento Word generado correctamente.','ok');
 }catch(err){status(mode,'Error al generar Word: '+err.message,'error')}
}
['nuevos','inventario'].forEach(mode=>{
 document.getElementById('file-'+mode).addEventListener('click',e=>e.currentTarget.value='');
 document.getElementById('file-'+mode).addEventListener('change',e=>load(mode,e.target.files[0]));
 document.getElementById('search-'+mode).onclick=()=>search(mode);
 document.getElementById('scan-'+mode).addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();search(mode)}});
 document.getElementById('clear-'+mode).onclick=()=>{states[mode].labels=[];document.getElementById('result-'+mode).innerHTML='';render(mode);status(mode,'Rótulos de esta pestaña eliminados.','ok')};
 document.getElementById('word-'+mode).onclick=()=>makeWord(mode);
});
function activate(mode){
 ['nuevos','inventario'].forEach(m=>{
  document.getElementById('tab-'+m).classList.toggle('active',m===mode);
  document.getElementById('panel-'+m).classList.toggle('active',m===mode);
 });
}
document.getElementById('tab-nuevos').onclick=()=>activate('nuevos');
document.getElementById('tab-inventario').onclick=()=>activate('inventario');
render('nuevos');render('inventario');
