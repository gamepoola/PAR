const $=x=>document.getElementById(x),key='radisson.manual.inventory.nov2026.v1';
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt=v=>v==null?'—':Number(v).toLocaleString('th-TH',{maximumFractionDigits:4});
let edits={};
function col(n){let c='';while(n){n--;c=String.fromCharCode(65+n%26)+c;n=Math.floor(n/26)}return c}
function values(s,item,d){
 let close;for(let i=0;i<=d;i++){
 const [initial,...tail]=item[7][i],open=tail[5]?close:initial;
 const qty=tail.slice(0,4).map((val,k)=>{const a=col(11+6*i+k)+item[0];return Object.hasOwn(edits[s]||{},a)?edits[s][a]:val});
 close=tail[6]?Number(open||0)-Number(qty[0]||0)+Number(qty[1]||0)+Number(qty[2]||0)-Number(qty[3]||0):tail[4];
 if(i===d)return {open,qty,close};
 }}
function closeCell(cell,item,close){cell.textContent=fmt(close);cell.className=close==null?'':close<0?'neg':item[4]>0&&close/item[4]<item[5]?'low':''}
function draw(){
 const s=+$('shop').value,d=+$('day').value,g=$('group').value,q=$('search').value.toLowerCase().trim();
 const list=stock[s].filter(a=>(!g||a[3]===g)&&(a[1]+' '+a[2]).toLowerCase().includes(q));$('count').textContent=list.length+' รายการ';
 $('rows').innerHTML=list.map(a=>{const v=values(s,a,d);return '<tr data-r="'+a[0]+'"><td>'+esc(a[2])+'<small>'+esc(a[1])+' · ตัวคูณต่อหน่วยบรรจุ '+fmt(a[4])+'</small></td><td>'+fmt(a[5])+' / '+fmt(a[6])+'</td><td>'+fmt(v.open)+'</td>'+v.qty.map((n,k)=>'<td><input aria-label="'+['SALE','ISS.','TF/IN','TF/OUT'][k]+' '+esc(a[1])+'" type="number" step="any" data-k="'+k+'" value="'+esc(n)+'"></td>').join('')+'<td></td></tr>'}).join('');Array.from($('rows').rows).forEach((tr,i)=>closeCell(tr.lastElementChild,list[i],values(s,list[i],d).close));
}
function groups(){const gs=[...new Set(stock[+$('shop').value].map(a=>a[3]))];$('group').innerHTML='<option value="">ทุกหมวด</option>'+gs.map(g=>'<option>'+esc(g)+'</option>').join('');draw()}
$('day').innerHTML=Array.from({length:30},(_,i)=>'<option value="'+i+'">'+(i+1)+' พ.ย. 2026</option>').join('');
$('day').value=Math.min(29,new Date().getDate()-1);$('shop').onchange=groups;$('group').onchange=$('day').onchange=draw;$('search').oninput=draw;
$('rows').onchange=e=>{const input=e.target;if(!input.matches('input'))return;const tr=input.closest('tr'),s=+$('shop').value,d=+$('day').value,k=+input.dataset.k,a=stock[s].find(a=>a[0]===+tr.dataset.r),addr=col(11+6*d+k)+a[0];
 if(input.validity.badInput){draw();return}const n=input.value===''?null:Number(input.value);edits[s]??={};if(n===a[7][d][1+k])delete edits[s][addr];else edits[s][addr]=n;
 closeCell(tr.lastElementChild,a,values(s,a,d).close);
 try{$('status').textContent='บันทึกในเครื่องแล้ว · '+a[1]}catch{$('status').textContent='เบราว์เซอร์บันทึกลงเครื่องไม่ได้ กรุณาส่งออก Excel ก่อนปิดหน้า'}};
function patch(xml,changes){
 const rest=new Map(Object.entries(changes));const result=xml.replace(/<c\b([^>]*?\br="([A-Z]+\d+)"[^>]*?)(?:\s*\/>|>([\s\S]*?)<\/c>)/g,(all,attrs,addr,body)=>{if(!rest.has(addr))return all;if(body?.includes('<f'))throw Error('ช่องสูตร '+addr);const v=rest.get(addr);rest.delete(addr);return v===null?'<c'+attrs+'/>':'<c'+attrs+'><v>'+Number(v)+'</v></c>'});
 if(rest.size)throw Error('ไม่พบช่อง '+rest.keys().next().value);return result;
}
async function exportBook(){const zip=await JSZip.loadAsync(template,{base64:true});for(const [s,changes] of Object.entries(edits)){if(!['0','1','2'].includes(s)||!Object.keys(changes).length)continue;const path='xl/worksheets/sheet'+(+s+1)+'.xml';zip.file(path,patch(await zip.file(path).async('string'),changes))}
 const path='xl/workbook.xml';zip.file(path,(await zip.file(path).async('string')).replace(/<calcPr\b[^>]*\/>/,'<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'));return zip.generateAsync({type:'blob',compression:'DEFLATE'})}
$('exp').onclick=async()=>{const b=$('exp');b.disabled=true;try{const url=URL.createObjectURL(await exportBook()),a=document.createElement('a');a.href=url;a.download='Radisson_Bar_Par_Nov2026_Updated.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(url),60000);$('status').textContent='ส่งออกแล้ว · เปิดไฟล์ใน Microsoft Excel เพื่อคำนวณสูตรทั้งหมด'}catch(e){$('status').textContent='ส่งออกไม่สำเร็จ · '+e.message}finally{b.disabled=false}};
groups();
const recipeKey='bar.recipes.2026-11.v2';
let recipeBook={},pending=null,dayReports=edits.reports||{};
const canon=v=>v.normalize('NFKC').trim().replace(/\s+/g,' ').toLowerCase();
const loc=()=>({s:+$('shop').value,d:+$('day').value});
function records(text,width){
 const source=text.replace(/^\uFEFF/,'').trim(),sep=source.split(/\r?\n/)[0].includes('\t')?'\t':',';
 const rows=[];let row=[],word='',quoted=false;
 for(let i=0;i<=source.length;i++){const c=source[i];if(c==='"'){if(quoted&&source[i+1]==='"'){word+='"';i++}else quoted=!quoted}
 else if(!quoted&&(c===sep||c==='\n'||c===undefined)){row.push(word.trim());word='';if(c!==sep){if(row.some(Boolean))rows.push(row);row=[]}}else if(c!==undefined)word+=c}
 if(quoted)throw Error('เครื่องหมายคำพูดใน CSV ไม่ครบ');if(rows.length&&/^(เมนู|menu|item|item name|description)$/i.test(rows[0][0])&&!Number.isFinite(Number(rows[0].at(-1))))rows.shift();
 if(!rows.length||rows.some(r=>r.length!==width||r.some(v=>!v)))throw Error('ตรวจจำนวนคอลัมน์และช่องว่างในข้อมูล');return rows;
}
function resolveRecipes(s,text){const map=new Map();for(const [name,article,q] of records(text,3)){
 const matches=stock[s].filter(a=>String(a[1])===article);
 if(matches.length!==1||!(Number(q)>0&&Number.isFinite(Number(q))))throw Error('ตรวจ Article และปริมาณ · '+name+' / '+article);
 const id=canon(name),parts=map.get(id)||[];if(parts.some(x=>x.a===matches[0]))throw Error('วัตถุดิบซ้ำในสูตร '+name);
 parts.push({a:matches[0],q:Number(q)});map.set(id,parts)}return map}
let sourceInfo=null;
function cancel(){pending=null;document.getElementById('auditResult')?.remove();$('proposal').textContent='';$('post').hidden=true}
function context(){sourceInfo=null;cancel();$('recipes').value=recipeBook[loc().s]||'';$('sales').value='';$('salesFile').value='';$('recipeFile').value='';$('notice').textContent='';$('recipeNotice').textContent=''}
$('shop').addEventListener('change',context);$('day').addEventListener('change',context);
$('recipes').oninput=$('sales').oninput=cancel;
$('recipeFile').onchange=async()=>{const f=$('recipeFile').files[0];if(f){cancel();$('recipes').value=await f.text()}};
$('storeRecipes').onclick=()=>{try{const {s}=loc(),t=$('recipes').value;const map=resolveRecipes(s,t);
 const next={...recipeBook,[s]:t};recipeBook=next;cancel();$('recipeNotice').textContent='บันทึกแล้ว '+map.size+' สูตร'}catch(e){$('recipeNotice').textContent=e.message}};
function usage(s,recipe,sales){
 const map=resolveRecipes(s,recipe),total=new Map(),audit=[],missing=new Set();
 for(const [name,v] of records(sales,2)){const count=Number(v);if(!Number.isFinite(count)||count<0)throw Error('ตรวจจำนวนขายสุทธิ '+name);if(count===0)continue;
 const parts=map.get(canon(name));if(!parts){missing.add(name);continue}
 for(const {a,q} of parts){total.set(a,(total.get(a)||0)+q*count);audit.push(name+' × '+fmt(count)+' → '+a[1]+' '+a[2]+' × '+fmt(q)+' = '+fmt(count*q))}}
 if(missing.size)throw Error('ยังไม่มีสูตรสำหรับ '+[...missing].join(', '));return {total,audit};
}
const csvField=v=>'"'+String(v??'').replaceAll('"','""')+'"';
$('salesFile').onchange=async()=>{sourceInfo=null;cancel();const f=$('salesFile').files[0];if(!f)return;$('sales').value='';
 const fields=['salesFile','shop','day','previewSales','sales'];fields.forEach(id=>$(id).disabled=true);
 try{if(/\.(csv|tsv|txt)$/i.test(f.name))$('sales').value=await f.text();
 else{if(f.size>8*1024*1024)throw Error('ไฟล์ใหญ่เกิน 8 MB');$('notice').textContent='กำลังอ่านรายงานด้วย OpenAI…';
 const resp=await api('/api/read',{method:'POST',headers:{'Content-Type':f.type,'X-Bar-App':'daily'},body:f});const result=await resp.json();if(!resp.ok)throw Error(result.error);
 $('sales').value=result.lines.map(l=>csvField(l.menu)+','+(l.quantity??'')).join('\n');
 const isPDF=/\.pdf$/i.test(f.name);
 if(!isPDF){$('corrections').open=true;$('notice').textContent='ตรวจรายการ แล้วกดตรวจยอดอีกครั้ง';return}
 const dateMatch=/^2026-11-(0[1-9]|[12]\d|30)$/.exec(result.date);
 const bar=['Banquet','The Taste','Mangrove'].findIndex(b=>canon(b)===canon(result.outlet||''));
 if(!dateMatch||bar<0){$('sales').value='';throw Error('ลงยอดไม่ได้: PDF ต้องระบุบาร์เดียวและวันที่เดียวในเดือน พ.ย. 2026 กรุณาเลือกหรือออกรายงานใหม่')}
 $('shop').value=String(bar);renderIngredients();$('day').value=String(Number(dateMatch[1])-1);groups();$('recipes').value=recipeBook[bar]||'';
 const bytes=await crypto.subtle.digest('SHA-256',await f.arrayBuffer());
 const hash=Array.from(new Uint8Array(bytes),v=>v.toString(16).padStart(2,'0')).join('');
 sourceInfo={file:f.name,hash,notes:result.notes||''};
 if(edits.imports?.[bar+'-'+loc().d]?.hash===hash){$('notice').textContent='PDF นี้ลงยอดของบาร์และวันที่นี้แล้ว';return}
 await $('previewSales').onclick(true);return}
 $('corrections').open=true;
 $('notice').textContent='นำเข้าแล้ว ตรวจเมนูและจำนวนขายก่อนคำนวณ'}catch(e){$('notice').textContent=e.message}finally{fields.forEach(id=>$(id).disabled=false)}};
$('previewSales').onclick=async(auto)=>{cancel();try{const {s,d}=loc();if(!recipeBook[s]||recipeBook[s]!==$('recipes').value)throw Error('กรุณาบันทึกสูตรของบาร์ก่อนคำนวณ');
 const {total,audit}=usage(s,recipeBook[s],$('sales').value),prior=dayReports[s+'-'+d]||[];
 for(const r of prior){const a=stock[s].find(x=>x[0]===r);if(a&&!total.has(a))total.set(a,0)}
 if(!total.size)throw Error('ไม่มีรายการให้ลงยอด');const changes=Array.from(total,([a,n])=>({a,n:Number(n.toFixed(6)),old:values(s,a,d).qty[0]||0}));
 $('proposal').textContent='รายละเอียดตามสูตร\n'+audit.join('\n')+'\n\nยอด SALE รวมทั้งวัน (เดิม → ใหม่)\n'+changes.map(x=>x.a[1]+' '+x.a[2]+'   '+fmt(x.old)+' → '+fmt(x.n)).join('\n');
 const candidate={s,d,changes};pending=candidate;
 $('notice').textContent='OpenAI กำลังตรวจชื่อเมนู จำนวน สูตร และยอดตัด…';
 const src=$('salesFile').files[0];let original=null;
 if(src && /\.(pdf|png|jpe?g)$/i.test(src.name)){
 original=await new Promise((ok,fail)=>{const r=new FileReader();r.onload=()=>ok(r.result);r.onerror=fail;r.readAsDataURL(src)});
 }else if(src)original=(await src.text()).slice(0,150000);
 const res=await api('/api/check',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({original,day:d+1,outlet:$('shop').selectedOptions[0].textContent,sales:$('sales').value,recipes:recipeBook[s],usage:changes.map(x=>({article:x.a[1],name:x.a[2],amount:x.n}))})});
 const report=await res.json();if(pending!==candidate)return;if(!res.ok)throw Error(report.error||'ตรวจสอบไม่สำเร็จ');
 if(sourceInfo?.notes)report.issues.unshift(sourceInfo.notes);candidate.review=report.issues;candidate.sales=$('sales').value;candidate.source=sourceInfo;
 const box=document.createElement('div');box.id='auditResult';box.style.whiteSpace='pre-line';box.textContent=report.issues.length?'OpenAI พบจุดที่ต้องตรวจ\n'+report.issues.join('\n'):'OpenAI ตรวจแล้ว ไม่พบข้อสังเกตจากข้อมูลที่ส่งตรวจ';
 $('proposal').after(box);
 if(auto===true&&!report.issues.length){await $('post').onclick();return}
 if(report.issues.length){$('corrections').open=true;const label=document.createElement('label'),cb=document.createElement('input');cb.type='checkbox';label.append(cb,' ตรวจเทียบรายงานต้นฉบับและยอมรับยอดนี้แล้ว');box.append(document.createElement('br'),label);cb.onchange=()=>{$('post').hidden=!cb.checked}}else $('post').hidden=false;
 $('notice').textContent='การยืนยันจะแทนยอด SALE เฉพาะรายการที่แสดง การนำเข้ารายงานเดิมซ้ำจะไม่บวกยอดซ้ำ'}catch(e){cancel();$('corrections').open=true;$('notice').textContent=e.message}};
$('post').onclick=()=>{if(!pending)return;const {s,d,changes}=pending;if(changes.some(x=>(values(s,x.a,d).qty[0]||0)!==x.old)){cancel();$('notice').textContent='ยอดเปลี่ยนแล้ว กรุณาคำนวณอีกครั้ง';return}
 const next=structuredClone(edits);next[s]??={};for(const x of changes)next[s][col(11+d*6)+x.a[0]]=x.n;
 next.imports={...next.imports,[s+'-'+d]:{...pending.source,file:pending.source?.file||$('salesFile').files[0]?.name||'กรอกเอง',sales:pending.sales,review:pending.review,at:new Date().toISOString()}};
 next.reports={...dayReports,[s+'-'+d]:changes.map(x=>x.a[0])};try{edits=next;dayReports=next.reports;draw();cancel();$('notice').textContent='ลงยอด SALE วันที่ '+(d+1)+' พ.ย. แล้ว พร้อมส่งออก Excel'}catch{$('notice').textContent='พื้นที่จัดเก็บไม่พอ ยังไม่ได้ลงยอด'}};
context();

if(true){
 let revision=0;
 const main=document.querySelector('main');
 function applyState(a){edits=a.edits;recipeBook=a.recipes;dayReports=edits.reports||{};revision=a.rev;groups();context()}
 async function getState(){const res=await api('/api/state');if(!res.ok)throw Error();return res.json()}
 main.inert=true;
 getState().then(a=>{applyState(a);main.inert=false;$('status').textContent='เปิดฐานข้อมูลกลางแล้ว'}).catch(()=>{$('status').textContent='เปิดฐานข้อมูลไม่ได้ กรุณาโหลดหน้าใหม่'});
 for(const [el,event] of [['rows','onchange'],['post','onclick'],['storeRecipes','onclick']]){
 const handler=$(el)[event];$(el)[event]=async function(e){
 const previous=JSON.stringify([edits,recipeBook]);handler.call(this,e);if(previous===JSON.stringify([edits,recipeBook]))return;
 main.inert=true;try{const response=await api('/api/state',{method:'POST',headers:{'Content-Type':'application/json','X-Bar-App':'daily'},body:JSON.stringify({edits,recipes:recipeBook,rev:revision})});
 if(!response.ok)throw Error();revision=(await response.json()).rev;$('status').textContent='บันทึกฐานข้อมูลหลักแล้ว';main.inert=false}
 catch{$('notice').textContent='ยืนยันการบันทึกไม่ได้ กรุณาโหลดหน้าใหม่และตรวจรายงาน';$('status').textContent='ยืนยันการบันทึกไม่ได้ กรุณาโหลดหน้าใหม่เพื่อตรวจยอดล่าสุด';alert($('status').textContent)}
 }}
}

function dailyRows(s,d){
 const report=edits.imports?.[s+'-'+d],outlet=['Banquet','The Taste','Mangrove'][s],date='2026-11-'+String(d+1).padStart(2,'0');
 const rows=[['วันที่',date,'บาร์',outlet],['เอกสาร',report?.file||'ข้อมูลในตารางพาร์'],['เวลานำเข้า (UTC)',report?.at||'']];
 if(report){rows.push(['ข้อสังเกตจากการตรวจ',...(report.review||[])],[],['เมนู POS','จำนวนขายสุทธิ'],...records(report.sales,2))}
 rows.push([],['Article','รายการ','ยอดเปิด','SALE','ISS.','TF/IN','TF/OUT','ยอดปิด']);
 for(const item of stock[s]){const v=values(s,item,d);rows.push([item[1],item[2],v.open,...v.qty,v.close])}return rows;
}
$('dailyCsv').onclick=()=>{const {s,d}=loc();const text=dailyRows(s,d).map(row=>row.map(v=>csvField(typeof v==='string'&&/^\s*[=+@-]/.test(v)?"'"+v:v)).join(',')).join('\r\n');
 const url=URL.createObjectURL(new Blob(['\uFEFF'+text],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download=['Banquet','The_Taste','Mangrove'][s]+'_Par_2026-11-'+String(d+1).padStart(2,'0')+'.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),60000)};

function expandRecipe(recipe, scale = 1, ancestors = []) {
 if (ancestors.includes(recipe.id)) throw Error('สูตรย่อยอ้างถึงตัวเอง');
 return recipe.ingredients.flatMap(line => {
   if (line.type === 'A') return [{...line, amount: line.act_qty * scale}];
   const child = catalog.recipes.find(item => item.id === line.subrecipe_id);
   if (!child || line.unit !== 'pc') throw Error('รอตรวจสูตรย่อยและหน่วย • ' + line.description);
   return expandRecipe(child, scale * line.act_qty, [...ancestors, recipe.id]);
 });
}
function renderIngredients() {
 const recipe = catalog.recipes.find(item => item.id === $('mc').value);
 const box = $('ingredients');
 box.replaceChildren(); $('catalogStatus').textContent = ''; $('approve').disabled = true;
 $('alias').value = recipe.menu;
 try {
   const lines = expandRecipe(recipe);
   const source = document.createElement('p');
   source.textContent = recipe.source + ' • หน้า ' + recipe.page + ' • ตรวจปริมาณต่อหนึ่งเสิร์ฟ แล้วระบุจำนวนหน่วย SALE ต่อ 1 หน่วยสูตร';
   box.append(source);
   for (const line of lines) {
     const label = document.createElement('label'); label.className = 'ingredient';
     label.textContent = line.description + ' • ' + line.article + ' • '+line.amount+' '+line.unit+' ';
     const matches = stock[loc().s].filter(a => String(a[1]) === line.article);
     if (matches.length !== 1) { label.append(' — ต้องจับคู่รหัสสต็อกของบาร์ในชุดสูตร CSV'); box.append(label); continue; }
     const input = document.createElement('input'); input.type='number'; input.min='0'; input.step='any';
     input.dataset.article=line.article; input.dataset.qty=line.amount; input.placeholder='ตัวคูณ SALE'; label.append(input); box.append(label);
   }
   $('approve').disabled = box.querySelectorAll('input').length !== lines.length;
 } catch(error) { $('catalogStatus').textContent=error.message; }
}
$('mc').replaceChildren(...catalog.recipes.map(recipe => new Option(recipe.group+' · '+recipe.menu,recipe.id)));
$('mc').onchange=renderIngredients;
$('shop').addEventListener('change',renderIngredients);
renderIngredients();
$('approve').onclick=async()=>{
 try {
   const name=$('alias').value.trim();
   if(!name) throw Error('ระบุชื่อเมนู POS');
   const sums=new Map();
   for(const input of $('ingredients').querySelectorAll('input')){
     const amount=Number(input.value)*Number(input.dataset.qty);
     if(!Number.isFinite(amount)||amount<=0) throw Error('ตรวจตัวคูณ SALE ให้ครบทุกวัตถุดิบ');
     sums.set(input.dataset.article,(sums.get(input.dataset.article)||0)+amount);
   }
   const prior=recipeBook[loc().s]||'';
   const rows=prior.trim()?records(prior,3):[];
   if(rows.some(row=>canon(row[0])===canon(name))&&!confirm('แทนสูตรเดิมของ '+name+' สำหรับบาร์นี้?'))return;
   const next=rows.filter(row=>canon(row[0])!==canon(name)).concat(Array.from(sums,([code,qty])=>[name,code,qty]));
   $('recipes').value=next.map(row=>row.map(csvField).join(',')).join('\n');
   await $('storeRecipes').onclick();
   $('catalogStatus').textContent=$('status').textContent+' · '+$('recipeNotice').textContent;
 } catch(error){$('catalogStatus').textContent=error.message}
};
