const DAYS=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const pad=n=>String(n).padStart(2,'0');
const iso=d=>d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());
async function loadConfig(){const c=await (await fetch('data/config.json')).json();
  document.getElementById('name').textContent=c.name;document.getElementById('sub').textContent=c.title;
  const i=document.getElementById('logo');i.src=c.logo;i.alt=c.name+' logo';i.onerror=()=>{i.style.display='none'};return c}
function mondayOf(d){const x=new Date(d);x.setHours(0,0,0,0);x.setDate(x.getDate()-((x.getDay()+6)%7));return x}
const inRange=(h,a,b)=>h>=a&&h<b;
// PT slots just before and after each boxing-gym class: cfg.around = {before: hours, after: hours}
function aroundWindows(cfg,blocked){const a=cfg.around||{},nb=+a.before||0,na=+a.after||0,o=[];
  blocked.filter(b=>/boxing/i.test(b.label||'')).forEach(b=>{if(nb>0)o.push([Math.max(0,b.from-nb),b.from]);if(na>0)o.push([b.to,Math.min(24,b.to+na)])});return o}
// Rules for one date: the weekly pattern, unless the sheet has rows for that exact week and day, which replace the pattern for that day only
function dayRules(cfg,d){const day=DAYS[d.getDay()],wk=iso(mondayOf(d)),ov=(cfg.weeks||[]).filter(w=>w.week===wk&&w.day===day);
  if(ov.length){const bl=ov.filter(w=>w.type==='shift');return{open:ov.filter(w=>w.type==='open').map(w=>[w.from,w.to]).concat(aroundWindows(cfg,bl)),blocked:bl}}
  const bl=(cfg.blocked||[]).filter(b=>b.day===day);return{open:((cfg.open||{})[day]||[]).concat(aroundWindows(cfg,bl)),blocked:bl}}
// is hour h on date d open for PT? open windows minus shifts/classes minus away dates
function isOpen(cfg,d,h){const k=iso(d);
  if(h<(cfg.earliest||0))return false;
  if((cfg.away||[]).some(a=>k>=a.from&&k<=(a.to||a.from)))return false;
  const r=dayRules(cfg,d);
  if(!r.open.some(w=>inRange(h,w[0],w[1])))return false;
  return !r.blocked.some(b=>inRange(h,b.from,b.to))}
function hourRange(cfg){let lo=24,hi=0;const add=(a,b)=>{lo=Math.min(lo,a);hi=Math.max(hi,b)};
  Object.values(cfg.open).forEach(l=>l.forEach(w=>add(w[0],w[1])));(cfg.blocked||[]).forEach(b=>add(b.from,b.to));(cfg.weeks||[]).forEach(w=>add(w.from,w.to));aroundWindows(cfg,(cfg.blocked||[]).concat((cfg.weeks||[]).filter(w=>w.type==='shift'))).forEach(w=>add(w[0],w[1]));
  lo=Math.max(lo,cfg.earliest||0);return lo>=hi?[7,20]:[lo,hi]}
// busy = Set of (null while the live list is still loading: open slots show … and can't be tapped)
// busy = Set of "YYYY-MM-DD HH"
function renderWeek(el,cfg,busy,start,onFree){
  let h='<div class="h"></div>';const days=[];
  for(let i=0;i<7;i++){const d=new Date(start);d.setDate(d.getDate()+i);days.push(d);
    h+=`<div class="h">${DAYS[d.getDay()].slice(0,3)}<br>${d.getDate()}/${d.getMonth()+1}</div>`}
  const [lo,hi]=hourRange(cfg),now=new Date();
  for(let hr=lo;hr<hi;hr++){h+=`<div class="t">${pad(hr)}:00</div>`;
    days.forEach(d=>{const k=iso(d)+' '+pad(hr);
      const past=new Date(d.getFullYear(),d.getMonth(),d.getDate(),hr)<now;
      const open=isOpen(cfg,d,hr)&&!past;
      h+=!open?`<div class="off">–</div>`:busy===null?`<div class="wait" aria-label="Checking">…</div>`:busy.has(k)?`<div class="busy" aria-label="Taken">×</div>`:`<div class="free" data-k="${k}" role="button" tabindex="0">Free</div>`})}
  el.innerHTML=h;el.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){const n=e.target.closest&&e.target.closest('.free');if(n&&onFree){e.preventDefault();onFree(n.dataset.k)}}};el.onclick=e=>{const n=e.target.closest&&e.target.closest('.free');if(n&&onFree)onFree(n.dataset.k)}}

// Joe's Schedule tab (served with the busy list) overrides the saved schedule, so his changes go live without anyone in between
function applySchedule(cfg,b){if(b&&b.schedule&&b.schedule.open){cfg.open=b.schedule.open;cfg.blocked=b.schedule.blocked||[];cfg.away=b.schedule.away||[];cfg.weeks=b.schedule.weeks||[];cfg.around=b.schedule.around||{before:0,after:0}}return cfg}
