const DAYS=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const pad=n=>String(n).padStart(2,'0');
const iso=d=>d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());
async function loadConfig(){const c=await (await fetch('data/config.json')).json();
  document.getElementById('name').textContent=c.name;document.getElementById('sub').textContent=c.title;
  const i=document.getElementById('logo');i.src=c.logo;i.alt=c.name+' logo';i.onerror=()=>{i.style.display='none'};return c}
function mondayOf(d){const x=new Date(d);x.setHours(0,0,0,0);x.setDate(x.getDate()-((x.getDay()+6)%7));return x}
const inRange=(h,a,b)=>h>=a&&h<b;
// is hour h on date d open for PT? open windows minus shifts/classes minus away dates
function isOpen(cfg,d,h){const day=DAYS[d.getDay()],k=iso(d);
  if((cfg.away||[]).some(a=>k>=a.from&&k<=(a.to||a.from)))return false;
  if(!(cfg.open[day]||[]).some(w=>inRange(h,w[0],w[1])))return false;
  return !(cfg.blocked||[]).some(b=>b.day===day&&inRange(h,b.from,b.to))}
function hourRange(cfg){let lo=24,hi=0;const add=(a,b)=>{lo=Math.min(lo,a);hi=Math.max(hi,b)};
  Object.values(cfg.open).forEach(l=>l.forEach(w=>add(w[0],w[1])));(cfg.blocked||[]).forEach(b=>add(b.from,b.to));
  return lo>hi?[7,20]:[lo,hi]}
// busy = Set of "YYYY-MM-DD HH"
function renderWeek(el,cfg,busy,start,onFree){
  let h='<div class="h"></div>';const days=[];
  for(let i=0;i<7;i++){const d=new Date(start);d.setDate(d.getDate()+i);days.push(d);
    h+=`<div class="h">${DAYS[d.getDay()].slice(0,3)}<br>${d.getDate()}/${d.getMonth()+1}</div>`}
  const [lo,hi]=hourRange(cfg),now=new Date();
  for(let hr=lo;hr<hi;hr++){h+=`<div class="t">${pad(hr)}:00</div>`;
    days.forEach(d=>{const k=iso(d)+' '+pad(hr);
      const past=new Date(d.getFullYear(),d.getMonth(),d.getDate(),hr)<now;
      const open=isOpen(cfg,d,hr)&&!past&&!busy.has(k);
      h+=open?`<div class="free" data-k="${k}">Free</div>`:`<div class="busy">–</div>`})}
  el.innerHTML=h;if(onFree)el.querySelectorAll('.free').forEach(n=>n.onclick=()=>onFree(n.dataset.k))}
