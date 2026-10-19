// Shared settings for every page.
// GoatCounter code for visit and search counting. Leave "" to switch counting off.
const GC_CODE="afcwinkfield";
if(GC_CODE){const g=document.createElement("script");g.async=true;g.dataset.goatcounter=`https://${GC_CODE}.goatcounter.com/count`;g.src="//gc.zgo.at/count.js";document.head.appendChild(g)}
function track(path,title){
  if(!GC_CODE)return;
  const send=()=>{if(window.goatcounter&&window.goatcounter.count){window.goatcounter.count({path,title,event:true});return true}return false};
  if(!send())setTimeout(send,1500);
}

// Set to false to hide the "Cock-up of the Season" awards from the trophy room.
const SHOW_FUN_AWARDS=true;

// Social media pages shown as icons at the bottom of every page. Change a link here and it changes everywhere.
const SOCIALS=[
  ["Instagram","https://www.instagram.com/afcwinkfield",'<rect x="3" y="3" width="18" height="18" rx="5"/><circle cx="12" cy="12" r="4"/><path d="M17.5 6.5h.01"/>'],
  ["Facebook","https://facebook.com/afcwinkfield",'<path d="M17 2h-3a4 4 0 0 0-4 4v3H7v4h3v9h4v-9h3l1-4h-4V6a1 1 0 0 1 1-1h2z"/>'],
  ["LinkedIn","https://www.linkedin.com/company/afc-winkfield",'<path d="M16 8a6 6 0 0 1 6 6v7h-4v-7a2 2 0 0 0-4 0v7h-4v-7a6 6 0 0 1 6-6z"/><rect x="2" y="9" width="4" height="12"/><circle cx="4" cy="4" r="2"/>'],
  ["X","https://www.x.com/afcwinkfield",'<path d="M4 4l16 16M20 4l-6.5 6.5M4 20l6.5-6.5"/>']
];
function socialLinks(){
  return '<div style="display:flex;justify-content:center;gap:14px;padding:14px 0 4px">'+SOCIALS.map(([n,u,p])=>`<a href="${u}" target="_blank" rel="noopener noreferrer" aria-label="AFC Winkfield on ${n}" title="${n}" onclick="track('social/${n}','Social: ${n}')" onmouseover="this.style.color='#D1A04A';this.style.borderColor='#D1A04A'" onmouseout="this.style.color='#A39D90';this.style.borderColor='#2A2A2E'" style="width:46px;height:46px;border-radius:50%;border:1px solid #2A2A2E;display:flex;align-items:center;justify-content:center;color:#A39D90;background:#17171A;transition:color .15s,border-color .15s"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${p}</svg></a>`).join("")+'</div>';
}
