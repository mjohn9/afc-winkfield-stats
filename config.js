// Shared settings for every page.
// GoatCounter code for visit and search counting. Leave "" to switch counting off.
const GC_CODE="afcwinkfield";
if(GC_CODE){const g=document.createElement("script");g.async=true;g.dataset.goatcounter=`https://${GC_CODE}.goatcounter.com/count`;g.src="//gc.zgo.at/count.js";document.head.appendChild(g)}
function track(path,title){
  if(!GC_CODE)return;
  const send=()=>{if(window.goatcounter&&window.goatcounter.count){window.goatcounter.count({path,title,event:true});return true}return false};
  if(!send())setTimeout(send,1500);
}
