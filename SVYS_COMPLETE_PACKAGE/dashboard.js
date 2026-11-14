/* UI enhancements only — reads the existing `state`, never writes data or touches auth. */
(function(){
var $=function(id){return document.getElementById(id)};
// Toasts replace blocking alert() popups; messages are unchanged.
var host=document.createElement('div');host.className='toasts';host.setAttribute('aria-live','polite');document.body.appendChild(host);
window.alert=function(msg){var t=document.createElement('div');t.className='toast'+(/could not|not saved|not available|enter |invalid|incorrect|mandatory|too large|failed|unreachable|no payments|please/i.test(msg)?' err':'');t.textContent=msg;host.appendChild(t);setTimeout(function(){t.classList.add('out');setTimeout(function(){t.remove()},320)},4500)};
// Mobile drawer
function nav(open){document.body.classList.toggle('nav-open',open)}
$('menuBtn').addEventListener('click',function(){nav(!document.body.classList.contains('nav-open'))});
$('backdrop').addEventListener('click',function(){nav(false)});
document.addEventListener('keydown',function(e){if(e.key==='Escape'){nav(false);document.querySelectorAll('.modal.open').forEach(function(m){m.classList.remove('open')})}});
document.querySelectorAll('.sidebar .nav-item').forEach(function(b){b.addEventListener('click',function(){nav(false)})});
var title={navMonths:'Monthly Collection',navMembers:'Our Members'};
var _go=window.goSlide;window.goSlide=function(i){_go(i);$('topTitle').textContent=i?title.navMembers:title.navMonths;var m=document.querySelector('.slider-anchor');if(m&&m.scrollIntoView)m.scrollIntoView({behavior:'smooth',block:'start'})};
// Charts + activity
var esc=function(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]})};
function enhance(){
  var total=state.payments.reduce(function(s,p){return s+Number(p.amount||0)},0),target=state.members.length*expectedYear();
  var user=window.svysCurrentUser;var name=user&&(user.displayName||(user.email||'').split('@')[0]);
  $('greet').textContent=name?'Namaste, '+name:'Namaste';
  var W=600,H=210,pad=28,bw=(W-pad*2)/MONTHS.length,max=1,data=MONTHS.map(function(m){var c=state.payments.filter(function(p){return p.month===m[0]}).reduce(function(s,p){return s+Number(p.amount||0)},0);var e=state.members.length*m[1];max=Math.max(max,c,e);return{n:m[0].slice(0,3),c:c,e:e}});
  var svg='<svg class="chart" viewBox="0 0 '+W+' '+(H+22)+'" role="img" aria-label="Monthly collected versus expected amount">';
  for(var g=0;g<=3;g++){var y=H-g*H/3.4;svg+='<line x1="'+pad+'" x2="'+(W-pad)+'" y1="'+y+'" y2="'+y+'" stroke="#d6a64f33"/>'}
  data.forEach(function(d,i){var x=pad+i*bw+bw*.18,w=bw*.64,he=d.e/max*(H-12),hc=d.c/max*(H-12);
    svg+='<rect x="'+x+'" y="'+(H-he)+'" width="'+w+'" height="'+he+'" rx="5" fill="#f1e4cd"/>';
    svg+='<rect class="bar-fill" style="animation-delay:'+i*45+'ms" x="'+x+'" y="'+(H-hc)+'" width="'+w+'" height="'+hc+'" rx="5" fill="url(#bg)"><title>'+d.n+': '+money(d.c)+' of '+money(d.e)+'</title></rect>';
    svg+='<text x="'+(x+w/2)+'" y="'+(H+16)+'" text-anchor="middle">'+d.n+'</text>'});
  svg+='<defs><linearGradient id="bg" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#E89B35"/><stop offset="1" stop-color="#762B36"/></linearGradient></defs></svg>';
  $('chartMonths').innerHTML=svg;
  var pct=target?Math.min(100,total/target*100):0,r=62,C=2*Math.PI*r;
  $('chartDonut').innerHTML='<svg viewBox="0 0 160 160" width="190" height="190" role="img" aria-label="Collection progress '+pct.toFixed(0)+' percent"><circle cx="80" cy="80" r="'+r+'" fill="none" stroke="#f1e4cd" stroke-width="16"/><circle class="donut-arc" cx="80" cy="80" r="'+r+'" fill="none" stroke="url(#bg)" stroke-width="16" stroke-linecap="round" stroke-dasharray="'+(C*pct/100)+' '+C+'" transform="rotate(-90 80 80)"/></svg><div class="center"><b>'+pct.toFixed(0)+'%</b><span>of yearly minimum</span></div>';
  $('donutLegend').innerHTML='<span><i style="background:#762B36"></i>Collected '+money(total)+'</span><span><i style="background:#f1e4cd"></i>Target '+money(target)+'</span>';
  var recent=state.payments.slice().sort(function(a,b){return String(b.date).localeCompare(String(a.date))}).slice(0,6);
  $('recentActivity').innerHTML=recent.length?recent.map(function(p){var m=state.members.find(function(x){return String(x.id)===String(p.memberId)});return'<div class="act"><div class="avatar">'+esc(initials(m?m.name:'?'))+'</div><div><b>'+esc(m?m.name:'Member')+'</b><small>'+esc(p.month)+' · '+esc(p.date||'—')+'</small></div><span class="amt">+'+money(p.amount)+'</span></div>'}).join(''):'<div class="empty">No payments recorded yet.</div>';
}
var _r=window.render;window.render=function(){_r.apply(this,arguments);try{enhance()}catch(e){console.error('Dashboard enhancement failed:',e)}};
})();
