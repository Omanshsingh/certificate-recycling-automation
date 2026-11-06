/* Desktop analytics, using saved issuance history and certificate PDFs. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const number = value => Number(value || 0).toLocaleString();
  const set = (id,value) => { $(id).textContent=value; };
  const day = value => new Date(`${value}T12:00:00`);
  const dateLabel = value => { const parts=String(value || '').split('-'); return parts.length===3 ? `${parts[2]}/${parts[1]}/${parts[0].slice(-2)}` : String(value || ''); };
  const fullDateLabel = dateLabel;
  const periods = {today:'today',week:'in the last 7 days',month:'this month',quarter:'in the last 90 days'};
  const node = (tag,className='',text) => {
    const el=document.createElement(tag); el.className=className;
    if(text!==undefined) el.textContent=text; return el;
  };
  const svg = (tag,attrs={},text) => {
    const el=document.createElementNS('http://www.w3.org/2000/svg',tag);
    for(const [key,value] of Object.entries(attrs)) el.setAttribute(key,value);
    if(text!==undefined) el.textContent=text; return el;
  };
  let request=0,controller,current,chartIndex=-1,tooltip,tooltipOwner,dismissedOwner;
  function hideTooltip(owner) {
    if(owner && owner!==tooltipOwner) return;
    if(tooltipOwner) {
      const ids=(tooltipOwner.getAttribute('aria-describedby') || '').split(' ').filter(id=>id && id!=='analytics-tooltip');
      if(ids.length) tooltipOwner.setAttribute('aria-describedby',ids.join(' '));
      else tooltipOwner.removeAttribute('aria-describedby');
    }
    if(tooltip) tooltip.hidden=true;
    tooltipOwner=null;
  }
  function showTooltip(owner,title,rows,caption,anchor) {
    if(dismissedOwner===owner || $('analytics-workspace').hidden || $('analytics-content').inert) return;
    if(!tooltip) {
      tooltip=node('div','analytics-tooltip'); tooltip.id='analytics-tooltip';
      tooltip.setAttribute('role','tooltip'); tooltip.hidden=true; $('analytics-workspace').append(tooltip);
    }
    if(tooltipOwner!==owner) hideTooltip();
    tooltipOwner=owner;
    const ids=new Set((owner.getAttribute('aria-describedby') || '').split(' ').filter(Boolean)); ids.add(tooltip.id);
    owner.setAttribute('aria-describedby',[...ids].join(' '));
    const content=[node('strong','analytics-tooltip-title',title)];
    for(const row of rows) {
      const line=node('div',`analytics-tooltip-row ${row.kind || ''}`);
      const label=node('span','analytics-tooltip-label');
      if(row.kind) { const swatch=node('i'); swatch.setAttribute('aria-hidden','true'); label.append(swatch); }
      label.append(document.createTextNode(row.label));
      line.append(label,node('b','',row.value)); content.push(line);
    }
    if(caption) content.push(node('span','analytics-tooltip-caption',caption));
    tooltip.replaceChildren(...content); tooltip.hidden=false;
    const bounds=owner.getBoundingClientRect(),point=anchor || {x:bounds.left+bounds.width/2,y:bounds.bottom};
    const rect=tooltip.getBoundingClientRect(),gap=14,padding=12;
    let x=point.x+gap,y=point.y+gap;
    if(x+rect.width>innerWidth-padding) x=point.x-rect.width-gap;
    if(y+rect.height>innerHeight-padding) y=point.y-rect.height-gap;
    tooltip.style.left=`${Math.max(padding,Math.min(x,innerWidth-rect.width-padding))}px`;
    tooltip.style.top=`${Math.max(padding,Math.min(y,innerHeight-rect.height-padding))}px`;
  }
  function attachTooltip(owner,content) {
    const show=event=>{ const info=content(); showTooltip(owner,info.title,info.rows,info.caption,event?.type.startsWith('pointer')?{x:event.clientX,y:event.clientY}:undefined); };
    owner.addEventListener('pointerenter',event=>{ dismissedOwner=null; show(event); });
    owner.addEventListener('pointermove',show);
    owner.addEventListener('pointerleave',()=>{ if(document.activeElement!==owner) hideTooltip(owner); dismissedOwner=null; });
    owner.addEventListener('focus',event=>{ dismissedOwner=null; show(event); });
    owner.addEventListener('blur',()=>hideTooltip(owner));
  }
  const review = detail => $('analytics-workspace').dispatchEvent(new CustomEvent('analytics:review',{detail:{...detail,asOf:current?.range.end}}));
  const files = () => $('analytics-workspace').dispatchEvent(new CustomEvent('analytics:files'));
  async function load() {
    hideTooltip(); dismissedOwner=null;
    const token=++request; controller?.abort(); controller=new AbortController();
    $('analytics-workspace').setAttribute('aria-busy','true'); $('analytics-content').inert=true;
    $('analytics-message').hidden=false;
    $('analytics-message').replaceChildren(node('span','','Loading dashboard…'));
    try {
      const query=new URLSearchParams({period:$('analytics-period').value,plant:$('analytics-plant').value});
      const response=await fetch(`/api/analytics?${query}`,{signal:controller.signal});
      if(response.status===404) throw new Error('Close and reopen the certificate app to load the dashboard.');
      if(!response.headers.get('content-type')?.includes('application/json')) throw new Error('The dashboard could not load. Try again or reopen the app.');
      const data=await response.json();
      if(!response.ok) throw new Error(data.error || 'The dashboard could not load.');
      if(token!==request) return;
      current=data; $('analytics-content').hidden=false; chartIndex=-1; render(data);
      const notes=[];
      if(!data.coverage.scan_complete) notes.push('Some Excel entries could not be read.');
      if(data.coverage.unknown_generation_dates) notes.push(`${number(data.coverage.unknown_generation_dates)} certificates need dates.`);
      if(data.coverage.undated_inwards) {
        const count=data.coverage.undated_inwards;
        notes.push(count===1 ? '1 inward needs a received date.' : `${number(count)} inwards need received dates.`);
      }
      if(data.coverage.missing_cod) notes.push(`${number(data.coverage.missing_cod)} certificates need COD numbers.`);
      if(data.coverage.folder_needs_details) notes.push(`${number(data.coverage.folder_needs_details)} saved PDFs need their date or type checked.`);
      $('analytics-message').hidden=!notes.length;
      if(notes.length) {
        const actions=node('div','analytics-message-actions');
        const addAction=(label, action)=>{
          const button=node('button','',label);
          button.type='button'; button.addEventListener('click',action); actions.append(button);
        };
        if(!data.coverage.scan_complete) addAction('Open Settings',()=>$('open-settings').click());
        if(data.coverage.undated_inwards) addAction('Check entries',()=>review({status:'needs_inward_date'}));
        if(data.coverage.folder_needs_details) addAction('View saved PDFs',files);
        const message=$('analytics-message');
        message.replaceChildren(node('span','',notes.join(' ')));
        if(actions.childElementCount) message.append(actions);
      }
    } catch(error) {
      if(token!==request || error.name==='AbortError') return;
      $('analytics-content').hidden=true; $('analytics-message').hidden=false;
      const retry=node('button','','Try again'); retry.type='button'; retry.addEventListener('click',load);
      $('analytics-message').replaceChildren(node('span','',error.message),retry);
    } finally {
      if(token===request) { $('analytics-workspace').setAttribute('aria-busy','false'); $('analytics-content').inert=false; }
    }
  }
  function bar(value,maximum) {
    const track=node('div','analytics-bar-track'),fill=node('div','analytics-bar-fill');
    fill.style.setProperty('--bar-width',`${maximum?100*value/maximum:0}%`); track.append(fill); return track;
  }
  function render(data) {
    const count=data.totals.generated,previous=data.totals.previous_generated;
    set('analytics-range',`${dateLabel(data.range.start)} – ${dateLabel(data.range.end)}`);
    set('analytics-created-label',`Certificates created ${periods[$('analytics-period').value]}`); set('analytics-created',number(count));
    const days=data.series.length;
    set('analytics-comparison',`${days===1?'Yesterday':`Previous ${days} days`}: ${number(previous)}`);
    $('analytics-comparison').title=`${dateLabel(data.range.previous_start)} – ${dateLabel(data.range.previous_end)}`;
    set('analytics-today',number(data.calendar.today)); set('analytics-week',number(data.calendar.week));
    set('analytics-received',number(data.totals.received)); set('analytics-received-period',`${$('analytics-period').selectedOptions[0].textContent} · By received date`);
    const time=data.totals.median_turnaround;
    set('analytics-turnaround',time==null?'—':`${number(time)} ${time===1?'day':'days'}`);
    set('analytics-time-sample',time==null?'No linked inwards yet':`${number(data.totals.turnaround_sample)} certificates`);
    set('analytics-type-total',`${number(count)} total`);
    const maximum=Math.max(1,...data.types.map(item=>item.count));
    $('analytics-types').replaceChildren(...data.types.map(item=>{
      const row=node('div',`analytics-bar ${item.key}`),heading=node('div','analytics-bar-top');
      row.tabIndex=0; row.setAttribute('role','img');
      row.setAttribute('aria-label',`${item.label}: ${number(item.count)} certificates.`);
      heading.append(node('span','',item.label),node('strong','',number(item.count)));
      row.append(heading,bar(item.count,maximum));
      attachTooltip(row,()=>({title:item.label,rows:[{label:'Certificates created',value:number(item.count)},
        {label:'Share of total',value:`${count?Math.round(100*item.count/count):0}% of ${number(count)}`}],caption:`${dateLabel(data.range.start)}–${dateLabel(data.range.end)}`}));
      return row;
    }));
    $('analytics-plants').classList.toggle('single-plant',data.plants.length===1);
    $('analytics-plants').replaceChildren(...data.plants.map(item=>{
      const button=node('button','analytics-plant'); button.type='button';
      const active=$('analytics-plant').value===item.key,action=active?'Show both plants':`Show ${item.label}`;
      button.setAttribute('aria-label',`${item.label}: ${number(item.generated)} certificates. ${action}.`);
      button.setAttribute('aria-pressed',String(active));
      const footer=node('div','analytics-plant-share'),share=count?Math.round(100*item.generated/count):0;
      footer.append(bar(item.generated,count),node('span','',`${share}%`));
      button.append(node('span','analytics-plant-name',item.label),node('strong','',number(item.generated)),footer,node('span','analytics-plant-action',`${active?'Show both plants':'View plant'} →`));
      attachTooltip(button,()=>({title:item.label,rows:[{label:'Certificates created',value:number(item.generated)},
        {label:'Share of total',value:`${share}% of ${number(count)}`},{label:'Client entries',value:number(item.received)}],
        caption:`${dateLabel(data.range.start)}–${dateLabel(data.range.end)} · Click to ${active?'show both plants':`show ${item.label}`}`}));
      button.addEventListener('click',()=>{ $('analytics-plant').value=active?'all':item.key; load(); }); return button;
    }));
    $('analytics-unassigned').hidden=!data.coverage.no_plant;
    set('analytics-unassigned',`${number(data.coverage.no_plant)} saved ${data.coverage.no_plant===1?'certificate has':'certificates have'} no plant listed.`);
    set('analytics-overdue-count',number(data.queue.overdue)); $('analytics-overdue').classList.toggle('has-waiting',data.queue.overdue>0);
    const definitions=[['ready','Ready to create','ready'],['to_send','PDFs to send','awaiting_send'],['attention','Check details','pending_review'],['destruction','HDD certificates','destruction_pending']];
    $('analytics-queue').replaceChildren(...definitions.map(([key,label,status])=>{
      const button=node('button',`analytics-work-card ${key}`); button.type='button';
      button.setAttribute('aria-label',`${label}: ${number(data.queue[key])}. View entries.`);
      button.append(node('span','analytics-work-label',label),node('strong','',number(data.queue[key])));
      button.addEventListener('click',()=>review({status,...(key==='destruction'?{type:'destruction'}:{})})); return button;
    }));
    drawChart(); set('analytics-updated',`Updated ${new Date().toLocaleTimeString(undefined,{hour:'numeric',minute:'2-digit'})}`);
  }
  function drawChart() {
    if(!current || $('analytics-workspace').hidden) return;
    const host=$('analytics-chart'),points=current.series || [],width=host.clientWidth || 600,height=host.clientHeight || 248;
    const left=45,right=18,top=22,bottom=36,plotWidth=width-left-right,plotHeight=height-top-bottom;
    const greatest=Math.max(1,...points.map(item=>item.generated));
    const raw=greatest/3,magnitude=10**Math.floor(Math.log10(raw));
    const step=Math.max(1,[1,2,5,10].find(value=>value*magnitude>=raw)*magnitude),ceiling=Math.ceil(greatest/step)*step;
    const x=index=>left+plotWidth*(points.length===1?.5:index/(points.length-1)),y=value=>top+plotHeight*(1-value/ceiling);
    const picture=svg('svg',{viewBox:`0 0 ${width} ${height}`,role:'img','aria-label':`${number(current.totals.generated)} certificates created, ${dateLabel(current.range.start)} to ${dateLabel(current.range.end)}. Each point shows one day.`});
    for(let value=0;value<=ceiling;value+=step) picture.append(svg('line',{x1:left,x2:width-right,y1:y(value),y2:y(value),stroke:'#e6ece5'}),svg('text',{x:left-10,y:y(value)+5,'text-anchor':'end'},number(value)));
    const pathFor=key=>points.map((item,i)=>`${i?'L':'M'}${x(i)},${y(item[key] || 0)}`).join(' ');
    if(points.length) {
      picture.append(svg('path',{d:`${pathFor('generated')} L${x(points.length-1)},${y(0)} L${x(0)},${y(0)} Z`,fill:'#edf5ef'}),svg('path',{d:pathFor('generated'),fill:'none',stroke:'#236c50','stroke-width':3,'stroke-linejoin':'round','stroke-linecap':'round'}));
      const labelCount=Math.min(points.length,Math.max(2,Math.floor(plotWidth/110)));
      const indices=new Set(Array.from({length:labelCount},(_,i)=>Math.round(i*(points.length-1)/Math.max(1,labelCount-1))));
      for(const index of indices) picture.append(svg('text',{x:x(index),y:height-9,'text-anchor':index===0?'start':index===points.length-1?'end':'middle'},dateLabel(points[index].date)));
      if(points.length===1) picture.append(svg('circle',{cx:x(0),cy:y(points[0].generated),r:5,fill:'#236c50'}));
    }
    const marker=svg('g',{'aria-hidden':'true'}),selected=Math.min(Math.max(0,chartIndex<0?points.length-1:chartIndex),points.length-1);
    function select(index) {
      chartIndex=index; const item=points[index]; if(!item) return;
      marker.replaceChildren(svg('line',{x1:x(index),x2:x(index),y1:top,y2:y(0),stroke:'#9cb5a3','stroke-dasharray':'3 4'}),
        svg('circle',{cx:x(index),cy:y(item.generated),r:6,fill:'#236c50',stroke:'#fff','stroke-width':2}));
      set('analytics-chart-detail',`${fullDateLabel(item.date)}: ${number(item.generated)} ${item.generated===1?'certificate':'certificates'} created`);
    }
    function showPoint(anchor) {
      const item=points[chartIndex]; if(!item) return;
      const bounds=host.getBoundingClientRect();
      showTooltip(host,fullDateLabel(item.date),[
        {label:'Certificates created',value:number(item.generated)},
      ],null,anchor || {x:bounds.left+x(chartIndex),y:bounds.top+y(item.generated)});
    }
    picture.append(marker); host.replaceChildren(picture); select(selected);
    host.onpointerenter=()=>{ dismissedOwner=null; };
    host.onpointermove=event=>{ const bounds=host.getBoundingClientRect(); select(Math.max(0,Math.min(points.length-1,Math.round((event.clientX-bounds.left-left)/plotWidth*Math.max(1,points.length-1))))); showPoint({x:event.clientX,y:event.clientY}); };
    host.onpointerleave=()=>{ if(document.activeElement!==host) hideTooltip(host); dismissedOwner=null; };
    host.onfocus=()=>{ dismissedOwner=null; showPoint(); };
    host.onblur=()=>hideTooltip(host);
    host.onkeydown=event=>{
      if(!['ArrowLeft','ArrowRight','Home','End'].includes(event.key)) return; event.preventDefault();
      select(event.key==='Home'?0:event.key==='End'?points.length-1:Math.max(0,Math.min(points.length-1,chartIndex+(event.key==='ArrowRight'?1:-1))));
      dismissedOwner=null; showPoint();
    };
  }
  document.addEventListener('DOMContentLoaded',()=>{
    $('analytics-plant').addEventListener('change',load); $('analytics-period').addEventListener('change',load);
    $('analytics-overdue').addEventListener('click',()=>review({status:'overdue_unissued'}));
    const refreshVisible=()=>{ if(!document.hidden && !$('analytics-workspace').hidden && !document.querySelector('dialog[open]')) load(); };
    window.setInterval(refreshVisible,60000); window.addEventListener('focus',refreshVisible);
    new ResizeObserver(drawChart).observe($('analytics-chart'));
    document.addEventListener('keydown',event=>{ if(event.key==='Escape' && tooltipOwner) { dismissedOwner=tooltipOwner; hideTooltip(); } });
    window.addEventListener('scroll',()=>{
      if(tooltipOwner && !tooltipOwner.matches(':hover')) hideTooltip();
    },{capture:true,passive:true});
    window.addEventListener('resize',()=>hideTooltip());
    window.addEventListener('blur',()=>hideTooltip());
    document.addEventListener('pointerdown',event=>{ if(tooltipOwner && !tooltipOwner.contains(event.target)) hideTooltip(); });
  });
  window.CertificateAnalytics={load};
})();
