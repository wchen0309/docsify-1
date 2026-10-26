(function(){
  function parseDate(value){
    if(!value) return null;
    if(value instanceof Date && !isNaN(value)) return new Date(value.getFullYear(),value.getMonth(),value.getDate());
    const s=String(value).trim();
    if(!s || s.startsWith('=')) return null;
    const m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if(m){
      const d=new Date(Number(m[1]),Number(m[2])-1,Number(m[3]));
      return isNaN(d)?null:d;
    }
    const d=new Date(s);
    return isNaN(d)?null:new Date(d.getFullYear(),d.getMonth(),d.getDate());
  }
  function dayDiff(a,b){ return Math.round((b-a)/86400000); }
  function fmt(d){ return (d.getMonth()+1)+'/'+d.getDate(); }
  function esc2(v){ return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
  function categoryOf(t){ return typeof canonicalCategory==='function'?canonicalCategory(t.category):String(t.category||''); }

  window.ganttZoomMode=window.ganttZoomMode||'week';
  window.ganttZoomFactor=window.ganttZoomFactor||1;
  window.setGanttZoom=function(mode){
    window.ganttZoomMode=mode;
    window.ganttZoomFactor=1;
    if(typeof renderDetail==='function') renderDetail();
  };
  window.adjustGanttZoom=function(delta){
    window.ganttZoomFactor=Math.max(0.55,Math.min(2.4,Math.round((window.ganttZoomFactor+delta)*100)/100));
    if(typeof renderDetail==='function') renderDetail();
  };

  window.tlDate=function(v){
    if(v===null||v===undefined||v==='') return '';
    if(v instanceof Date&&!isNaN(v)) return v.toISOString().slice(0,10);
    if(typeof v==='number'&&v>20000&&v<80000&&window.XLSX){
      const d=XLSX.SSF.parse_date_code(v);
      if(d) return String(d.y).padStart(4,'0')+'-'+String(d.m).padStart(2,'0')+'-'+String(d.d).padStart(2,'0');
    }
    const s=String(v).trim();
    if(!s || s.startsWith('=')) return '';
    const m=s.match(/^(\d{4})[./年-](\d{1,2})[./月-](\d{1,2})/);
    if(m){
      const iso=m[1]+'-'+m[2].padStart(2,'0')+'-'+m[3].padStart(2,'0');
      return parseDate(iso)?iso:'';
    }
    return '';
  };

  const style=document.createElement('style');
  style.textContent=`
    .gantt-toolbar{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin:0 0 10px}
    .gantt-toolbar .gantt-label{font-size:12px;color:var(--muted);margin-right:2px}
    .gantt-toolbar button{padding:7px 10px;border-radius:9px;font-size:12px}
    .gantt-toolbar button.active{background:#252b40;color:#fff;box-shadow:none}
    .gantt-zoom-pct{font-size:12px;color:var(--muted);min-width:42px;text-align:center}
    .gantt-scroll{overflow:auto;border:1px solid var(--line);border-radius:13px;background:#fff;max-height:72vh}
    .gantt-grid{--cat-w:96px;--task-w:330px;--owner-w:110px;min-width:max-content;font-size:12px}
    .gantt-row{display:grid;align-items:stretch;border-bottom:1px solid var(--line);min-height:52px}
    .gantt-head{position:sticky;top:0;z-index:6;background:#f8f9fb;min-height:66px;font-weight:700}
    .gantt-cat,.gantt-task,.gantt-owner{position:sticky;background:#fff;border-right:1px solid var(--line);padding:10px;display:flex;align-items:center;z-index:3}
    .gantt-head .gantt-cat,.gantt-head .gantt-task,.gantt-head .gantt-owner{background:#f8f9fb;z-index:8}
    .gantt-cat{left:0;width:var(--cat-w);font-weight:700;color:var(--accent)}
    .gantt-task{left:var(--cat-w);width:var(--task-w);display:block;z-index:4}
    .gantt-owner{left:calc(var(--cat-w) + var(--task-w));width:var(--owner-w);z-index:5;color:#475467}
    .gantt-task b{display:block;font-size:13px;line-height:1.35;margin-bottom:3px;color:var(--ink)}
    .gantt-task small{display:block;color:var(--muted);line-height:1.3;max-width:300px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .gantt-time{position:relative;background:#fff;overflow:hidden}
    .gantt-head .gantt-time{background:#f8f9fb;overflow:visible}
    .gantt-months{height:30px;position:relative;border-bottom:1px solid var(--line)}
    .gantt-month{position:absolute;top:0;height:30px;padding:7px 8px;border-right:1px solid var(--line);white-space:nowrap;font-weight:700}
    .gantt-ticks{height:35px;position:relative}
    .gantt-tick{position:absolute;top:0;height:35px;border-right:1px solid #edf0f5;padding:6px 3px;text-align:center;color:#667085;white-space:nowrap;font-size:11px;overflow:hidden}
    .gantt-gridline{position:absolute;top:0;bottom:0;border-right:1px solid #f0f2f6}
    .gantt-bar{position:absolute;top:14px;height:24px;border-radius:7px;background:var(--soft);box-shadow:inset 0 0 0 1px #d9dcff;overflow:hidden;min-width:4px}
    .gantt-bar i{display:block;height:100%;background:var(--accent);opacity:.82}
    .gantt-bar span{position:absolute;inset:0;display:flex;align-items:center;padding:0 7px;font-size:11px;color:#20263a;white-space:nowrap}
    .gantt-empty{position:absolute;left:10px;top:17px;color:var(--muted)}
    .gantt-today{position:absolute;top:0;bottom:0;width:2px;background:#ef6a5b;opacity:.65;z-index:2}
    @media(max-width:760px){.gantt-scroll{max-height:68vh}.gantt-grid{--cat-w:72px;--task-w:230px;--owner-w:88px}.gantt-cat,.gantt-task,.gantt-owner{padding:8px}.gantt-task small{max-width:205px}.gantt-row{min-height:58px}.gantt-toolbar{position:sticky;left:0}}
  `;
  document.head.appendChild(style);

  function zoomToolbar(){
    const mode=window.ganttZoomMode||'week';
    const pct=Math.round((window.ganttZoomFactor||1)*100);
    return '<div class="gantt-toolbar"><span class="gantt-label">时间尺度</span>'+
      ['day','week','month'].map(m=>'<button class="'+(mode===m?'active':'')+'" onclick="setGanttZoom(\''+m+'\')">'+({day:'日',week:'周',month:'月'}[m])+'</button>').join('')+
      '<span style="width:6px"></span><button onclick="adjustGanttZoom(-0.15)" title="缩小">−</button><span class="gantt-zoom-pct">'+pct+'%</span><button onclick="adjustGanttZoom(0.15)" title="放大">＋</button></div>';
  }

  function gridCols(timeWidth){
    return 'var(--cat-w) var(--task-w) var(--owner-w) '+timeWidth+'px';
  }

  window.gantt=function(ts){
    const rows=(ts||[]).slice();
    const toolbar=zoomToolbar();
    const scheduled=rows.map(t=>({t,start:parseDate(t.planStart),end:parseDate(t.planEnd)})).filter(x=>x.start&&x.end&&x.end>=x.start);
    if(!scheduled.length){
      const cols=gridCols(520);
      return toolbar+'<div class="gantt-scroll"><div class="gantt-grid">'+
        '<div class="gantt-row gantt-head" style="grid-template-columns:'+cols+'"><div class="gantt-cat">类别</div><div class="gantt-task">任务描述</div><div class="gantt-owner">负责人</div><div class="gantt-time"><div class="gantt-empty">暂无有效计划日期</div></div></div>'+
        rows.map(t=>'<div class="gantt-row" style="grid-template-columns:'+cols+'" onclick="editTask(\''+esc2(t.id)+'\')"><div class="gantt-cat">'+esc2(categoryOf(t))+'</div><div class="gantt-task"><b>'+esc2(t.name||'未命名任务')+'</b><small>'+esc2(t.desc||'')+'</small></div><div class="gantt-owner">'+esc2(t.owner||'未分配')+'</div><div class="gantt-time"><div class="gantt-empty">未排期</div></div></div>').join('')+
        '</div></div>';
    }

    let min=new Date(Math.min(...scheduled.map(x=>+x.start))), max=new Date(Math.max(...scheduled.map(x=>+x.end)));
    min=new Date(min.getFullYear(),min.getMonth(),min.getDate());
    max=new Date(max.getFullYear(),max.getMonth(),max.getDate());
    const totalDays=dayDiff(min,max)+1;
    const mode=window.ganttZoomMode||'week';
    const factor=window.ganttZoomFactor||1;
    const baseScale=mode==='day'?28:mode==='week'?12:4;
    const scale=baseScale*factor;
    const timeWidth=Math.max(620,totalDays*scale);
    const cols=gridCols(timeWidth);

    let tickHtml='',gridHtml='';
    if(mode==='month'){
      let c=new Date(min.getFullYear(),min.getMonth(),1);
      while(c<=max){
        const start=new Date(Math.max(+c,+min));
        const next=new Date(c.getFullYear(),c.getMonth()+1,1);
        const end=new Date(Math.min(+new Date(next.getTime()-86400000),+max));
        const left=dayDiff(min,start)*scale;
        const width=(dayDiff(start,end)+1)*scale;
        tickHtml+='<div class="gantt-tick" style="left:'+left+'px;width:'+width+'px">'+(c.getMonth()+1)+'月</div>';
        gridHtml+='<i class="gantt-gridline" style="left:'+left+'px"></i>';
        c=next;
      }
    }else{
      const step=mode==='day'?1:7;
      for(let i=0;i<totalDays;i+=step){
        const d=new Date(min); d.setDate(d.getDate()+i);
        const w=Math.min(step,totalDays-i)*scale;
        const label=mode==='day'?String(d.getDate()):fmt(d);
        tickHtml+='<div class="gantt-tick" style="left:'+(i*scale)+'px;width:'+w+'px">'+label+'</div>';
        gridHtml+='<i class="gantt-gridline" style="left:'+(i*scale)+'px"></i>';
      }
    }

    let monthHtml='';
    let cursor=new Date(min.getFullYear(),min.getMonth(),1);
    while(cursor<=max){
      const ms=new Date(Math.max(+cursor,+min));
      const next=new Date(cursor.getFullYear(),cursor.getMonth()+1,1);
      const me=new Date(Math.min(+new Date(next.getTime()-86400000),+max));
      const left=dayDiff(min,ms)*scale;
      const width=(dayDiff(ms,me)+1)*scale;
      monthHtml+='<div class="gantt-month" style="left:'+left+'px;width:'+width+'px">'+ms.getFullYear()+'年'+(ms.getMonth()+1)+'月</div>';
      cursor=next;
    }

    const today=new Date(); today.setHours(0,0,0,0);
    const todayOffset=today>=min&&today<=max?dayDiff(min,today)*scale:null;
    const todayHtml=todayOffset===null?'':'<i class="gantt-today" style="left:'+todayOffset+'px" title="今天"></i>';

    const head='<div class="gantt-row gantt-head" style="grid-template-columns:'+cols+'"><div class="gantt-cat">类别</div><div class="gantt-task">任务描述</div><div class="gantt-owner">负责人</div><div class="gantt-time"><div class="gantt-months">'+monthHtml+'</div><div class="gantt-ticks">'+tickHtml+'</div></div></div>';

    const body=rows.map(t=>{
      const start=parseDate(t.planStart),end=parseDate(t.planEnd);
      let timeline=gridHtml+todayHtml;
      if(start&&end&&end>=start){
        const left=dayDiff(min,start)*scale;
        const width=Math.max(5,(dayDiff(start,end)+1)*scale);
        const pct=Math.max(0,Math.min(100,Number(t.completion)||0));
        timeline+='<div class="gantt-bar" style="left:'+left+'px;width:'+width+'px" title="'+esc2((t.planStart||'')+' 至 '+(t.planEnd||''))+'"><i style="width:'+pct+'%"></i><span>'+esc2(t.planStart||'')+' → '+esc2(t.planEnd||'')+'</span></div>';
      }else{
        const bad=(t.planStart||t.planEnd)?'日期待修正':'未排期';
        timeline+='<div class="gantt-empty">'+bad+'</div>';
      }
      return '<div class="gantt-row" style="grid-template-columns:'+cols+'" onclick="editTask(\''+esc2(t.id)+'\')"><div class="gantt-cat">'+esc2(categoryOf(t))+'</div><div class="gantt-task"><b>'+esc2(t.name||'未命名任务')+'</b><small>'+esc2(t.desc||'')+'</small></div><div class="gantt-owner">'+esc2(t.owner||'未分配')+'</div><div class="gantt-time">'+timeline+'</div></div>';
    }).join('');

    return toolbar+'<div class="gantt-scroll"><div class="gantt-grid">'+head+body+'</div></div>';
  };
})();