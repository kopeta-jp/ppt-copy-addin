(function(root){
  'use strict';
  function create({status,action}){
    const $=id=>document.getElementById(id);
    let hits=[],font='',index=-1,scanning=false,cancelled=false,enabled=false;
    const matches=()=>hits.filter(h=>h.font.toLowerCase()===font.toLowerCase());
    function controls(allowed=enabled){
      enabled=allowed;
      $('audit-scope').disabled=!allowed;
      $('audit-scan').disabled=!allowed;
      $('audit-prev').disabled=!allowed||index<=0;
      $('audit-next').disabled=!allowed||index<0||index>=matches().length-1;
      $('audit-cancel').disabled=!scanning;$('audit-cancel').hidden=!scanning;
      $('audit-fonts').querySelectorAll('button').forEach(b=>b.disabled=!allowed);
    }
    function render(){
      const list=$('audit-fonts');list.replaceChildren();
      const fonts=new Map();for(const h of hits){const key=h.font.toLowerCase();if(!fonts.has(key))fonts.set(key,{name:h.font,count:0});fonts.get(key).count++;}
      for(const f of [...fonts.values()].sort((a,b)=>a.name.localeCompare(b.name,'ja'))){
        const button=document.createElement('button');button.className='audit-font';button.textContent=`${f.name} · ${f.count}箇所`;
        button.setAttribute('aria-pressed',String(font.toLowerCase()===f.name.toLowerCase()));
        button.addEventListener('click',()=>action(async()=>{font=f.name;index=-1;render();await navigate(0);}));list.append(button);
      }
      controls();
    }
    function reset(){hits=[];font='';index=-1;render();$('audit-position').textContent='';}
    async function scan(){
      reset();scanning=true;cancelled=false;controls(false);
      try{
        const result=await PowerPoint.run(async context=>{
          const slides=context.presentation.slides;slides.load('items/id');
          const selected=context.presentation.getSelectedSlides();selected.load('items/id');await context.sync();
          const ids=new Set(selected.items.map(s=>s.id));
          if($('audit-scope').value==='current'&&ids.size!==1)throw new Error('スライドを1枚だけ選択してください。');
          const result=[],excluded=new Set();
          function record(meta,name,start,length,text){
            if(!name||!text.trim())return;
            const prev=result[result.length-1];
            if(prev&&prev.slideId===meta.slideId&&prev.shapeId===meta.shapeId&&prev.row===meta.row&&prev.col===meta.col&&prev.font===name&&prev.start+prev.length===start){prev.length+=length;prev.text+=text;}
            else result.push({...meta,font:name,start,length,text});
          }
          for(let i=0;i<slides.items.length;i++){
            if(cancelled)throw new Error('チェックを中止しました。');
            const slide=slides.items[i];if($('audit-scope').value==='current'&&!ids.has(slide.id))continue;
            status(`フォントを確認中：スライド ${i+1}/${slides.items.length}`);
            slide.shapes.load('items/id,items/type');await context.sync();
            const objects=[];
            for(let shapeOrder=0;shapeOrder<slide.shapes.items.length;shapeOrder++){
              const shape=slide.shapes.items[shapeOrder];
              const meta={slideId:slide.id,slideNumber:i+1,shapeId:shape.id,shapeOrder};
              if(['Group','SmartArt','Chart'].includes(shape.type)){excluded.add(shape.type);continue;}
              if(shape.type==='Table'){
                if(!Office.context.requirements.isSetSupported('PowerPointApi','1.9')){excluded.add('Table');continue;}
                const table=shape.getTable();table.load('rowCount,columnCount');objects.push({meta,table});
              }else{
                const frame=shape.getTextFrameOrNullObject();frame.load('hasText');objects.push({meta,frame});
              }
            }
            await context.sync();const ranges=[],cells=[];
            for(const o of objects){
              if(o.table){for(let row=0;row<o.table.rowCount;row++)for(let col=0;col<o.table.columnCount;col++){
                const cell=o.table.getCellOrNullObject(row,col);cell.load('text');cells.push({meta:{...o.meta,row,col},cell});
              }}else if(!o.frame.isNullObject&&o.frame.hasText){const range=o.frame.textRange;range.load('text');range.font.load('name');ranges.push({meta:o.meta,range});}
            }
            await context.sync();const mixed=[];
            for(const o of ranges){
              const text=o.range.text;if(!text.trim())continue;
              if(o.range.font.name)record(o.meta,o.range.font.name,0,text.length,text);
              else mixed.push(o);
            }
            // Split only mixed ranges, batching each tree level instead of one sync per character.
            for(const o of mixed){
              let queue=[{start:0,length:o.range.text.length}],parts=[];
              while(queue.length){
                if(cancelled)throw new Error('チェックを中止しました。');
                const batch=queue.splice(0,256).map(p=>({...p,range:o.range.getSubstring(p.start,p.length)}));
                batch.forEach(p=>p.range.font.load('name'));await context.sync();
                for(const p of batch){
                  if(p.range.font.name)parts.push({...p,font:p.range.font.name});
                  else if(p.length>1){const n=Math.floor(p.length/2);queue.push({start:p.start,length:n},{start:p.start+n,length:p.length-n});}
                  else excluded.add('未取得の文字');
                }
              }
              for(const p of parts.sort((a,b)=>a.start-b.start))record(o.meta,p.font,p.start,p.length,o.range.text.slice(p.start,p.start+p.length));
            }
            const valid=cells.filter(o=>!o.cell.isNullObject&&o.cell.text.trim());
            valid.forEach(o=>{o.cell.font.load('name');o.cell.load('textRuns');});if(valid.length)await context.sync();
            for(const o of valid){
              if(o.cell.font.name)record(o.meta,o.cell.font.name,0,o.cell.text.length,o.cell.text);
              else{let start=0;for(const run of o.cell.textRuns){const text=run.text||'';record(o.meta,run.font?.name,start,text.length,text);if(!run.font?.name&&text.trim())excluded.add('未取得の表内文字');start+=text.length;}}
            }
          }
          return {hits:result.sort((a,b)=>a.slideNumber-b.slideNumber||a.shapeOrder-b.shapeOrder||(a.row??-1)-(b.row??-1)||(a.col??-1)-(b.col??-1)||a.start-b.start),excluded:[...excluded]};
        });
        if(cancelled)throw new Error('チェックを中止しました。');
        hits=result.hits;render();
        $('audit-position').textContent=result.excluded.length?`対象外：${result.excluded.join('・')}`:'';
        status(hits.length?`${new Set(hits.map(h=>h.font.toLowerCase())).size}種類のフォントを検出。`:'対象の文字がありません。',result.excluded.length?'warning':'success');
      }finally{scanning=false;controls(false);}
    }
    async function navigate(next){
      const list=matches();if(next<0||next>=list.length)return;
      const hit=list[next];
      try{
        await PowerPoint.run(async context=>{
          const slide=context.presentation.slides.getItem(hit.slideId);const shape=slide.shapes.getItem(hit.shapeId);
          if(hit.row!==undefined){
            const cell=shape.getTable().getCellOrNullObject(hit.row,hit.col);cell.load('text');await context.sync();
            if(cell.isNullObject||cell.text.slice(hit.start,hit.start+hit.length)!==hit.text)throw new Error('文字が変更されています。再チェックしてください。');
            context.presentation.setSelectedSlides([hit.slideId]);await context.sync();slide.setSelectedShapes([hit.shapeId]);await context.sync();
          }else{
            const frame=shape.getTextFrameOrNullObject();frame.load('hasText');await context.sync();
            if(frame.isNullObject||!frame.hasText)throw new Error('文字が変更されています。再チェックしてください。');
            const full=frame.textRange;full.load('text');await context.sync();
            if(full.text.slice(hit.start,hit.start+hit.length)!==hit.text)throw new Error('文字が変更されています。再チェックしてください。');
            const range=full.getSubstring(hit.start,hit.length);range.font.load('name');await context.sync();
            if((range.font.name||'').toLowerCase()!==font.toLowerCase())throw new Error('フォントが変更されています。再チェックしてください。');
            context.presentation.setSelectedSlides([hit.slideId]);await context.sync();range.setSelected();await context.sync();
          }
        });
      }catch(e){throw new Error(`${e.message||'使用箇所を選択できません。'} 必要に応じて再チェックしてください。`);}
      index=next;const sample=hit.text.replace(/\s+/g,' ').slice(0,60);
      $('audit-position').textContent=`${next+1}/${list.length} · スライド ${hit.slideNumber}${hit.row!==undefined?` · 表 ${hit.row+1}行${hit.col+1}列（外枠選択）`:''}\n${sample}`;
      render();status(`${font}の使用箇所を選択。`,'success');
    }
    $('audit-scan').addEventListener('click',()=>action(scan));
    $('audit-prev').addEventListener('click',()=>action(()=>navigate(index-1)));
    $('audit-next').addEventListener('click',()=>action(()=>navigate(index+1)));
    $('audit-cancel').addEventListener('click',()=>{cancelled=true;status('中止しています…');});
    $('audit-scope').addEventListener('change',reset);
    return {controls};
  }
  root.FileStyleAudit={create};
})(window);
