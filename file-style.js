(function(){
  'use strict';
  const C=window.FileStyleCore;
  let palette=C.clean(null), ready=false, busy=false;
  const $=id=>document.getElementById(id);
  function status(message,type=''){ $('status').textContent=message;$('status').className=type; }
  function controls(){document.querySelectorAll('button').forEach(b=>b.disabled=!ready||busy);}
  async function action(fn){
    if(!ready||busy)return;
    busy=true;controls();
    try{await fn();}catch(error){console.error(error);status(error.message||'処理に失敗しました。','error');}
    finally{busy=false;controls();}
  }
  function render(){
    for(const kind of ['colors','fonts']){
      const list=$(kind);list.replaceChildren();
      if(!palette[kind].length){const p=document.createElement('p');p.className='hint';p.textContent='まだ登録されていません。';list.append(p);}
      palette[kind].forEach((item,index)=>{
        const entry=document.createElement('div');entry.className='entry';
        const apply=document.createElement(kind==='colors'?'div':'button');apply.className=kind==='colors'?'color-card':'apply';
        if(kind==='fonts')apply.title='選択したオブジェクトに適用';
        if(kind==='colors'){const swatch=document.createElement('span');swatch.className='swatch';swatch.style.backgroundColor=item.value;apply.append(swatch);}
        const text=document.createElement('span');text.className='label';text.textContent=item.label;
        const value=document.createElement('small');value.textContent=item.value;text.append(value);apply.append(text);
        if(kind==='colors'){
          const actions=document.createElement('div');actions.className='color-actions';
          for(const [target,label,icon] of [['text','文字','T'],['fill','塗り','■'],['line','線','□']]){
            const button=document.createElement('button');button.className='color-action';
            button.title=`${item.value}を${label}に適用`;button.setAttribute('aria-label',`${item.label}を${label}に適用`);
            const symbol=document.createElement('span');symbol.className='action-icon';symbol.textContent=icon;symbol.setAttribute('aria-hidden','true');
            const caption=document.createElement('span');caption.textContent=label;button.append(symbol,caption);
            button.addEventListener('click',()=>action(()=>applyStyle(kind,item.value,target)));actions.append(button);
          }
          apply.append(actions);
        }else apply.addEventListener('click',()=>action(()=>applyStyle(kind,item.value)));
        const remove=document.createElement('button');remove.className='delete';remove.textContent='削除';remove.setAttribute('aria-label',item.label+'を削除');
        remove.addEventListener('click',()=>action(async()=>{await refresh();const next=C.clean(palette);const at=next[kind].findIndex(x=>x.value===item.value);if(at>=0)next[kind].splice(at,1);await persist(next);}));
        entry.append(apply,remove);list.append(entry);
      });
    }
    const size=C.bytes(palette);$('capacity').textContent=`登録データ：約 ${(size/1024).toFixed(1)} KB（設定JSON／PPTX全体の増分ではありません）`;
    controls();
  }
  function callback(method){return new Promise((resolve,reject)=>method(result=>result.status===Office.AsyncResultStatus.Succeeded?resolve():reject(new Error(result.error?.message||'ファイル内の登録情報にアクセスできませんでした。'))));}
  async function refresh(){
    await callback(cb=>Office.context.document.settings.refreshAsync(cb));
    palette=C.clean(Office.context.document.settings.get(C.key));render();
  }
  async function persist(next){
    const size=C.bytes(next);if(size>C.limit)throw new Error('登録データが64 KBを超えました。保存を停止しました。');
    const settings=Office.context.document.settings,previous=settings.get(C.key);
    settings.set(C.key,next);
    try{await callback(cb=>settings.saveAsync({overwriteIfStale:false},cb));}
    catch(error){if(previous==null)settings.remove(C.key);else settings.set(C.key,previous);throw new Error('登録情報を保存できませんでした。再読み込み後に再試行してください。 '+error.message);}
    palette=next;render();
    status(size>=C.warning?'登録データが32 KBを超えました。容量を確認してください。閉じる前に ⌘S で保存してください。':'ファイル内に登録しました。閉じる前に ⌘S で保存してください。',size>=C.warning?'warning':'success');
  }
  async function register(kind,value,label){await refresh();await persist(C.add(palette,kind,value,label));}
  async function textTargets(context,shapes){
    const selected=context.presentation.getSelectedTextRangeOrNullObject();selected.load('text');await context.sync();
    if(!selected.isNullObject&&selected.text.length>0)return [selected];
    const frames=shapes.map(s=>s.getTextFrameOrNullObject());frames.forEach(f=>f.load('hasText'));await context.sync();
    return frames.filter(f=>!f.isNullObject&&f.hasText).map(f=>f.textRange);
  }
  async function selection(context){const selected=context.presentation.getSelectedShapes();selected.load('items/type');await context.sync();return selected.items;}
  const supported=s=>['GeometricShape','TextBox','Line'].includes(s.type);
  async function capture(kind){
    const values=await PowerPoint.run(async context=>{
      const shapes=await selection(context),found=[];
      if(kind==='colors'){
        const objects=shapes.filter(supported);
        objects.forEach(s=>{if(s.type!=='Line')s.fill.load('type,foregroundColor');s.lineFormat.load('visible,color');});
        await context.sync();
        objects.forEach(s=>{if(s.type!=='Line'&&s.fill.type==='Solid')found.push(s.fill.foregroundColor);if(s.lineFormat.visible)found.push(s.lineFormat.color);});
      }
      {
        const ranges=await textTargets(context,shapes);if(!ranges.length&&kind==='fonts')throw new Error('文字のあるオブジェクト、または文字範囲を選択してください。');
        const property=kind==='fonts'?'name':'color';ranges.forEach(r=>r.font.load(property));await context.sync();
        let chars=0;const fragments=[];
        for(const range of ranges){
          if(range.font[property])found.push(range.font[property]);
          else {range.load('text');fragments.push(range);}
        }
        if(fragments.length){
          await context.sync();const parts=[];
          for(const range of fragments){chars+=range.text.length;if(chars>4000)throw new Error('書式が混在した文字が4,000文字を超えています。登録したい部分を範囲選択してください。');
            for(let i=0;i<range.text.length;i++){const p=range.getSubstring(i,1);p.font.load(property);parts.push(p);}}
          await context.sync();parts.forEach(p=>{if(p.font[property])found.push(p.font[property]);});
        }
      }
      return [...new Set(found.map(v=>kind==='colors'?C.color(v):v).filter(Boolean))];
    });
    if(!values.length)throw new Error('登録できる単色またはフォントがありません。単色のオブジェクトを選ぶか、手入力してください。');
    await refresh();let next=palette;values.forEach(v=>{next=C.add(next,kind,v,kind==='colors'?$('color-name').value:$('font-label').value);});await persist(next);
  }
  async function applyStyle(kind,value,target){
    const result=await PowerPoint.run(async context=>{
      const shapes=await selection(context);let count=0,skipped=0;
      if(kind==='fonts'||target==='text'){
        const ranges=await textTargets(context,shapes);if(!ranges.length)throw new Error('文字のあるオブジェクト、または文字範囲を選択してください。');
        for(const range of ranges){if(kind==='fonts')range.font.name=value;else range.font.color=value;count++;}
        skipped=Math.max(0,shapes.length-ranges.length);
      }else{
        for(const shape of shapes){if(!supported(shape)||(target==='fill'&&shape.type==='Line')){skipped++;continue;}
          if(target==='fill')shape.fill.setSolidColor(value);else {shape.lineFormat.color=value;shape.lineFormat.visible=true;}count++;
        }
        if(!count)throw new Error('適用できる図形が選択されていません。');
      }
      await context.sync();return {count,skipped};
    });
    status(`${result.count}件に適用しました。${result.skipped?` 対象外の${result.skipped}件はスキップしました。`:''} ⌘Z で戻せます。`,'success');
  }
  document.addEventListener('DOMContentLoaded',()=>{
    $('picker').addEventListener('input',()=>{$('hex').value=$('picker').value.toUpperCase();});
    $('hex').addEventListener('input',()=>{const color=C.color($('hex').value);if(color)$('picker').value=color;});
    $('color-form').addEventListener('submit',e=>{e.preventDefault();action(()=>register('colors',$('hex').value,$('color-name').value));});
    $('font-form').addEventListener('submit',e=>{e.preventDefault();action(()=>register('fonts',$('font-name').value,$('font-label').value));});
    $('capture-color').addEventListener('click',()=>action(()=>capture('colors')));
    $('capture-font').addEventListener('click',()=>action(()=>capture('fonts')));
    $('reload').addEventListener('click',()=>action(async()=>{await refresh();status('このファイルの登録情報を読み込みました。','success');}));
    Office.onReady(async info=>{
      try{
        if(info.host!==Office.HostType.PowerPoint)throw new Error('PowerPointのアドインとして開いてください。');
        if(!Office.context.requirements.isSetSupported('PowerPointApi','1.8'))throw new Error('PowerPointApi 1.8以上が必要です。');
        palette=C.clean(Office.context.document.settings.get(C.key));ready=true;render();
        status('このファイルの色とフォントを登録できます。');
        Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged,()=>{if(!busy)action(async()=>{await refresh();status('登録情報を更新しました。');});},result=>{if(result.status!==Office.AsyncResultStatus.Succeeded)console.info('SettingsChanged unavailable; manual refresh supported.');});
      }catch(e){status(e.message,'error');}
    }).catch(e=>status(e.message,'error'));
  });
})();
