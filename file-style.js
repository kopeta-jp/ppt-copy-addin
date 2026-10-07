(function(){
  'use strict';
  const C=window.FileStyleCore;
  let palette=C.clean(null), ready=false, busy=false, exportUrl=null;
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
    for(const kind of ['colors','fonts','sizes']){
      const list=$(kind);list.replaceChildren();
      if(!palette[kind].length){const p=document.createElement('p');p.className='hint';p.textContent='未登録';list.append(p);}
      palette[kind].forEach((item,index)=>{
        const entry=document.createElement('div');entry.className='entry';
        const apply=document.createElement(kind==='colors'?'div':'button');apply.className=kind==='colors'?'color-card':'apply';
        if(kind!=='colors')apply.title='選択したオブジェクトに適用';
        if(kind==='colors'){const swatch=document.createElement('span');swatch.className='swatch';swatch.style.backgroundColor=item.value;apply.append(swatch);}
        const text=document.createElement('span');text.className='label';text.textContent=item.label;
        const value=document.createElement('small');value.textContent=kind==='sizes'?`${item.value} pt`:item.value;
        if(kind!=='sizes'||item.label!==`${item.value} pt`)text.append(value);apply.append(text);
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
    const size=C.bytes(palette);$('capacity').textContent=`登録データ：約 ${(size/1024).toFixed(1)} KB`;
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
  async function exportStyles(){
    await refresh();
    if(exportUrl)URL.revokeObjectURL(exportUrl);
    const text=C.exportStyles(palette);
    exportUrl=URL.createObjectURL(new Blob([text],{type:'application/json;charset=utf-8'}));
    const link=$('download-styles');link.href=exportUrl;link.download='file-style-styles.json';
    $('export-result').hidden=false;$('export-json').value=text;
    link.click();
    status('書き出しました。','success');
  }
  async function importText(text){
    const incoming=C.parseStyles(text);
    await refresh();
    const merged=C.mergeStyles(palette,incoming);
    const a=merged.added;
    if(a.colors+a.fonts+a.sizes===0){status('すべて登録済みでした。登録情報は変更していません。','success');return;}
    await persist(merged.palette);
    const large=C.bytes(merged.palette)>=C.warning;
    status(`読み込みました：色 ${a.colors}件・フォント ${a.fonts}件・サイズ ${a.sizes}件を追加。${large?'登録データが32 KBを超えています。容量を確認してください。':''}閉じる前に ⌘S で保存してください。`,large?'warning':'success');
  }
  async function textTargets(context,shapes){
    const selected=context.presentation.getSelectedTextRangeOrNullObject();selected.load('text');await context.sync();
    if(!selected.isNullObject&&selected.text.length>0)return [selected];
    const tableShapes=shapes.filter(s=>s.type==='Table');
    if(tableShapes.length&&!Office.context.requirements.isSetSupported('PowerPointApi','1.9'))throw new Error('表内の文字変更にはPowerPointApi 1.9以上が必要です。PowerPointを更新してください。');
    const frames=shapes.filter(s=>s.type!=='Table').map(s=>s.getTextFrameOrNullObject());frames.forEach(f=>f.load('hasText'));
    const tables=tableShapes.map(s=>s.getTable());tables.forEach(t=>t.load('rowCount,columnCount'));await context.sync();
    const ranges=frames.filter(f=>!f.isNullObject&&f.hasText).map(f=>f.textRange);
    ranges.coveredCount=ranges.length+tables.length;
    const cells=[];
    for(const table of tables){
      for(let row=0;row<table.rowCount;row++)for(let col=0;col<table.columnCount;col++){
        const cell=table.getCellOrNullObject(row,col);cell.load('text');cells.push(cell);
      }
    }
    if(cells.length)await context.sync();
    const validCells=cells.filter(c=>!c.isNullObject);
    for(const cell of validCells)ranges.push({font:cell.font,cell,text:cell.text});
    ranges.tableCellCount=validCells.length;
    return ranges;
  }
  async function selection(context){const selected=context.presentation.getSelectedShapes();selected.load('items/type');await context.sync();return selected.items;}
  async function textSizeTargets(context,ranges){
    const targets=[];
    for(const range of ranges){
      if(range.cell)continue;
      const frame=range.getParentTextFrame(),shape=frame.getParentShape();
      shape.load('type,width,height,left,top');targets.push({frame,shape});
    }
    if(targets.length)await context.sync();
    return targets.filter(t=>['TextBox','GeometricShape'].includes(t.shape.type)).map(t=>({...t,isTextBox:t.shape.type==='TextBox',width:t.shape.width,height:t.shape.height,left:t.shape.left,top:t.shape.top}));
  }
  const supported=s=>['GeometricShape','TextBox','Line'].includes(s.type);
  const supportsFill=s=>supported(s)&&s.type!=='Line';
  const supportsLine=s=>supported(s)||s.type==='Image';
  async function capture(kind){
    const values=await PowerPoint.run(async context=>{
      const shapes=await selection(context),found=[];
      if(kind==='colors'){
        const objects=shapes.filter(supportsLine);
        objects.forEach(s=>{if(supportsFill(s))s.fill.load('type,foregroundColor');s.lineFormat.load('visible,color');});
        await context.sync();
        objects.forEach(s=>{if(supportsFill(s)&&s.fill.type==='Solid')found.push(s.fill.foregroundColor);if(s.lineFormat.visible)found.push(s.lineFormat.color);});
      }
      {
        const ranges=await textTargets(context,shapes);if(!ranges.length&&kind!=='colors')throw new Error('文字のあるオブジェクト、または文字範囲を選択してください。');
        const property=kind==='fonts'?'name':kind==='sizes'?'size':'color';ranges.forEach(r=>r.font.load(property));await context.sync();
        let chars=0;const fragments=[],cellFragments=[];
        for(const range of ranges){
          if(range.cell&&range.text.length===0)continue;
          if(range.font[property])found.push(range.font[property]);
          else if(range.cell){range.cell.load('textRuns');cellFragments.push(range.cell);}
          else {range.load('text');fragments.push(range);}
        }
        if(cellFragments.length){await context.sync();cellFragments.forEach(cell=>cell.textRuns.forEach(run=>{if(run.font?.[property])found.push(run.font[property]);}));}
        if(fragments.length){
          await context.sync();const parts=[];
          for(const range of fragments){chars+=range.text.length;if(chars>4000)throw new Error('書式が混在した文字が4,000文字を超えています。登録したい部分を範囲選択してください。');
            for(let i=0;i<range.text.length;i++){const p=range.getSubstring(i,1);p.font.load(property);parts.push(p);}}
          await context.sync();parts.forEach(p=>{if(p.font[property])found.push(p.font[property]);});
        }
      }
      return [...new Set(found.map(v=>kind==='colors'?C.color(v):v).filter(Boolean))];
    });
    if(!values.length)throw new Error('登録できる色・フォント・サイズがありません。対象を選ぶか、手入力してください。');
    await refresh();let next=palette;values.forEach(v=>{next=C.add(next,kind,v,kind==='colors'?$('color-name').value:kind==='fonts'?$('font-label').value:'');});await persist(next);
  }
  async function applyStyle(kind,value,target){
    const result=await PowerPoint.run(async context=>{
      const shapes=await selection(context);let count=0,skipped=0,tableCellCount=0,heightFitCount=0;
      if(kind==='fonts'||kind==='sizes'||target==='text'){
        const ranges=await textTargets(context,shapes);if(!ranges.length)throw new Error('文字のあるオブジェクト、または文字範囲を選択してください。');
        const sizeTargets=kind==='sizes'?await textSizeTargets(context,ranges):[];
        const fit=sizeTargets.filter(item=>item.isTextBox);
        // Disable any previous shape-to-text fit before changing text size.
        // Geometric shapes stay fixed; only actual text boxes get height fitting.
        for(const item of sizeTargets)item.frame.autoSizeSetting='AutoSizeNone';
        for(const range of ranges){if(kind==='fonts')range.font.name=value;else if(kind==='sizes')range.font.size=value;else range.font.color=value;count++;}
        for(const item of fit){item.frame.wordWrap=true;item.frame.autoSizeSetting='AutoSizeShapeToFitText';}
        if(sizeTargets.length){
          await context.sync();
          for(const item of sizeTargets){
            item.shape.width=item.width;item.shape.left=item.left;item.shape.top=item.top;
            if(!item.isTextBox)item.shape.height=item.height;
          }
          heightFitCount=fit.length;
        }
        skipped=Math.max(0,shapes.length-(ranges.coveredCount??ranges.length));
        tableCellCount=ranges.tableCellCount||0;
      }else{
        if(target==='line'){
          shapes.filter(supportsLine).forEach(s=>s.lineFormat.load('visible,weight,transparency'));
          await context.sync();
        }
        for(const shape of shapes){if(!(target==='fill'?supportsFill(shape):supportsLine(shape))){skipped++;continue;}
          if(target==='fill')shape.fill.setSolidColor(value);else {
            const line=shape.lineFormat;
            const newLine=line.visible!==true;
            if(newLine||!Number.isFinite(line.weight)||line.weight<=0)line.weight=1;
            if(newLine){line.dashStyle='Solid';line.style='Single';}
            if(newLine||!Number.isFinite(line.transparency)||line.transparency>=1)line.transparency=0;
            line.color=value;line.visible=true;
          }count++;
        }
        if(!count)throw new Error('適用できる図形が選択されていません。');
      }
      await context.sync();return {count,skipped,tableCellCount,heightFitCount};
    });
    status(`${result.count}件に適用。${result.tableCellCount?` 表：${result.tableCellCount}セル。`:''}${result.skipped?` 対象外：${result.skipped}件。`:''}`,'success');
  }
  document.addEventListener('DOMContentLoaded',()=>{
    $('picker').addEventListener('input',()=>{$('hex').value=$('picker').value.toUpperCase();});
    $('hex').addEventListener('input',()=>{const color=C.color($('hex').value);if(color)$('picker').value=color;});
    $('color-form').addEventListener('submit',e=>{e.preventDefault();action(()=>register('colors',$('hex').value,$('color-name').value));});
    $('font-form').addEventListener('submit',e=>{e.preventDefault();action(()=>register('fonts',$('font-name').value,$('font-label').value));});
    $('size-form').addEventListener('submit',e=>{e.preventDefault();action(()=>register('sizes',$('font-size').value,''));});
    $('capture-color').addEventListener('click',()=>action(()=>capture('colors')));
    $('capture-font').addEventListener('click',()=>action(()=>capture('fonts')));
    $('capture-size').addEventListener('click',()=>action(()=>capture('sizes')));
    $('reload').addEventListener('click',()=>action(async()=>{await refresh();status('このファイルの登録情報を読み込みました。','success');}));
    $('export-styles').addEventListener('click',()=>action(exportStyles));
    $('import-styles').addEventListener('click',()=>{if(ready&&!busy)$('import-file').click();});
    $('import-file').addEventListener('change',()=>{
      const file=$('import-file').files[0];$('import-file').value='';
      if(!file)return;
      action(async()=>{if(file.size>C.transferLimit)throw new Error('設定ファイルは256 KB以内で読み込んでください。');await importText(await file.text());});
    });
    $('import-text').addEventListener('click',()=>action(()=>importText($('import-json').value)));
    window.addEventListener('unload',()=>{if(exportUrl)URL.revokeObjectURL(exportUrl);});
    Office.onReady(async info=>{
      try{
        if(info.host!==Office.HostType.PowerPoint)throw new Error('PowerPointのアドインとして開いてください。');
        if(!Office.context.requirements.isSetSupported('PowerPointApi','1.8'))throw new Error('PowerPointApi 1.8以上が必要です。');
        palette=C.clean(Office.context.document.settings.get(C.key));ready=true;render();
        status('準備完了');
        Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged,()=>{if(!busy)action(async()=>{await refresh();status('登録情報を更新しました。');});},result=>{if(result.status!==Office.AsyncResultStatus.Succeeded)console.info('SettingsChanged unavailable; manual refresh supported.');});
      }catch(e){status(e.message,'error');}
    }).catch(e=>status(e.message,'error'));
  });
})();
