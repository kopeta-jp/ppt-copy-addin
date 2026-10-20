(function(){
  'use strict';
  const C=window.FileStyleCore;
  let palette=C.clean(null), ready=false, busy=false, exportUrl=null, slideBlob=null, slideUrl=null, audit=null;
  const $=id=>document.getElementById(id);
  function status(message,type=''){ $('status').textContent=message;$('status').className=type; }
  function controls(){document.querySelectorAll('button').forEach(b=>b.disabled=!ready||busy);$('round-replace').disabled=!ready||busy;$('slide-retry').disabled=!ready||busy||!slideBlob;if(audit)audit.controls(ready&&!busy);}
  async function action(fn){
    if(!ready||busy)return;
    busy=true;controls();
    try{await fn();}catch(error){console.error(error);status(error.message||'処理に失敗しました。','error');}
    finally{busy=false;controls();}
  }
  function render(){
    for(const kind of ['colors','fonts','sizes','dimensions','roundings']){
      const list=$(kind);list.replaceChildren();
      if(!palette[kind].length){const p=document.createElement('p');p.className='hint';p.textContent='未登録';list.append(p);}
      palette[kind].forEach((item,index)=>{
        const entry=document.createElement('div');entry.className='entry';
        const apply=document.createElement(kind==='colors'||kind==='dimensions'?'div':'button');apply.className=kind==='colors'?'color-card':kind==='dimensions'?'dimension-card':'apply';
        if(kind!=='colors')apply.title='選択したオブジェクトに適用';
        if(kind==='colors'){const swatch=document.createElement('span');swatch.className='swatch';swatch.style.backgroundColor=item.value;apply.append(swatch);}
        const text=document.createElement('span');text.className='label';text.textContent=item.label;
        if(kind==='roundings')text.textContent+=item.mode==='radius'?` · 半径 ${(item.value*25.4/72).toFixed(2)} mm`:'（旧・再登録）';
        const value=document.createElement('small');value.textContent=kind==='sizes'?`${item.value} pt`:kind==='dimensions'?`${item.value} cm`:item.value;
        if(kind!=='roundings'&&(!['sizes','dimensions'].includes(kind)||item.label!==value.textContent))text.append(value);apply.append(text);
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
        }else if(kind==='dimensions'){
          const actions=document.createElement('div');actions.className='dimension-actions';
          for(const [axis,label] of [['width','横幅'],['height','縦幅']]){
            const button=document.createElement('button');button.textContent=label;
            button.setAttribute('aria-label',`${label}を${item.value} cmに合わせる`);
            button.addEventListener('click',()=>action(()=>applyDimension(item.value,axis)));actions.append(button);
          }
          apply.append(actions);
        }else apply.addEventListener('click',()=>action(()=>kind==='roundings'?applyRounding(item):applyStyle(kind,item.value)));
        const remove=document.createElement('button');remove.className='delete';remove.textContent='削除';remove.setAttribute('aria-label',item.label+'を削除');
        remove.addEventListener('click',()=>action(async()=>{await refresh();const next=C.clean(palette);const at=next[kind].findIndex(x=>x.value===item.value&&(kind!=='roundings'||x.mode===item.mode));if(at>=0)next[kind].splice(at,1);await persist(next);}));
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
    if(a.colors+a.fonts+a.sizes+a.dimensions+a.roundings===0){status(a.lineSpacings?'行間機能は廃止されています。登録情報は変更していません。':'すべて登録済みでした。登録情報は変更していません。','success');return;}
    await persist(merged.palette);
    const large=C.bytes(merged.palette)>=C.warning;
    status(`追加：色 ${a.colors}・フォント ${a.fonts}・サイズ ${a.sizes}・寸法 ${a.dimensions}・角丸 ${a.roundings}。${large?'登録データが32 KBを超えています。':''} ⌘Sで保存。`,large?'warning':'success');
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
        // Geometric shapes stay fixed; text boxes fit their unwrapped text in both dimensions.
        for(const item of sizeTargets)item.frame.autoSizeSetting='AutoSizeNone';
        for(const range of ranges){if(kind==='fonts')range.font.name=value;else if(kind==='sizes')range.font.size=value;else range.font.color=value;count++;}
        for(const item of fit){item.frame.wordWrap=false;item.frame.autoSizeSetting='AutoSizeShapeToFitText';}
        if(sizeTargets.length){
          await context.sync();
          for(const item of sizeTargets){
            item.shape.left=item.left;item.shape.top=item.top;
            if(!item.isTextBox){item.shape.width=item.width;item.shape.height=item.height;}
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
  async function applyDimension(value,axis){
    const cm=C.dimension(value);
    if(cm===null||!['width','height'].includes(axis))throw new Error('寸法が不正です。');
    const result=await PowerPoint.run(async context=>{
      const shapes=await selection(context);
      const targets=shapes.filter(s=>['Image','GeometricShape'].includes(s.type));
      if(!targets.length)throw new Error('画像または図形を選択してください。');
      targets.forEach(s=>s.load('width,height,left,top'));await context.sync();
      const plans=targets.map(shape=>{
        const {width,height,left,top}=shape;
        if(!Number.isFinite(width)||!Number.isFinite(height)||width<=0||height<=0)throw new Error('寸法を取得できないオブジェクトがあります。');
        const points=cm*72/2.54,factor=points/(axis==='width'?width:height);
        const w=axis==='width'?points:width*factor,h=axis==='height'?points:height*factor;
        if(!Number.isFinite(w)||!Number.isFinite(h)||w>4032||h>4032)throw new Error('変換後のサイズが大きすぎます。小さい数値を指定してください。');
        return {shape,width:w,height:h,left,top};
      });
      const frames=plans.filter(p=>p.shape.type==='GeometricShape').map(p=>p.shape.getTextFrameOrNullObject());
      frames.forEach(f=>f.load('hasText'));if(frames.length)await context.sync();
      // Prevent text autofit from overriding the explicitly requested shape dimensions.
      frames.filter(f=>!f.isNullObject&&f.hasText).forEach(f=>f.autoSizeSetting='AutoSizeNone');
      for(const p of plans){p.shape.width=p.width;p.shape.height=p.height;p.shape.left=p.left;p.shape.top=p.top;}
      await context.sync();return {count:plans.length,skipped:shapes.length-plans.length};
    });
    status(`${result.count}件の${axis==='width'?'横幅':'縦幅'}を${cm} cmに変更。${result.skipped?` 対象外：${result.skipped}件。`:''}`,'success');
  }
  async function roundedSelection(context){
    if(!Office.context.requirements.isSetSupported('PowerPointApi','1.10'))throw new Error('角丸の操作にはPowerPointApi 1.10以上が必要です。');
    const shapes=await selection(context);
    const geometric=shapes.filter(s=>s.type==='GeometricShape');
    geometric.forEach(s=>s.load('id,name,width,height,left,top,rotation,zOrderPosition'));await context.sync();
    const rounded=[],rectangles=[],geometry=new Map();
    if(geometric.length){
      if(typeof JSZip==='undefined')throw new Error('角丸判定用のファイルがありません。file-style-jszip.min.jsをアップしてください。');
      // Office.js does not expose the preset geometry of an existing shape.
      // Read the exported slide XML only; never import or replace the source shapes.
      const owners=geometric.map(shape=>({shape,slide:shape.getParentSlide()}));
      owners.forEach(o=>o.slide.load('id'));await context.sync();
      const groups=new Map();for(const o of owners){if(!groups.has(o.slide.id))groups.set(o.slide.id,{slide:o.slide,shapes:[]});groups.get(o.slide.id).shapes.push(o.shape);}
      const element=(node,name)=>Array.from(node.getElementsByTagName('*')).find(n=>n.localName===name);
      const sameBox=(record,shape)=>record.box&&['left','top','width','height'].every(k=>Number.isFinite(shape[k])&&Math.abs(record.box[k]-shape[k])<=0.15);
      for(const group of groups.values()){
        const exported=group.slide.exportAsBase64();await context.sync();
        const zip=await JSZip.loadAsync(exported.value,{base64:true});
        const files=Object.keys(zip.files).filter(p=>/^ppt\/slides\/slide\d+\.xml$/.test(p));
        if(files.length!==1)throw new Error('角丸の種類を確認できませんでした。');
        const xml=new DOMParser().parseFromString(await zip.file(files[0]).async('string'),'application/xml');
        if(xml.getElementsByTagName('parsererror').length)throw new Error('図形情報を読み込めませんでした。');
        const tree=element(xml,'spTree');
        // Strict OOXML and exported decks can use different namespaces/names/IDs.
        // Use local names, then cross-check identity with unchanged geometry and stacking order.
        const nodes=tree?Array.from(tree.childNodes).filter(n=>n.nodeType===1&&['sp','pic','graphicFrame','grpSp','cxnSp','contentPart'].includes(n.localName)):[];
        const records=nodes.map((node,zOrder)=>{
          const props=element(node,'cNvPr'),geom=element(node,'prstGeom'),transform=element(node,'xfrm');
          const off=transform&&element(transform,'off'),ext=transform&&element(transform,'ext');
          const box=off&&ext?{left:Number(off.getAttribute('x'))/12700,top:Number(off.getAttribute('y'))/12700,width:Number(ext.getAttribute('cx'))/12700,height:Number(ext.getAttribute('cy'))/12700}:null;
          const guide=geom&&Array.from(geom.getElementsByTagName('*')).find(n=>n.localName==='gd'&&n.getAttribute('name')==='adj');
          const formula=guide?.getAttribute('fmla'),adjustment=guide&&/^val\s+-?\d+(\.\d+)?$/.test(formula||'')?Number(formula.trim().split(/\s+/)[1])/100000:guide?NaN:16667/100000;
          return {id:props?.getAttribute('id'),name:props?.getAttribute('name'),preset:node.localName==='sp'?geom?.getAttribute('prst'):null,box,zOrder,rawRatio:adjustment,ratio:Math.max(0,Math.min(.5,adjustment))};
        });
        for(const shape of group.shapes){
          let candidates=records.filter(r=>r.id===String(shape.id)&&(sameBox(r,shape)?r.name===shape.name||r.zOrder===shape.zOrderPosition:!r.box&&r.name===shape.name));
          if(candidates.length!==1)candidates=records.filter(r=>r.name===shape.name&&sameBox(r,shape));
          if(candidates.length!==1)candidates=records.filter(r=>sameBox(r,shape));
          if(candidates.length>1)candidates=candidates.filter(r=>r.zOrder===shape.zOrderPosition);
          if(candidates.length!==1)throw new Error('図形情報を照合できません。グループ内の図形は解除して再試行してください。');
          if(candidates[0].preset==='rect')rectangles.push(shape);
          if(candidates[0].preset==='roundRect'){
            if(!Number.isFinite(candidates[0].ratio))throw new Error('角の半径を読み取れませんでした。');
            rounded.push(shape);geometry.set(shape,candidates[0]);
          }
        }
      }
    }
    rounded.forEach(s=>s.adjustments.load('count'));if(rounded.length)await context.sync();
    if(rounded.some(s=>s.adjustments.count!==1))throw new Error('対応していない角丸設定が含まれています。');
    return {shapes,rounded,rectangles,geometry};
  }
  async function captureRounding(){
    const captured=await PowerPoint.run(async context=>{
      const {shapes,rounded,geometry}=await roundedSelection(context);
      if(shapes.length!==1||rounded.length!==1)throw new Error('基準の角丸四角形を1つだけ選択してください。');
      const result=rounded[0].adjustments.get(0);await context.sync();
      if(C.rounding(result.value)===null)throw new Error('角丸の設定を取得できません。');
      const shape=rounded[0],short=Math.min(shape.width,shape.height),ratio=geometry.get(shape).ratio;
      if(!Number.isFinite(short)||short<=0)throw new Error('図形の寸法を取得できません。');
      const rawRatio=geometry.get(shape).rawRatio,scale=rawRatio>0?result.value/rawRatio:1;
      if(!Number.isFinite(scale)||scale<=0)throw new Error('丸みの調整値と半径を対応づけられません。');
      return {radius:ratio*short,scale};
    });
    await refresh();await persist(C.add(palette,'roundings',captured.radius,$('rounding-name').value,{mode:'radius',scale:captured.scale}));
  }
  async function applyRounding(item){
    if(item.mode!=='radius')throw new Error('旧設定には半径情報がありません。基準の角丸を再登録してください。');
    const radius=item.value;
    if(C.rounding(radius)===null)throw new Error('角丸の設定が不正です。');
    const result=await PowerPoint.run(async context=>{
      const {shapes,rounded,rectangles,geometry}=await roundedSelection(context);
      if(rectangles.length)return {convert:true};
      if(!rounded.length)throw new Error('角丸四角形を選択してください。');
      const values=rounded.map(shape=>{
        const short=Math.min(shape.width,shape.height);
        if(!Number.isFinite(short)||short<=0)throw new Error('図形の寸法を取得できません。');
        if(radius>short/2+1e-7)throw new Error('登録した半径が収まらない図形があります。小さい丸みを登録してください。変更は行っていません。');
        return {shape,short,current:shape.adjustments.get(0),ratio:geometry.get(shape).rawRatio};
      });await context.sync();
      const plans=values.map(o=>{
        const scale=o.ratio>0?o.current.value/o.ratio:item.scale;
        if(!Number.isFinite(scale)||scale<=0)throw new Error('丸みの調整値を計算できません。');
        return {shape:o.shape,value:Math.min(.5,radius/o.short)*scale};
      });
      plans.forEach(p=>p.shape.adjustments.set(0,p.value));await context.sync();
      return {count:rounded.length,skipped:shapes.length-rounded.length};
    });
    if(result.convert){await applyXml('round',radius);return;}
    status(`${result.count}件の角半径を統一。${result.skipped?` 対象外：${result.skipped}件。`:''}`,'success');
  }
  async function applyXml(kind,value){
    const replace=$('round-replace').checked;
    const result=await window.FileStyleXml.apply({kind,value},{replace,status});
    status(`${result.count}件を${replace?'反映':'変更済みコピーに反映'}。${result.slides}スライド。リンクは引き継ぎ対象外です。`,'warning');
  }
  async function fitToSlide(){
    if(!Office.context.requirements.isSetSupported('PowerPointApi','1.10'))throw new Error('スライドサイズの取得にはPowerPointApi 1.10以上が必要です。');
    const result=await PowerPoint.run(async context=>{
      const page=context.presentation.pageSetup;page.load('slideWidth,slideHeight');
      const shapes=await selection(context);
      if(!Number.isFinite(page.slideWidth)||!Number.isFinite(page.slideHeight)||page.slideWidth<=0||page.slideHeight<=0)throw new Error('スライドサイズを取得できません。');
      const targets=shapes.filter(s=>['Image','GeometricShape','TextBox','Group'].includes(s.type));
      if(!targets.length)throw new Error('画像または図形を選択してください。');
      const frames=targets.filter(s=>['GeometricShape','TextBox'].includes(s.type)).map(s=>s.getTextFrameOrNullObject());
      frames.forEach(f=>f.load('hasText'));if(frames.length)await context.sync();
      frames.filter(f=>!f.isNullObject).forEach(f=>f.autoSizeSetting='AutoSizeNone');
      for(const shape of targets){shape.width=page.slideWidth;shape.height=page.slideHeight;shape.left=0;shape.top=0;}
      await context.sync();return {count:targets.length,skipped:shapes.length-targets.length};
    });
    status(`${result.count}件をスライド全面に合わせました。${result.skipped?` 対象外：${result.skipped}件。`:''}`,'success');
  }
  function writeSlide(blobPromise){
    if(typeof navigator==='undefined'||!navigator.clipboard||typeof navigator.clipboard.write!=='function'||typeof ClipboardItem==='undefined')throw new Error('画像コピー非対応');
    return navigator.clipboard.write([new ClipboardItem({'image/png':blobPromise})]);
  }
  function clearSlide(){
    if(slideUrl)URL.revokeObjectURL(slideUrl);
    slideUrl=null;slideBlob=null;$('slide-fallback').hidden=true;
    $('slide-preview').removeAttribute('src');$('slide-download').removeAttribute('href');
  }
  async function copySlide(){
    clearSlide();status('画像を生成中…');
    const rendered=PowerPoint.run(async context=>{
      const slides=context.presentation.getSelectedSlides();slides.load('items');await context.sync();
      if(slides.items.length!==1)throw new Error('スライドを1枚だけ選択してください。');
      const result=slides.items[0].getImageAsBase64({width:2560});await context.sync();
      if(!result.value)throw new Error('画像を生成できませんでした。');
      const base64=result.value.replace(/^data:image\/png;base64,/, '').replace(/\s/g,'');
      const binary=atob(base64),bytes=Uint8Array.from(binary,c=>c.charCodeAt(0));
      if([137,80,78,71,13,10,26,10].some((v,i)=>bytes[i]!==v))throw new Error('PNG画像を生成できませんでした。');
      const blob=new Blob([bytes],{type:'image/png'});
      slideBlob=blob;slideUrl=URL.createObjectURL(blob);
      $('slide-preview').src=slideUrl;$('slide-download').href=slideUrl;
      $('slide-download').download=`slide-${Date.now()}.png`;
      return blob;
    });
    // Invoke clipboard.write in the click turn; WebKit can accept a promised PNG.
    let copied;
    try{copied=Promise.resolve(writeSlide(rendered)).then(()=>true,()=>false);}catch(_){copied=Promise.resolve(false);}
    await rendered;
    if(await copied){status('画像をコピーしました。⌘Vで貼り付け。','success');}
    else{$('slide-fallback').hidden=false;status('直接コピー不可。再コピー・右クリック・PNG保存を利用してください。','warning');}
  }
  async function retrySlide(){
    if(!slideBlob)return;
    try{await writeSlide(Promise.resolve(slideBlob));status('画像をコピーしました。⌘Vで貼り付け。','success');}
    catch(_){$('slide-fallback').hidden=false;status('画像の右クリックまたはPNG保存を利用してください。','warning');}
  }
  document.addEventListener('DOMContentLoaded',()=>{
    audit=window.FileStyleAudit.create({status,action});
    $('fit-to-slide').addEventListener('click',()=>action(fitToSlide));
    $('capture-rounding').addEventListener('click',()=>action(captureRounding));
    $('slide-copy').addEventListener('click',()=>action(copySlide));
    $('slide-retry').addEventListener('click',()=>action(retrySlide));
    $('picker').addEventListener('input',()=>{$('hex').value=$('picker').value.toUpperCase();});
    $('hex').addEventListener('input',()=>{const color=C.color($('hex').value);if(color)$('picker').value=color;});
    $('color-form').addEventListener('submit',e=>{e.preventDefault();action(()=>register('colors',$('hex').value,$('color-name').value));});
    $('font-form').addEventListener('submit',e=>{e.preventDefault();action(()=>register('fonts',$('font-name').value,$('font-label').value));});
    $('size-form').addEventListener('submit',e=>{e.preventDefault();action(()=>register('sizes',$('font-size').value,''));});
    $('dimension-form').addEventListener('submit',e=>{e.preventDefault();action(()=>register('dimensions',$('dimension-value').value,''));});
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
    window.addEventListener('unload',()=>{if(exportUrl)URL.revokeObjectURL(exportUrl);if(slideUrl)URL.revokeObjectURL(slideUrl);});
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
