(function(root){
  'use strict';
  const descendants=(node,name)=>node?Array.from(node.getElementsByTagName('*')).filter(n=>n.localName===name):[];
  const child=(node,name)=>Array.from(node?.childNodes||[]).find(n=>n.nodeType===1&&n.localName===name);
  function parse(text){const doc=new DOMParser().parseFromString(text,'application/xml');if(descendants(doc,'parsererror').length)throw new Error('スライドXMLが不正です。');return doc;}
  function records(doc){
    const tree=descendants(doc,'spTree')[0];if(!tree)throw new Error('図形情報がありません。');
    return Array.from(tree.childNodes).filter(n=>n.nodeType===1&&['sp','pic','graphicFrame','grpSp','cxnSp','contentPart'].includes(n.localName)).map((node,order)=>{
      const props=descendants(node,'cNvPr')[0],x=descendants(node,'xfrm')[0],off=x&&child(x,'off'),ext=x&&child(x,'ext');
      return {node,order,id:props?.getAttribute('id'),name:props?.getAttribute('name'),box:off&&ext?{left:+off.getAttribute('x')/12700,top:+off.getAttribute('y')/12700,width:+ext.getAttribute('cx')/12700,height:+ext.getAttribute('cy')/12700}:null};
    });
  }
  function match(list,target){
    const same=r=>r.box&&['left','top','width','height'].every(k=>Math.abs(r.box[k]-target[k])<.15);
    let found=list.filter(r=>r.id===String(target.id)&&same(r)&&(r.name===target.name||r.order===target.zOrderPosition));
    if(found.length!==1)found=list.filter(r=>r.name===target.name&&same(r));
    if(found.length!==1)found=list.filter(same);
    if(found.length>1)found=found.filter(r=>r.order===target.zOrderPosition);
    if(found.length!==1)throw new Error('図形を照合できません。グループを解除して再試行してください。');
    return found[0];
  }
  async function transform(base64,targets,operation){
    if(operation.kind!=='round'||!Number.isFinite(operation.value)||operation.value<0)throw new Error('指定値が不正です。');
    const zip=await JSZip.loadAsync(base64,{base64:true}),paths=Object.keys(zip.files).filter(p=>/^ppt\/slides\/slide\d+\.xml$/.test(p));
    if(paths.length!==1)throw new Error('1スライドずつ処理してください。');
    const doc=parse(await zip.file(paths[0]).async('string')),list=records(doc),checks=[];let skipped=0;
    const seen=new Set();
    for(const target of targets){
      const r=match(list,target);if(seen.has(r.order))continue;seen.add(r.order);
      const node=r.node,body=child(node,'txBody'),props=child(node,'spPr'),geom=props&&child(props,'prstGeom');
      {
        if(!geom||!['rect','roundRect'].includes(geom.getAttribute('prst'))){skipped++;continue;}
        const short=Math.min(r.box.width,r.box.height);
        if(operation.value>short/2+1e-7)throw new Error('登録半径が収まらない図形があります。小さい半径を登録してください。');
        if(!Number.isFinite(short)||short<=0)throw new Error('図形サイズが不正です。');
        const ns=geom.namespaceURI,prefix=geom.prefix||'a';geom.setAttribute('prst','roundRect');
        let av=child(geom,'avLst');if(!av){av=doc.createElementNS(ns,`${prefix}:avLst`);geom.insertBefore(av,geom.firstChild);}
        while(av.firstChild)av.removeChild(av.firstChild);
        const gd=doc.createElementNS(ns,`${prefix}:gd`),adjustment=Math.round(operation.value/short*100000);gd.setAttribute('name','adj');gd.setAttribute('fmla',`val ${adjustment}`);av.appendChild(gd);
        checks.push({order:r.order,kind:'round',value:adjustment,text:body?descendants(body,'t').map(n=>n.textContent).join(''):''});
      }
    }
    if(!checks.length)throw new Error('四角形または角丸四角形を選択してください。');
    zip.file(paths[0],new XMLSerializer().serializeToString(doc));
    // Keep unrelated ZIP members untouched. Hyperlink validation/remapping is out of scope.
    return {base64:await zip.generateAsync({type:'base64',compression:'DEFLATE'}),checks,count:checks.length,skipped};
  }
  async function verify(base64,checks){
    const zip=await JSZip.loadAsync(base64,{base64:true}),path=Object.keys(zip.files).find(p=>/^ppt\/slides\/slide\d+\.xml$/.test(p));
    if(!path)throw new Error('挿入結果を確認できません。');
    const doc=parse(await zip.file(path).async('string')),list=records(doc);
    for(const check of checks){
      const node=list.find(r=>r.order===check.order)?.node;if(!node)throw new Error('挿入後の図形がありません。');
      const text=descendants(child(node,'txBody')||node,'t').map(n=>n.textContent).join('');
      if(text!==check.text)throw new Error('挿入後の文字が一致しません。');
      {
        const geom=descendants(child(node,'spPr'),'prstGeom')[0],gd=geom&&descendants(geom,'gd').find(n=>n.getAttribute('name')==='adj');
        if(geom?.getAttribute('prst')!=='roundRect'||gd?.getAttribute('fmla')!==`val ${check.value}`)throw new Error('角丸の反映を確認できません。');
      }
    }
  }
  async function apply(operation,{replace=false,status=()=>{}}={}){
    if(operation.kind!=='round')throw new Error('この操作は対応していません。');
    if(!Office.context.requirements.isSetSupported('PowerPointApi','1.8'))throw new Error('PowerPointApi 1.8以上が必要です。');
    return PowerPoint.run(async context=>{
      const selected=context.presentation.getSelectedShapes();selected.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/zOrderPosition');await context.sync();
      const targets=selected.items.filter(s=>['GeometricShape'].includes(s.type));
      if(!targets.length)throw new Error('文字のある図形または四角形を選択してください。');
      const owners=targets.map(shape=>({shape,slide:shape.getParentSlide()}));owners.forEach(o=>o.slide.load('id'));await context.sync();
      const groups=new Map();for(const o of owners){if(!groups.has(o.slide.id))groups.set(o.slide.id,{slide:o.slide,targets:[]});
        const meta={};for(const key of ['id','name','type','left','top','width','height','zOrderPosition'])meta[key]=o.shape[key];groups.get(o.slide.id).targets.push(meta);
      }
      const prepared=[];let completed=0;
      // Preflight all target slides before any insertion or deletion.
      for(const group of groups.values()){
        status('変更済みスライドを準備中…');const result=group.slide.exportAsBase64();await context.sync();
        prepared.push({...group,...await transform(result.value,group.targets,operation)});
      }
      let count=0;
      try{for(const group of prepared){
        const slides=context.presentation.slides;slides.load('items/id');await context.sync();const before=new Set(slides.items.map(s=>s.id));
        context.presentation.insertSlidesFromBase64(group.base64,{targetSlideId:group.slide.id,formatting:'KeepSourceFormatting'});await context.sync();
        slides.load('items/id');await context.sync();const added=slides.items.filter(s=>!before.has(s.id));
        if(added.length!==1)throw new Error('挿入結果を特定できません。元スライドは削除していません。追加されたスライドを確認してください。');
        try{const exported=added[0].exportAsBase64();await context.sync();await verify(exported.value,group.checks);}catch(e){throw new Error(`${e.message} 元スライドは残しています。追加コピーを確認してください。`);}
        if(replace){group.slide.delete();await context.sync();}
        context.presentation.setSelectedSlides([added[0].id]);await context.sync();completed++;
        count+=group.count;
      }}catch(error){if(completed)throw new Error(`${completed}スライドは処理済みです。${error.message}`);throw error;}
      return {count,slides:completed,replace};
    });
  }
  root.FileStyleXml={transform,verify,apply};
})(window);
