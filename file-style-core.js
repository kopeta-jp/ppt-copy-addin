(function(root){
  'use strict';
  const key='fileStyle.palette.v1', limit=65536, warning=32768;
  function color(value){
    if(typeof value!=='string') return null;
    let s=value.trim(); if(!s.startsWith('#')) s='#'+s;
    if(/^#[0-9a-f]{3}$/i.test(s)) s='#'+s.slice(1).split('').map(c=>c+c).join('');
    return /^#[0-9a-f]{6}$/i.test(s)?s.toUpperCase():null;
  }
  function clean(raw){
    if(raw==null) return {version:1,colors:[],fonts:[],sizes:[],dimensions:[],roundings:[]};
    if(raw.version!==1||!Array.isArray(raw.colors)||!Array.isArray(raw.fonts)) throw new Error('保存データの形式が異なります。登録内容を保護するため変更を停止しました。');
    if(raw.sizes!==undefined&&!Array.isArray(raw.sizes))throw new Error('保存したサイズ情報の形式が異なります。');
    if(raw.dimensions!==undefined&&!Array.isArray(raw.dimensions))throw new Error("保存した寸法情報の形式が異なります。");
    if(raw.roundings!==undefined&&!Array.isArray(raw.roundings))throw new Error("保存した角丸情報の形式が異なります。");
    const result={version:1,colors:[],fonts:[],sizes:[],dimensions:[],roundings:[]};
    for(const item of raw.colors){const value=color(item.value);if(value&&!result.colors.some(x=>x.value===value)) result.colors.push({value,label:String(item.label||value).slice(0,60)});}
    for(const item of raw.fonts){const value=String(item.value||'').trim().slice(0,120);if(value&&!result.fonts.some(x=>x.value.toLowerCase()===value.toLowerCase())) result.fonts.push({value,label:String(item.label||value).slice(0,60)});}
    for(const item of raw.sizes||[]){const value=size(item.value);if(value!==null&&!result.sizes.some(x=>x.value===value))result.sizes.push({value,label:String(item.label||`${value} pt`).slice(0,60)});}
    for(const item of raw.dimensions||[]){const value=dimension(item.value);if(value!==null&&!result.dimensions.some(x=>x.value===value))result.dimensions.push({value,label:String(item.label||`${value} cm`).slice(0,60)});}
    for(const item of raw.roundings||[]){
      const value=rounding(item.value),mode=item.mode||'ratio';
      if(!['ratio','radius'].includes(mode))throw new Error('保存した角丸情報の形式が異なります。');
      const scale=item.scale===undefined?1:item.scale;
      if(mode==='radius'&&(!Number.isFinite(scale)||scale<=0))throw new Error('保存した角丸情報の形式が異なります。');
      if(value!==null&&!result.roundings.some(x=>x.value===value&&x.mode===mode))result.roundings.push({value,label:String(item.label||`丸み ${result.roundings.length+1}`).slice(0,60),mode,...(mode==='radius'?{scale}:{})});
    }
    return result;
  }
  function rounding(value){return typeof value==="number"&&Number.isFinite(value)&&value>=0?value:null;}
  function dimension(value){
    if(typeof value!=="number"&&typeof value!=="string")return null;
    if(typeof value==="string"&&!/^\d+(\.\d+)?$/.test(value.trim()))return null;
    const n=Number(value);return Number.isFinite(n)&&n>=0.01&&n<=100?n:null;
  }
  function size(value){
    if(typeof value!=='number'&&typeof value!=='string')return null;
    if(typeof value==='string'&&!/^\d+(\.\d+)?$/.test(value.trim()))return null;
    const n=Number(value);return Number.isFinite(n)&&n>=1&&n<=400?n:null;
  }
  function bytes(data){return new TextEncoder().encode(JSON.stringify(data)).length;}
  function add(data,kind,value,label,roundingOptions){
    const next=clean(data);
    if(kind==='colors'){value=color(value);if(!value) throw new Error('カラーコードは #2155CD のような6桁で入力してください。');}
    else if(kind==='fonts'){value=String(value||'').trim();if(!value||value.length>120) throw new Error('フォント名を120文字以内で入力してください。');}
    else if(kind==='sizes'){value=size(value);if(value===null)throw new Error('フォントサイズを1〜400 ptの数値で入力してください。');}
    else if(kind==='dimensions'){value=dimension(value);if(value===null)throw new Error('寸法を0.01〜100 cmの数値で入力してください。');}
    else if(kind==='roundings'){value=rounding(value);if(value===null)throw new Error('角丸の調整値が不正です。');}
    else throw new Error('登録種別が不正です。');
    const mode=roundingOptions?.mode||'ratio',scale=roundingOptions?.scale===undefined?1:roundingOptions.scale;
    if(kind==='roundings'&&(!['ratio','radius'].includes(mode)||!Number.isFinite(scale)||scale<=0))throw new Error('角丸の登録情報が不正です。');
    if(next[kind].some(x=>String(x.value).toLowerCase()===String(value).toLowerCase()&&(kind!=='roundings'||x.mode===mode))) return next;
    next[kind].push({value,label:String(label||(kind==='sizes'?`${value} pt`:kind==='dimensions'?`${value} cm`:kind==='roundings'?`丸み ${next.roundings.length+1}`:value)).trim().slice(0,60)||String(value),...(kind==='roundings'?{mode,...(mode==='radius'?{scale}:{})}:{})});
    if(bytes(next)>limit) throw new Error('登録データが64 KBを超えるため保存を停止しました。不要な登録を削除してください。');
    return next;
  }
  const transferLimit=262144;
  function exportStyles(data){
    const palette=clean(data);
    return JSON.stringify({format:'FILE-STYLE',formatVersion:palette.roundings.some(x=>x.mode==='radius')?2:1,palette},null,2);
  }
  function parseStyles(text){
    if(typeof text!=='string'||new TextEncoder().encode(text).length>transferLimit)throw new Error('設定ファイルは256 KB以内で読み込んでください。');
    let raw;
    try{raw=JSON.parse(text.replace(/^\uFEFF/,''));}catch(_){throw new Error('設定ファイルを読み込めません。FILE STYLEから書き出したJSONファイルを選んでください。');}
    if(!raw||raw.format!=='FILE-STYLE'||![1,2].includes(raw.formatVersion))throw new Error('FILE STYLEの設定ファイルではないか、未対応のバージョンです。');
    const data=raw.palette;
    if(!data||data.version!==1||!Array.isArray(data.colors)||!Array.isArray(data.fonts)||!Array.isArray(data.sizes))throw new Error('設定ファイルの登録情報が不正です。');
    if(data.dimensions!==undefined&&!Array.isArray(data.dimensions))throw new Error('設定ファイルの寸法情報が不正です。');
    if(data.roundings!==undefined&&!Array.isArray(data.roundings))throw new Error('設定ファイルの角丸情報が不正です。');
    for(const kind of ['colors','fonts','sizes','dimensions','roundings'])for(const item of data[kind]||[]){
      if(!item||typeof item!=='object'||Array.isArray(item)||typeof item.label!=='string'||item.label.length>60)throw new Error('設定ファイルに不正な登録項目があります。');
      if(kind==='colors'&&!color(item.value))throw new Error('設定ファイルに不正なカラーコードがあります。');
      if(kind==='fonts'&&(typeof item.value!=='string'||!item.value.trim()||item.value.length>120))throw new Error('設定ファイルに不正なフォント名があります。');
      if(kind==='roundings'&&rounding(item.value)===null)throw new Error('設定ファイルに不正な角丸情報があります。');
      if(kind==='roundings'&&item.mode!==undefined&&!['ratio','radius'].includes(item.mode))throw new Error('設定ファイルに不正な角丸方式があります。');
      if(kind==='roundings'&&item.scale!==undefined&&(!Number.isFinite(item.scale)||item.scale<=0))throw new Error('設定ファイルに不正な角丸倍率があります。');
      if(kind==='dimensions'&&dimension(item.value)===null)throw new Error('設定ファイルに不正な寸法があります。');
      if(kind==='sizes'&&size(item.value)===null)throw new Error('設定ファイルに不正なサイズがあります。');
    }
    const result=clean(data);
    if(bytes(result)>limit)throw new Error('登録データが64 KBを超えているため読み込めません。');
    return result;
  }
  function mergeStyles(existing,incoming){
    let merged=clean(existing);const before={colors:merged.colors.length,fonts:merged.fonts.length,sizes:merged.sizes.length,dimensions:merged.dimensions.length,roundings:merged.roundings.length};
    for(const kind of ['colors','fonts','sizes','dimensions','roundings'])for(const item of incoming[kind]||[])merged=add(merged,kind,item.value,item.label,kind==='roundings'?{mode:item.mode||'ratio',scale:item.scale}:undefined);
    const added={};for(const kind of ['colors','fonts','sizes','dimensions','roundings'])added[kind]=merged[kind].length-before[kind];
    return {palette:merged,added};
  }
  const api={key,limit,warning,color,size,dimension,rounding,clean,bytes,add,transferLimit,exportStyles,parseStyles,mergeStyles};
  if(typeof module!=='undefined'&&module.exports) module.exports=api;
  root.FileStyleCore=api;
})(typeof window!=='undefined'?window:globalThis);
