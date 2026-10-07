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
    if(raw==null) return {version:1,colors:[],fonts:[]};
    if(raw.version!==1||!Array.isArray(raw.colors)||!Array.isArray(raw.fonts)) throw new Error('保存データの形式が異なります。登録内容を保護するため変更を停止しました。');
    const result={version:1,colors:[],fonts:[]};
    for(const item of raw.colors){const value=color(item.value);if(value&&!result.colors.some(x=>x.value===value)) result.colors.push({value,label:String(item.label||value).slice(0,60)});}
    for(const item of raw.fonts){const value=String(item.value||'').trim().slice(0,120);if(value&&!result.fonts.some(x=>x.value.toLowerCase()===value.toLowerCase())) result.fonts.push({value,label:String(item.label||value).slice(0,60)});}
    return result;
  }
  function bytes(data){return new TextEncoder().encode(JSON.stringify(data)).length;}
  function add(data,kind,value,label){
    const next=clean(data);
    if(kind==='colors'){value=color(value);if(!value) throw new Error('カラーコードは #2155CD のような6桁で入力してください。');}
    else if(kind==='fonts'){value=String(value||'').trim();if(!value||value.length>120) throw new Error('フォント名を120文字以内で入力してください。');}
    else throw new Error('登録種別が不正です。');
    if(next[kind].some(x=>x.value.toLowerCase()===value.toLowerCase())) return next;
    next[kind].push({value,label:String(label||value).trim().slice(0,60)||value});
    if(bytes(next)>limit) throw new Error('登録データが64 KBを超えるため保存を停止しました。不要な登録を削除してください。');
    return next;
  }
  const api={key,limit,warning,color,clean,bytes,add};
  if(typeof module!=='undefined'&&module.exports) module.exports=api;
  root.FileStyleCore=api;
})(typeof window!=='undefined'?window:globalThis);
