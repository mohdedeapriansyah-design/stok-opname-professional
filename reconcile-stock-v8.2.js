/* Reconcile Stock Pro V8.2 - Android memory optimized */
window.ReconcileStockV82 = (() => {
  const FACTORS=[1.5,1.6,1.7,1.8,1.9,2.0];
  const norm=v=>String(v??'').trim().toLowerCase().replace(/[\s_-]+/g,'');
  const num=v=>{
    if(typeof v==='number') return Number.isFinite(v)?v:0;
    let s=String(v??'').trim().replace(/\s/g,''); if(!s)return 0;
    if(s.includes('.')&&s.includes(','))s=s.replace(/\./g,'').replace(',','.');
    else if(/^\d{1,3}(\.\d{3})+$/.test(s))s=s.replace(/\./g,'');
    else s=s.replace(',','.');
    const n=Number(s); return Number.isFinite(n)?n:0;
  };
  function findHeaderRow(ws,XLSX){
    const ref=XLSX.utils.decode_range(ws['!ref']||'A1:A1');
    for(let r=0;r<=Math.min(ref.e.r,30);r++){
      let sap=-1,qty=-1;
      for(let c=ref.s.c;c<=Math.min(ref.e.c,80);c++){
        const h=norm(ws[XLSX.utils.encode_cell({r,c})]?.v);
        if(h.includes('productsapcode')||h.includes('kodesap'))sap=c;
        if(h.includes('qtyfix'))qty=c;
      }
      if(sap>=0&&qty>=0)return {row:r,sapCol:sap,qtyCol:qty};
    }
    return null;
  }
  async function readSO(file,XLSX,onProgress){
    const wb=XLSX.read(await file.arrayBuffer(),{
      type:'array',dense:true,cellFormula:false,cellNF:false,cellStyles:false,cellHTML:false
    });
    const result=[];
    for(const name of wb.SheetNames||[]){
      const ws=wb.Sheets[name],hdr=findHeaderRow(ws,XLSX); if(!hdr)continue;
      const ref=XLSX.utils.decode_range(ws['!ref']||'A1:A1'),map=new Map();
      for(let r=hdr.row+1;r<=ref.e.r;r++){
        const sap=ws[XLSX.utils.encode_cell({r,c:hdr.sapCol})]?.v;
        if(sap===undefined||sap===null||String(sap).trim()==='')continue;
        const qty=ws[XLSX.utils.encode_cell({r,c:hdr.qtyCol})]?.v;
        map.set(String(sap).trim(),num(qty));
        if((r-hdr.row)%1500===0)await new Promise(requestAnimationFrame);
      }
      result.push({sheet:name,data:map});
    }
    return result;
  }
  async function processSOFiles(files,XLSX,handler,onProgress){
    const list=Array.from(files||[]),all=[];
    for(let i=0;i<list.length;i++){
      onProgress?.({current:i+1,total:list.length,name:list[i].name});
      const parsed=await readSO(list[i],XLSX);
      all.push(await handler(parsed,list[i],i));
      await new Promise(requestAnimationFrame);
    }
    return all;
  }
  function overstock(a,b,f,target){
    const x=num(a),y=num(b),factor=FACTORS.includes(Number(f))?Number(f):1.5,t=num(target);
    const result=(x+y)*factor;
    return {a:x,b:y,factor,total:x+y,result,target:t,difference:result-t,status:result>t?'OVERSTOK':'NORMAL'};
  }
  return {FACTORS,num,findHeaderRow,readSO,processSOFiles,overstock};
})();
