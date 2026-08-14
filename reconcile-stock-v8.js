/* Reconcile Stock Pro V8: lightweight SO parsing + manual overstock calculator */
window.ReconcileStockV8 = (() => {
  const factors = [1.5,1.6,1.7,1.8,1.9,2.0];
  function toNumber(v){
    if(typeof v==='number') return Number.isFinite(v)?v:0;
    if(v==null) return 0;
    let s=String(v).trim().replace(/\s/g,'');
    if(!s) return 0;
    if(s.includes('.')&&s.includes(',')) s=s.replace(/\./g,'').replace(',','.');
    else if(/^\d{1,3}(\.\d{3})+$/.test(s)) s=s.replace(/\./g,'');
    else s=s.replace(',','.');
    const n=Number(s); return Number.isFinite(n)?n:0;
  }
  function calculate(stockMonth1,stockMonth2,factor,targetNow){
    const a=toNumber(stockMonth1), b=toNumber(stockMonth2);
    const f=factors.includes(Number(factor))?Number(factor):1.5;
    const target=toNumber(targetNow), combined=a+b, potential=combined*f;
    return {stockMonth1:a,stockMonth2:b,factor:f,combined,potential,targetNow:target,
      difference:potential-target,status:potential>target?'OVERSTOK':'NORMAL'};
  }
  async function readSOFile(file,XLSX){
    const buf=await file.arrayBuffer();
    const wb=XLSX.read(buf,{type:'array',cellFormula:false,cellNF:false,cellStyles:false,cellHTML:false,dense:true});
    const result=[];
    for(const name of wb.SheetNames||[]){
      const ws=wb.Sheets[name]; if(!ws) continue;
      result.push({sheetName:name,rows:XLSX.utils.sheet_to_json(ws,{header:1,defval:''})});
    }
    return result;
  }
  async function processSequential(files,XLSX,handler,onProgress){
    const list=Array.from(files||[]), output=[];
    for(let i=0;i<list.length;i++){
      const f=list[i]; if(onProgress) onProgress(i+1,list.length,f.name);
      const parsed=await readSOFile(f,XLSX);
      output.push(await handler(parsed,f,i));
      await new Promise(r=>setTimeout(r,0));
    }
    return output;
  }
  function factorOptions(selected=1.5){return factors.map(f=>`<option value="${f}" ${Number(f)===Number(selected)?'selected':''}>${f.toFixed(1)}×</option>`).join('');}
  return {factors,toNumber,calculate,readSOFile,processSequential,factorOptions};
})();
