window.FileEngine = (() => {
  function normalizeHeader(v){return String(v ?? "").trim().toLowerCase().replace(/[\s_\-]+/g,"");}
  function findHeaders(rows){
    const keys = {
      sap: ["productsapcode","sapcode","kodesap","kodeproduk","productcode"],
      qty: ["qtyfix","qty","quantity","quantityfix","nominalso"]
    };
    for(let r=0;r<Math.min(rows.length,30);r++){
      const row=rows[r]||[];
      let sap=-1,qty=-1;
      row.forEach((v,c)=>{
        const h=normalizeHeader(v);
        if(sap<0 && keys.sap.includes(h)) sap=c;
        if(qty<0 && keys.qty.includes(h)) qty=c;
      });
      if(sap>=0 && qty>=0) return {headerRow:r,sapCol:sap,qtyCol:qty};
    }
    return null;
  }
  async function readWorkbook(file){
    const buffer=await file.arrayBuffer();
    return XLSX.read(buffer,{type:"array",cellFormula:true,cellNF:true,cellStyles:true});
  }
  async function readMultiple(files,onProgress){
    const list=Array.from(files||[]);
    const out=[];
    for(let i=0;i<list.length;i++){
      const wb=await readWorkbook(list[i]);
      out.push({file:list[i],workbook:wb});
      if(onProgress) onProgress(Math.round(((i+1)/list.length)*100),`Membaca ${list[i].name}`);
    }
    return out;
  }
  function cleanQty(value){
    if(value===null||value===undefined||value==="") return 0;
    const n=Number(String(value).replace(/,/g,"."));
    if(!Number.isFinite(n)) return 0;
    return Number.isInteger(n)?n:Number(n.toFixed(6));
  }
  function cleanSap(value){return String(value??"").trim();}
  return {findHeaders,readWorkbook,readMultiple,cleanQty,cleanSap};
})();