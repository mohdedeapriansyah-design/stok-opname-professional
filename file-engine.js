window.FileEngine = (() => {
  function normalizeHeader(v){
    return String(v ?? "").trim().toLowerCase().replace(/[\s_\-]+/g,"");
  }

  function findHeaders(rows){
    const keys = {
      sap: ["productsapcode","sapcode","kodesap","kodeproduk","kodeitem","productcode"],
      qty: ["qtyfix","qtyfixed","qty","quantity","quantityfix","nominalso"]
    };
    for(let r=0;r<Math.min(rows.length,30);r++){
      const row=rows[r]||[];
      let sap=-1,qty=-1;
      for(let c=0;c<row.length;c++){
        const h=normalizeHeader(row[c]);
        if(sap<0 && keys.sap.includes(h)) sap=c;
        if(qty<0 && keys.qty.includes(h)) qty=c;
        if(sap>=0 && qty>=0) break;
      }
      if(sap>=0 && qty>=0) return {headerRow:r,sapCol:sap,qtyCol:qty};
    }
    return null;
  }

  async function readWorkbook(file, options={}){
    const buffer=await file.arrayBuffer();
    const isTemplate=options.template===true;
    try {
      return XLSX.read(buffer,{
        type:"array",
        cellFormula:isTemplate,
        cellNF:isTemplate,
        cellStyles:isTemplate,
        cellHTML:false,
        dense:true
      });
    } finally {
      // ArrayBuffer is local; releasing the reference here allows GC after return.
    }
  }

  async function readAndProcess(files, processEntry, onProgress){
    const list=Array.from(files||[]);
    const acc = processEntry ? processEntry.createAccumulator() : null;
    if (!processEntry || !acc) throw new Error("Processor file tidak tersedia.");

    for(let i=0;i<list.length;i++){
      const file=list[i];
      let wb=null;
      try{
        wb=await readWorkbook(file);
        processEntry.process({file,workbook:wb},acc);
      } finally {
        // Important on Android: remove worksheet references immediately.
        if(wb){
          if(wb.Sheets){
            for(const name of Object.keys(wb.Sheets)) wb.Sheets[name]=null;
            wb.Sheets=null;
          }
          wb.SheetNames=null;
        }
        wb=null;
      }
      if(onProgress){
        onProgress(
          Math.round(((i+1)/list.length)*100),
          `Memproses ${i+1}/${list.length}: ${file.name}`
        );
      }
      // Let rendering/GC happen between large Excel files.
      await new Promise(resolve=>setTimeout(resolve,16));
    }
    return acc;
  }

  function cleanQty(value){
    if(value===null||value===undefined||value==="") return 0;
    const n=Number(String(value).replace(/,/g,"."));
    if(!Number.isFinite(n)) return 0;
    return Number.isInteger(n)?n:Number(n.toFixed(6));
  }
  function cleanSap(value){return String(value??"").trim();}
  return {findHeaders,readWorkbook,readAndProcess,cleanQty,cleanSap};
})();
