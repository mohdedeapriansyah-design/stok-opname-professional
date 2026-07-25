window.OverstokEngine = (() => {
  function calculate(juni,juli,target,factor=1.5){
    const total=(Number(juni)||0)+(Number(juli)||0),potensi=total*(Number(factor)||0),selisih=potensi-(Number(target)||0);
    return {total,potensi,target:Number(target)||0,selisih,status:potensi>Number(target||0)?"OVERSTOK":"AMAN"};
  }
  function parseBulk(text){
    return String(text||"").split(/\r?\n/).map(x=>x.trim()).filter(Boolean).filter(x=>!/^kode\s*sap/i.test(x)).map(line=>{
      const p=line.split(/\s*\|\s*|\t|;/).map(x=>x.trim());
      return {sap:p[0]||"",name:p[1]||"",qty:Number(String(p[2]||"0").replace(",","."))||0};
    });
  }
  return {calculate,parseBulk};
})();