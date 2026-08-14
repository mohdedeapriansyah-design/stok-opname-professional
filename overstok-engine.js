window.OverstokEngine = (() => {
  function toNumber(v){
    if(v===null||v===undefined||v==="") return 0;
    const n=Number(String(v).replace(/\./g, '').replace(/,/g, '.'));
    return Number.isFinite(n)?n:0;
  }

  // Rumus: (Stok Bulan 1 + Stok Bulan 2) x Faktor, lalu dibandingkan dengan Target Sekarang.
  function calculate(bulan1,bulan2,target,factor=1.5){
    const a=toNumber(bulan1), b=toNumber(bulan2), t=toNumber(target);
    const f=Math.min(2,Math.max(1.5,toNumber(factor)||1.5));
    const total=a+b;
    const potensi=total*f;
    const selisih=potensi-t;
    return {bulan1:a,bulan2:b,total,potensi,target:t,factor:f,selisih,status:potensi>t?"OVERSTOK":"AMAN"};
  }

  function parseBulk(text){
    return String(text||"").split(/\r?\n/).map(x=>x.trim()).filter(Boolean).filter(x=>!/^kode\s*sap/i.test(x)).map(line=>{
      const p=line.split(/\s*\|\s*|\t|;/).map(x=>x.trim());
      return {sap:p[0]||"",name:p[1]||"",qty:toNumber(p[2])};
    });
  }
  return {calculate,parseBulk};
})();
