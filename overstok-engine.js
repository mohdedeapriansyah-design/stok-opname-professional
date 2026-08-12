window.OverstokEngine = (() => {
  /*
   * Analisa per brand.
   * Rumus:
   * ((Nominal stok bulan 1 + Nominal stok bulan 2) * 2) / Omset bulan 3
   *
   * Hasil ratio dikembalikan dalam persen.
   * Status OVERSTOK jika ratio > target ratio.
   */
  function calculateBrand(stokBulan1, stokBulan2, omsetBulan3, targetPercent=100){
    const stok1 = Number(stokBulan1) || 0;
    const stok2 = Number(stokBulan2) || 0;
    const omset3 = Number(omsetBulan3) || 0;
    const target = Number(targetPercent) || 0;

    if (omset3 <= 0) {
      return {
        stok1, stok2, omset3, totalStok: stok1 + stok2,
        numerator: (stok1 + stok2) * 2,
        ratio: Infinity,
        target,
        selisih: Infinity,
        status: "OVERSTOK"
      };
    }

    const numerator = (stok1 + stok2) * 2;
    const ratio = (numerator / omset3) * 100;
    const selisih = ratio - target;

    return {
      stok1, stok2, omset3,
      totalStok: stok1 + stok2,
      numerator,
      ratio,
      target,
      selisih,
      status: ratio > target ? "OVERSTOK" : "AMAN"
    };
  }

  function calculateRows(rows, targetPercent=100){
    return (rows || []).map(row => ({
      ...row,
      ...calculateBrand(row.stok1,row.stok2,row.omset3,targetPercent)
    }));
  }

  function parseBulk(text){
    return String(text||"")
      .split(/\r?\n/)
      .map(x=>x.trim())
      .filter(Boolean)
      .filter(x=>!/^brand\s*[|;\t]/i.test(x))
      .map(line=>{
        const p=line.split(/\s*\|\s*|\t|;/).map(x=>x.trim());
        return {
          brand:p[0]||"",
          stok1:Number(String(p[1]||"0").replace(/,/g,""))||0,
          stok2:Number(String(p[2]||"0").replace(/,/g,""))||0,
          omset3:Number(String(p[3]||"0").replace(/,/g,""))||0
        };
      });
  }

  return {calculateBrand,calculateRows,parseBulk};
})();
