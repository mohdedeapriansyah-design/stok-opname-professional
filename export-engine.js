window.ExportEngine = (() => {
  function stamp(){return new Date().toISOString().slice(0,19).replace(/[:T]/g,"-");}
  function downloadXLSX(workbook,name=`Hasil_Reconcile_${stamp()}.xlsx`){XLSX.writeFile(workbook,name,{bookType:"xlsx"});}
  function exportNotFound(data,name=`SKU_Tidak_Ditemukan_${stamp()}.xlsx`){
    const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(data.length?data:[{Keterangan:"Tidak ada SKU yang tidak ditemukan"}]);
    XLSX.utils.book_append_sheet(wb,ws,"Tidak Ditemukan");XLSX.writeFile(wb,name);
  }
  function downloadOriginalTemplate(file){if(file){const a=document.createElement("a");a.href=URL.createObjectURL(file);a.download=file.name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);}}
  return {downloadXLSX,exportNotFound,downloadOriginalTemplate};
})();