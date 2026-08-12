window.ExportEngine = (() => {
  function stamp(){
    const d=new Date();
    return d.toISOString().slice(0,10);
  }

  // Non-blocking download: build XLSX as ArrayBuffer, then trigger Blob download.
  async function downloadXLSX(workbook,name=`Hasil_Reconcile_${stamp()}.xlsx`){
    if(!workbook) throw new Error("Workbook hasil tidak tersedia.");
    await new Promise(requestAnimationFrame);
    const data = XLSX.write(workbook,{
      bookType:"xlsx",
      type:"array",
      compression:true
    });
    const blob = new Blob([data],{
      type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;
    a.download=name;
    a.style.display="none";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),2000);
  }

  function exportNotFound(rows,name=`SKU_Tidak_Ditemukan_${stamp()}.xlsx`){
    const wb=XLSX.utils.book_new();
    const ws=XLSX.utils.json_to_sheet(rows||[]);
    XLSX.utils.book_append_sheet(wb,ws,"Tidak Ditemukan");
    return downloadXLSX(wb,name);
  }

  function downloadOriginalTemplate(file){
    if(!file) return;
    const url=URL.createObjectURL(file);
    const a=document.createElement("a");
    a.href=url;
    a.download=file.name;
    a.style.display="none";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
  }

  return {downloadXLSX,exportNotFound,downloadOriginalTemplate};
})();
