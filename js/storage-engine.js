window.StorageEngine = (() => {
  const KEY="stok_opname_history_v1";
  function get(){try{return JSON.parse(localStorage.getItem(KEY)||"[]")}catch{return[]}}
  function add(item){const data=get();data.unshift(item);localStorage.setItem(KEY,JSON.stringify(data.slice(0,50)))}
  function clear(){localStorage.removeItem(KEY)}
  return {get,add,clear};
})();