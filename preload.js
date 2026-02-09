const { contextBridge } = require('electron');
const XLSX = require('xlsx');

contextBridge.exposeInMainWorld('xlsx', {
  read: (data, options) => XLSX.read(data, options),
  utils: {
    sheetToJson: (sheet, options) => XLSX.utils.sheet_to_json(sheet, options)
  }
});
