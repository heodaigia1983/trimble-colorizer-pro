const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  openIfcFile: () => ipcRenderer.invoke('dialog:openIfcFile'),
  openXktFile: () => ipcRenderer.invoke('dialog:openXktFile'),
  readSampleIfc: () => ipcRenderer.invoke('readSampleIfc')
});
