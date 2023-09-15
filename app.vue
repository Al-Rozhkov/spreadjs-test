<template>
  <div class="app">
    <div class="main">
      <ClientOnly>
        <SpreadSheet @ready="onReady" />
      </ClientOnly>
    </div>
    <div class="sidebar">
      <button @click="onExport">Export to xslx</button>
    </div>
  </div>
</template>

<script setup>
import "@grapecity/spread-sheets-vue";
import "@grapecity/spread-sheets-io";
import GC from '@grapecity/spread-sheets';
// import "@grapecity/spread-excelio";
// import "@grapecity/spread-sheets-charts";
// import "@grapecity/spread-sheets-shapes";
// import "@grapecity/spread-sheets-slicers";
// import "@grapecity/spread-sheets-pivot-addon";
// import "@grapecity/spread-sheets-tablesheet";

const spread = ref(null);

function onReady(instance) {
  spread.value = instance;
}

function saveAs(blob, name = 'export.xlsx') {
  const data = window.URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = data;
  link.download = name;

  // this is necessary as link.click() does not work on the latest firefox
  link.dispatchEvent(
    new MouseEvent('click', {
      bubbles: true,
      cancelable: true,
      view: window
    })
  );

  setTimeout(() => {
    // For Firefox it is necessary to delay revoking the ObjectURL
    window.URL.revokeObjectURL(data);
    link.remove();
  }, 10);
}

function onExport() {
  if (!spread.value) {
    return;
  }

  let fileType = "xlsx";
  let fileName = "export." + fileType;
  const options = getOptions();

  options.fileType = GC.Spread.Sheets.FileType.excel;
  console.log(options);

  spread.value.export(
    (blob) => saveAs(blob, fileName),
    (e) => { console.error(e) },
    options
  );
}

function getOptions() {
  const optionsConfig = [
    {
      propName: "includeBindingSource",
      type: "boolean",
      default: false,
    },
    {
      propName: "includeStyles",
      type: "boolean",
      default: true,
    },
    {
      propName: "includeFormulas",
      type: "boolean",
      default: true,
    },
    {
      propName: "saveAsView",
      type: "boolean",
      default: false,
    },
    {
      propName: "rowHeadersAsFrozenColumns",
      type: "boolean",
      default: false,
    },
    {
      propName: "columnHeadersAsFrozenRows",
      type: "boolean",
      default: false,
    },
    {
      propName: "includeAutoMergedCells",
      type: "boolean",
      default: false,
    },
    {
      propName: "includeCalcModelCache",
      type: "boolean",
      default: false,
    },
    {
      propName: "includeUnusedNames",
      type: "boolean",
      default: true,
    },
    {
      propName: "includeEmptyRegionCells",
      type: "boolean",
      default: true,
    },
  ];

  const optionsValue = {
    includeBindingSource: false,
    includeStyles: true,
    includeFormulas: true,
    saveAsView: false,
    rowHeadersAsFrozenColumns: false,
    columnHeadersAsFrozenRows: false,
    includeAutoMergedCells: false,
    includeCalcModelCache: false,
    includeUnusedNames: true,
    includeEmptyRegionCells: true,
    encoding: "UTF-8",
    rowDelimiter: "\r\n",
    columnDelimiter: ",",
    sheetIndex: 0,
    row: 0,
    column: 0,
    rowCount: 200,
    columnCount: 20,
  };

  let options = {};
  optionsConfig.forEach((prop) => {
    let v = optionsValue[prop.propName];
    if (prop.type === "number") {
      v = +v;
    }
    options[prop.propName] = v;
  });

  return options;
}
</script>

<style>
.app {
  display: flex;
}

.app .main {
  flex: 1 0 75%;
}

.app .sidebar {
  flex: 1 0 25%;
  padding: 2rem;
}
</style>
