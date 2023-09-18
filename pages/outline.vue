<template>
    <div class="outline-page">
        <div class="main">
            <ClientOnly>
                <SpreadSheetOutline @ready="onReady" />
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
import { saveAs, getExportOptions } from '~/utils/utils'
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

function onExport() {
    if (!spread.value) {
        return;
    }

    const fileType = "xlsx";
    let fileName = "export." + fileType;
    const options = getExportOptions();

    options.fileType = GC.Spread.Sheets.FileType.excel;

    spread.value.export(
        (blob) => saveAs(blob, fileName),
        (e) => { console.error(e) },
        options
    );
}
</script>
  
<style lang="scss">
.outline-page {
    display: flex;

    .main {
        flex: 1 0 75%;
    }

    .sidebar {
        flex: 1 0 25%;
        padding: 2rem;
    }
}
</style>
  