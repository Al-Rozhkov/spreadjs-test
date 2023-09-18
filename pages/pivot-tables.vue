<template>
  <div class="spread-sheets-pivot">
    <gc-spread-sheets class="table" @workbookInitialized="initSpread">
    </gc-spread-sheets>

    <div class="options">
      <button @click="onExport">Export to xlsx</button>
    </div>
  </div>
</template>

<script setup>
import "@grapecity/spread-sheets-shapes";
import "@grapecity/spread-sheets-pivot-addon";
import "@grapecity/spread-sheets-charts";
import GC from "@grapecity/spread-sheets";
import { GcSpreadSheets } from "@grapecity/spread-sheets-vue";
import { ref } from "vue";
import { pivotSales } from "../data/pivotData";

const spreadRef = ref(null);

function initSpread(spread) {
  spreadRef.value = spread;
  spread.suspendPaint();
  spread.setSheetCount(4);
  let sheet1 = spread.getSheet(0);
  let sheet2 = spread.getSheet(1);
  let sheet3 = spread.getSheet(2);
  let sheet4 = spread.getSheet(3);
  let tableName = getDataSource(sheet4, pivotSales);
  initPivotTable(sheet1, tableName);
  addCustom(sheet2, tableName);
  initBlankPivot(sheet3, tableName);
  columnFit(sheet4, 0, 6, 100);
  spread.resumePaint();
}

function onExport() {
  if (!spreadRef.value) {
    return;
  }

  const fileType = "xlsx";
  let fileName = "export." + fileType;
  const options = getExportOptions();

  options.fileType = GC.Spread.Sheets.FileType.excel;

  spreadRef.value.export(
    (blob) => saveAs(blob, fileName),
    (e) => {
      console.error(e);
    },
    options
  );
}

function getDataSource(sheet, tableSource) {
  sheet.name("DataSource");
  sheet.setRowCount(117);
  sheet.setColumnWidth(0, 120);
  sheet.getCell(-1, 0).formatter("YYYY-mm-DD");
  sheet.getRange(-1, 4, 0, 2).formatter("$ #,##0");
  let table = sheet.tables.add("table", 0, 0, 117, 6);
  for (let i = 2; i <= 117; i++) {
    sheet.setFormula(i - 1, 5, "=D" + i + "*E" + i);
  }
  table.style(GC.Spread.Sheets.Tables.TableThemes["none"]);
  sheet.setArray(0, 0, tableSource);
  return table.name();
}

function addCustom(sheet, tableName) {
  sheet.name("Custom PivotTable");
  sheet.setRowCount(1000);
  let pivotTableOptions = {
    bandRows: true,
    bandColumns: true,
  };
  let pivotTable = sheet.pivotTables.add(
    "CustomPivotTable",
    tableName,
    1,
    1,
    GC.Spread.Pivot.PivotTableLayoutType.tabular,
    GC.Spread.Pivot.PivotTableThemes.light8,
    pivotTableOptions
  );
  pivotTable.suspendLayout();
  pivotTable.add(
    "salesperson",
    "Salesperson",
    GC.Spread.Pivot.PivotTableFieldType.rowField
  );
  pivotTable.add("car", "Cars", GC.Spread.Pivot.PivotTableFieldType.rowField);
  pivotTable.add(
    "date",
    "Date",
    GC.Spread.Pivot.PivotTableFieldType.columnField
  );
  pivotTable.add(
    "total",
    "Totals",
    GC.Spread.Pivot.PivotTableFieldType.valueField,
    GC.Pivot.SubtotalType.sum
  );
  let itemList = ["Alan", "John", "Tess"];
  pivotTable.labelFilter("Salesperson", {
    textItem: {
      list: itemList,
      isAll: false,
    },
  });
  pivotTable.sort("Salesperson", {
    sortType: GC.Pivot.SortType.asc,
  });
  let carList = ["Audi", "BMW", "Mercedes"];
  pivotTable.labelFilter("Cars", {
    textItem: {
      list: carList,
      isAll: false,
    },
  });
  pivotTable.sort("Cars", {
    sortType: GC.Pivot.SortType.asc,
  });
  let style = new GC.Spread.Sheets.Style();
  style.formatter = "$ #,##0";
  pivotTable.setStyle(
    {
      dataOnly: true,
    },
    style
  );
  pivotTable.resumeLayout();
  pivotTable.autoFitColumn();
}

function initBlankPivot(sheet, source) {
  sheet.name("Blank PivotTable");
  sheet.setColumnWidth(0, 20);
  sheet.pivotTables.add(
    "BlankPivotTable",
    source,
    1,
    1,
    GC.Spread.Pivot.PivotTableLayoutType.outline,
    GC.Spread.Pivot.PivotTableThemes.medium2
  );
}

function initPivotTable(sheet, tableName) {
  sheet.name("Basic PivotTable");
  sheet.setRowCount(1000);
  let pivotTableOptions = {
    bandRows: true,
    bandColumns: true,
  };
  let pivotTable = sheet.pivotTables.add(
    "PivotTable",
    tableName,
    1,
    1,
    GC.Spread.Pivot.PivotTableLayoutType.outline,
    GC.Spread.Pivot.PivotTableThemes.medium1,
    pivotTableOptions
  );
  pivotTable.suspendLayout();
  pivotTable.add(
    "salesperson",
    "Salesperson",
    GC.Spread.Pivot.PivotTableFieldType.rowField
  );
  pivotTable.add("car", "Cars", GC.Spread.Pivot.PivotTableFieldType.rowField);
  pivotTable.add(
    "date",
    "Date",
    GC.Spread.Pivot.PivotTableFieldType.columnField
  );
  let groupInfo = {
    originFieldName: "date",
    dateGroups: [
      {
        by: GC.Pivot.DateGroupType.quarters,
      },
    ],
  };
  pivotTable.group(groupInfo);
  pivotTable.add(
    "total",
    "Totals",
    GC.Spread.Pivot.PivotTableFieldType.valueField,
    GC.Pivot.SubtotalType.sum
  );
  let style = new GC.Spread.Sheets.Style();
  style.formatter = "$ #,##0";
  pivotTable.setStyle(
    {
      dataOnly: true,
    },
    style
  );
  pivotTable.resumeLayout();
  pivotTable.autoFitColumn();
}

function columnFit(sheet, start, end, width) {
  for (let i = start; i < end; i++) {
    sheet.setColumnWidth(i, width);
  }
}
</script>

<style lang="scss">
@import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";

.spread-sheets-pivot {
  display: flex;
  gap: 2rem;
  min-height: 600px;

  .table {
    flex: 1 0 65%;
  }

  .options {
    flex: 1 0 25%;
  }
}
</style>
