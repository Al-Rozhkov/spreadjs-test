<template>
    <div class="spread-sheets-charts">
        <gc-spread-sheets class="table" @workbookInitialized="initSpread">
            <gc-worksheet />
        </gc-spread-sheets>

        <div class="options">
            <button @click="onExport">Export to xslx</button>
        </div>
    </div>
</template>

<script setup>
import GC from "@grapecity/spread-sheets";
import { ref } from "vue";
import {
    GcSpreadSheets,
    GcWorksheet,
} from "@grapecity/spread-sheets-vue";
import "@grapecity/spread-sheets-charts";

const chartTypes = [
    [{
        typeDesc: "Clustered Column",
        type: GC.Spread.Sheets.Charts.ChartType.columnClustered
    },
    {
        typeDesc: "Stacked Column",
        type: GC.Spread.Sheets.Charts.ChartType.columnStacked
    },
    {
        typeDesc: "100% Stacked Column",
        type: GC.Spread.Sheets.Charts.ChartType.columnStacked100
    }
    ],
    [{
        typeDesc: "Line",
        type: GC.Spread.Sheets.Charts.ChartType.line
    },
    {
        typeDesc: "Stacked Line",
        type: GC.Spread.Sheets.Charts.ChartType.lineStacked
    },
    {
        typeDesc: "100% Stacked Line",
        type: GC.Spread.Sheets.Charts.ChartType.lineStacked100
    },
    {
        typeDesc: "Line With Markers",
        type: GC.Spread.Sheets.Charts.ChartType.lineMarkers
    },
    {
        typeDesc: "Stacked Line With Markers",
        type: GC.Spread.Sheets.Charts.ChartType.lineMarkersStacked
    },
    {
        typeDesc: "100% Stacked Line With Markers",
        type: GC.Spread.Sheets.Charts.ChartType.lineMarkersStacked100
    }
    ],
    [{
        typeDesc: "Pie",
        type: GC.Spread.Sheets.Charts.ChartType.pie
    },
    {
        typeDesc: "Doughnut",
        type: GC.Spread.Sheets.Charts.ChartType.doughnut
    }
    ],
    [{
        typeDesc: "Clustered Bar",
        type: GC.Spread.Sheets.Charts.ChartType.barClustered
    },
    {
        typeDesc: "Stacked Bar",
        type: GC.Spread.Sheets.Charts.ChartType.barStacked
    },
    {
        typeDesc: "100% Stacked Bar",
        type: GC.Spread.Sheets.Charts.ChartType.barStacked100
    }
    ],
    [{
        typeDesc: "Area",
        type: GC.Spread.Sheets.Charts.ChartType.area
    },
    {
        typeDesc: "Stacked Area",
        type: GC.Spread.Sheets.Charts.ChartType.areaStacked
    },
    {
        typeDesc: "100% Stacked Area",
        type: GC.Spread.Sheets.Charts.ChartType.areaStacked100
    }
    ],
    [{
        typeDesc: "Scatter",
        type: GC.Spread.Sheets.Charts.ChartType.xyScatter
    },
    {
        typeDesc: "Scatter With Smooth Lines And Markers",
        type: GC.Spread.Sheets.Charts.ChartType.xyScatterSmooth
    },
    {
        typeDesc: "Scatter With Smooth Lines",
        type: GC.Spread.Sheets.Charts.ChartType.xyScatterSmoothNoMarkers
    },
    {
        typeDesc: "Scatter With Straight Lines And Markers",
        type: GC.Spread.Sheets.Charts.ChartType.xyScatterLines
    },
    {
        typeDesc: "Scatter With Straight Lines",
        type: GC.Spread.Sheets.Charts.ChartType.xyScatterLinesNoMarkers
    },
    {
        typeDesc: "Bubble",
        type: GC.Spread.Sheets.Charts.ChartType.bubble
    }
    ],
    [{
        typeDesc: "High-Low-Close",
        type: GC.Spread.Sheets.Charts.ChartType.stockHLC
    },
    {
        typeDesc: "Open-High-Low-Close",
        type: GC.Spread.Sheets.Charts.ChartType.stockOHLC
    },
    {
        typeDesc: "Volume-High-Low-Close",
        type: GC.Spread.Sheets.Charts.ChartType.stockVHLC
    },
    {
        typeDesc: "Volume-Open-High-Low-Close",
        type: GC.Spread.Sheets.Charts.ChartType.stockVOHLC
    }
    ]
];

const spreadRef = ref(null);
const groupType = ref("0");
const chartType = ref("0");
const displayBlanksCells = ref("1");
const displayNaNAsBlank = ref(false);
const ignoreHidden = ref(true);

function initSpread(spread) {
    spreadRef.value = spread;
    spread.suspendPaint();
    spread.setSheetCount(2);
    let sheet1 = spread.getSheet(0);
    sheet1.name("Common Chart");
    let sheet2 = spread.getSheet(1);
    sheet2.name("Custom Chart");
    initSheet(sheet1);
    initSheet(sheet2);
    //add chart
    initChart(sheet1);
    initChart(sheet2);
    //custom chart
    customChartStyle(sheet2);
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
        (e) => { console.error(e) },
        options
    );
}

function _getElementById(id) {
    return document.getElementById(id);
}

function initSheet(sheet) {
    sheet.suspendPaint();
    //prepare data for chart
    let dataArray = [
        ["", "Chrome", "FireFox", "IE", "Safari", "Edge", "Opera", "Other"],
        ["2015", 0.5651, 0.1734, 0.1711, 0.427, 0, 0.184, 0.293],
        ["2016", 0.623, 0.1531, 0.1073, 0.464, 0.311, 0.166, 0.225],
        ["2017", 0.636, 0.1304, 0.834, 0.589, 0.443, 0.223, 0.246]
    ];
    sheet.setArray(0, 0, dataArray);
    sheet.resumePaint();
}

function initChart(sheet) {
    //add common chart
    sheet.charts.add(
        "Chart1",
        GC.Spread.Sheets.Charts.ChartType.columnClustered,
        0,
        100,
        800,
        300,
        "A1:H4"
    );
}

function customChartStyle(sheet) {
    let changeChart = sheet.charts.all()[0];
    changeChartStyle(changeChart);
}

function changeChartStyle(chart) {
    //change orientation
    switchOrientation(chart);
    //change legend
    changeChartLegend(chart);
    //change chartArea
    changeChartArea(chart);
    //change chartTitle
    changeChartTitle(chart);
    //change dataLabels
    changeChartDataLabels(chart);
    //change axisTitles
    changeChartAxisTitles(chart);
    //change axesLine
    changeChartAxesLine(chart);
    //change series
    changeSeries(chart);
    //change gridLine
    changeGridLine(chart);
    //change seriesBorder
    changeSeriesBorder(chart);
}

function switchOrientation(chart) {
    chart.switchDataOrientation();
}

function changeChartLegend(chart) {
    let legend = chart.legend();
    legend.visible = true;
    let legendPosition = GC.Spread.Sheets.Charts.LegendPosition;
    legend.position = legendPosition.top;
    chart.legend(legend);
}

function changeChartArea(chart) {
    let chartArea = chart.chartArea();
    chartArea.backColor = "rgba(93,93,93,1)";
    chartArea.color = "rgba(255,255,255,1)";
    chartArea.fontSize = 14;
    chart.chartArea(chartArea);
}

function changeChartTitle(chart) {
    let title = chart.title();
    title.text = "Browser Market Share";
    title.fontSize = 18;
    chart.title(title);
}

function changeChartDataLabels(chart) {
    let dataLabels = chart.dataLabels();
    dataLabels.showValue = true;
    dataLabels.showSeriesName = false;
    dataLabels.showCategoryName = false;
    dataLabels.format = "0.00%";
    let dataLabelPosition = GC.Spread.Sheets.Charts.DataLabelPosition;
    dataLabels.position = dataLabelPosition.outsideEnd;
    chart.dataLabels(dataLabels);
    let series0 = chart.series().get(0);
    series0.dataLabels = {
        showSeriesName: true,
        showCategoryName: true,
        separator: ";",
        position: GC.Spread.Sheets.Charts.DataLabelPosition.Center,
        color: "red",
        backColor: "white",
        borderColor: "blue",
        borderWidth: 2
    };
    chart.series().set(0, series0);

    let series2 = chart.series().get(2);
    series2.dataLabels = {
        showSeriesName: true,
        separator: "/",
        position: GC.Spread.Sheets.Charts.DataLabelPosition.insideEnd,
        color: "yellow",
        backColor: "white",
        borderColor: "green",
        borderWidth: 1
    };
    chart.series().set(2, series2);

    let series4 = chart.series().get(4);
    series4.dataLabels = {
        showCategoryName: true,
        separator: ":",
        position: GC.Spread.Sheets.Charts.DataLabelPosition.above,
        color: "blue",
        backColor: "white",
        borderColor: "red",
        borderWidth: 2.5
    };
    chart.series().set(4, series4);
}

function changeChartAxisTitles(chart) {
    let axes = chart.axes();
    axes.primaryCategory.title.text = "Year";
    axes.primaryCategory.title.fontSize = 14;
    chart.axes(axes);
}

function changeChartAxesLine(chart) {
    let axes = chart.axes();
    axes.primaryValue.format = "0%";
    chart.axes(axes);
}

function changeSeries(chart) {
    let series = chart.series();
    let seriesItem = series.get(6);
    seriesItem.backColor = "#a3cf62";
    series.set(6, seriesItem);
}

function changeGridLine(chart) {
    let axes = chart.axes();
    axes.primaryCategory.majorGridLine.visible = false;
    axes.primaryValue.majorGridLine.visible = false;
    chart.axes(axes);
}

function changeSeriesBorder(chart) {
    let series = chart.series().get();
    for (let i = 0; i < series.length; i++) {
        let seriesItem = series[i];
        seriesItem.border.color = "rgb(255,255,255)";
        seriesItem.border.width = 1;
        chart.series().set(i, seriesItem);
    }
}

function getActiveChart(sheet) {
    let activeChart = null;
    sheet.charts.all().forEach(function (chart) {
        if (chart.isSelected()) {
            activeChart = chart;
        }
    });
    return activeChart;
}

function judgeIsEmptyOneCell(sheet, range) {
    if (range.rowCount === 1 && range.colCount === 1) {
        let cell = sheet.getCell(range.row, range.col);
        if (!cell.text()) {
            return true;
        }
    }
    return false;
}

</script>

<style lang="scss">
@import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";

.spread-sheets-charts {
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
