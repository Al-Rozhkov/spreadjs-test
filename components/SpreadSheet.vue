<template>
    <gc-spread-sheets :hostClass="hostClass" @workbookInitialized="initSpread">
        <gc-worksheet :dataSource="dataTable">
            <gc-column :width="width" :dataField="'preferredFullName'"></gc-column>
            <gc-column :width="width" :dataField="'jobTitleName'"></gc-column>
            <gc-column :width="width" :dataField="'phoneNumber'"></gc-column>
            <gc-column :width="width" :dataField="'region'"></gc-column>
        </gc-worksheet>
    </gc-spread-sheets>
</template>

<script>
import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";
import GC from '@grapecity/spread-sheets';
import {
    GcSpreadSheets,
    GcWorksheet,
    GcColumn,
} from "@grapecity/spread-sheets-vue";
import { data } from '~/data/data'

export default {
    components: {
        GcSpreadSheets,
        GcWorksheet,
        GcColumn,
    },
    data() {
        return {
            hostClass: "spread-host",
            width: 100,
        };
    },
    computed: {
        dataTable() {
            return data;
        },
    },
    methods: {
        initSpread(spread) {
            this.spread = spread;
            this.$emit('ready', spread);

            if (this.dataTable.length > 0) {
                spread.fromJSON(data[0]);
                let sheet = spread.getSheet(0);
                sheet.suspendPaint();

                const direction = GC.Spread.Sheets.Outlines.OutlineDirection.backward;
                sheet.rowOutlines.direction(direction);

                sheet.rowOutlines.group(3, 4);
                sheet.rowOutlines.group(8, 3);
                sheet.rowOutlines.group(12, 2);
                sheet.rowOutlines.group(15, 3);
                sheet.rowOutlines.group(3, 18);

                // sheet.columnOutlines.group(1, 4);
                // sheet.columnOutlines.group(6, 4);
                sheet.resumePaint();
            }
        }
    }
};
</script>

<style scoped>
.spread-host {
    width: 100%;
    height: 600px;
}
</style>
