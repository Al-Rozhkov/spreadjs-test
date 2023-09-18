export function saveAs(blob: Blob, name = 'export.xlsx') {
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

export function getExportOptions() {
    type OptionsConfigItem = {
        propName: string;
        type: string;
        default: boolean;
    }

    const optionsConfig: OptionsConfigItem[] = [
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

    const optionsValue: any = {
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

    let options: any = {};
    optionsConfig.forEach((prop: OptionsConfigItem) => {
        let v = optionsValue[prop.propName];
        if (prop.type === "number") {
            v = +v;
        }
        options[prop.propName] = v;
    });

    return options;
}
