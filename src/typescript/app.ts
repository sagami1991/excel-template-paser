import "babel-polyfill";
import * as XlsxPopulate from "xlsx-populate";

namespace ExcelNameSpace {
    export interface ICell {
        value(arg: string | number): void;
        value(): string | number | boolean | Date;
    }
    export interface ISheet {
        cell: (range: string) => ExcelNameSpace.ICell;
        active(): void;
        find(value: string): ICell[] | null;
    }
    export interface IWorkBook {
        sheet: (id: number) => ExcelNameSpace.ISheet;
        outputAsync: () => Promise<string | Uint8Array | ArrayBuffer | Blob>;
        find(value: string): ICell[] | null;
        // find(value: string, replace: string | number): ICell[] | null;
    }
}

interface IParseCellInfo {
    sheetNo: number;
    /** ${key名}に埋め込まれる */
    key: string;
    value: string | number;
}

class ExcelTemplatePaser {
    private templateUrl: string;
    private parseData: IParseCellInfo[];
    constructor(props: { templateUrl: string, paseData: IParseCellInfo[] }) {
        this.templateUrl = props.templateUrl;
        this.parseData = props.paseData;
    }

    private getWorkbook() {
        return new Promise<ExcelNameSpace.IWorkBook>((resolve, reject) => {
            const req = new XMLHttpRequest();
            req.open("GET", this.templateUrl, true);
            req.responseType = "arraybuffer";
            req.onreadystatechange = () => {
                if (req.readyState === 4) {
                    if (req.status === 200) {
                        resolve(XlsxPopulate.fromDataAsync(req.response));
                    } else {
                        reject(`Received a ${req.status} HTTP code`);
                    }
                }
            };
            req.send();
        });
    }

    private async fillForm() {
        const workbook = await this.getWorkbook();
        for (const cellParseData of this.parseData) {
            const sheet = workbook.sheet(cellParseData.sheetNo);
            if (!sheet) {
                continue;
            }
            const cell = sheet.find("${" + cellParseData.key + "}");
            if (cell && cell[0]) {
                cell[0].value(cellParseData.value);
            }
        }
        workbook.sheet(0).active();
        return workbook.outputAsync();
    }

    private async generateBlob() {
        const blob = await this.fillForm().catch(e => console.error(e.message));
        return blob;
    }

    public async download() {
        const blob = await this.generateBlob();
        const filename = `example.xlsx`;

        if (window.navigator && window.navigator.msSaveOrOpenBlob) {
            window.navigator.msSaveOrOpenBlob(blob, filename);
        } else {
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            document.body.appendChild(a);
            a.href = url;
            a.download = filename;
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

        }
    }
}

document.querySelector("button.excelForm-submit-button")!.addEventListener("click", () => {
    const parseData: IParseCellInfo[] = [{
        sheetNo: 0,
        key: "saleA",
        value: parseInt((document.querySelector(".excelForm-item-saleA") as HTMLInputElement).value, 10)
    }, {
        sheetNo: 0,
        key: "saleB",
        value: parseInt((document.querySelector(".excelForm-item-saleB") as HTMLInputElement).value, 10)
    }, {
        sheetNo: 0,
        key: "saleC",
        value: parseInt((document.querySelector(".excelForm-item-saleC") as HTMLInputElement).value, 10)
    }, {
        sheetNo: 0,
        key: "saleD",
        value: parseInt((document.querySelector(".excelForm-item-saleD") as HTMLInputElement).value, 10)
    }, {
        sheetNo: 0,
        key: "saleE",
        value: parseInt((document.querySelector(".excelForm-item-saleE") as HTMLInputElement).value, 10)
    }, {
        sheetNo: 1,
        key: "sheet2ValueA",
        value: (document.querySelector(".excelForm-item-sheet2A") as HTMLInputElement).value
    }, {
        sheetNo: 1,
        key: "sheet2ValueB",
        value: (document.querySelector(".excelForm-item-sheet2B") as HTMLInputElement).value
    },
    ];
    const excel = new ExcelTemplatePaser({
        templateUrl: "template/template_a.xlsx",
        paseData: parseData
    });
    excel.download();
});
