import * as xlsx from 'xlsx';
let process: NodeJS.Process = require('process');
import * as fs from 'fs';
import ExportSheet, { ExportSheetData } from './exportSheet';
import { ExportSheetConfig } from './exportTool';

export class ExportExcelData {
    excel: string;//excel表路径
    workBook: xlsx.WorkBook;
    sheetDatas: ExportSheetData[];

    private constructor() {

    }

    public static async parseExcelData(excelPath: string, sheetConfigs: ExportSheetConfig[]): Promise<ExportExcelData> {
        var realPath = process.cwd() + excelPath;
        let isExist = await fs.existsSync(realPath);
        if (!isExist) {
            console.error(excelPath + " 文件不存在!");
            return;
        }

        let workBook = xlsx.readFile(realPath);
        let excelData = new ExportExcelData;
        excelData.excel = excelPath;
        excelData.workBook = workBook;
        excelData.sheetDatas = [];
        for (let i = 0; i < sheetConfigs.length; i++) {
            const sheetConf = sheetConfigs[i];
            let sheetData = new ExportSheetData;
            sheetData.excel = excelPath;
            sheetData.sheetName = sheetConf.sheetName;
            sheetData.type = sheetConf.type;
            sheetData.key = sheetConf.key;
            sheetData.workSheet = workBook.Sheets[sheetConf.sheetName];

            excelData.sheetDatas.push(sheetData);
        }
        return excelData;
    }
}

export default class ExportExcel {
    private _exportData: ExportExcelData;

    public constructor(exportData: ExportExcelData) {
        this._exportData = exportData;
    }

    public parseExcel() {
        let sheets = this._exportData.sheetDatas;
        for (let i = 0; i < sheets.length; i++) {
            const sheet = sheets[i];
            console.log(`------开始导出表:${sheet.sheetName}------`);
            let exportSheet = new ExportSheet(sheet);
            exportSheet.parseSheet();
            console.log(`------结束导出表:${sheet.sheetName}------`);
        }
    }
}