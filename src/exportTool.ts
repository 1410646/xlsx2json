import { ExportType } from './exportSheet';
import ExportExcel, { ExportExcelData } from './exportExcel';
import IParser, { TsParser } from './IParser';
import * as fs from 'fs';

export default class ExportTool {
    private _configPath: string
    private _excels: ExportConfig;
    public static parser: IParser

    public constructor(configPath: string) {
        this._configPath = configPath;
    }

    public async initConfig() {
        let configPath = this._configPath;
        let isOk = await fs.existsSync(configPath);
        if (!isOk) {
            throw new Error(configPath + " 文件不存在!");
        }
        let exportConfigStr = fs.readFileSync(configPath).toString();
        this._excels = JSON.parse(exportConfigStr);
        if (!this._excels) {
            throw new Error("export.json 配置错误!");
        }
    }

    public async parseAllExcelData() {
        let excels = this._excels;
        let excelPaths = Object.keys(excels);

        for (let i = 0; i < excelPaths.length; i++) {
            const excelPath = excelPaths[i];
            const sheetConfs = this._excels[excelPath];

            console.log(`--------------开始导出${excelPath}--------------`);
            let excelData = await ExportExcelData.parseExcelData(excelPath, sheetConfs);
            let exportExcel: ExportExcel = new ExportExcel(excelData);
            exportExcel.parseExcel();
            console.log(`--------------结束导出${excelPath}--------------`);
        }
    }
}

export interface ExportConfig {
    [excel: string]: ExportSheetConfig[];
}

export interface ExportSheetConfig {
    sheetName: string;
    excel: string;
    type: ExportType;
    key: string;
}

export let xlsxRootPath:string;
export let outJsonRoot:string;
export let outInterfaceRoot:string;
(async function main() {
    // let parser = new TsParseTypePerformance;
    // let data = new ExportExcelData;
    // data.excel = "/res/excel/config.xlsx";
    // data.type = ExportType.CONFIG;
    // let e = new ExportExcel(data, parser);
    // e.loadExcel();

    xlsxRootPath = process.cwd() + "/res/excel/";
    outJsonRoot = process.cwd() + "/out/json/";
    outInterfaceRoot = process.cwd() + "/out/codeInterfaces/";
    ExportTool.parser = new TsParser;
    let exportTool = new ExportTool(process.cwd() + "/res/config/export.json");
    await exportTool.initConfig();
    await exportTool.parseAllExcelData();

    // //尝试导出
    // let json = (await fs.readFileSync(process.cwd() + "/out/json/pet.json")).toString();
    // let data1: { [index: number]: ConfPet } = JSON.parse(json);
})();